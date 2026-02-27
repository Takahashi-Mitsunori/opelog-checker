# app.py
# CMチェックツール（オペログPDF + フォーマット連絡表CSV/Excel）→ マーカー付きPDF生成
#
# 仕様（確定）:
# - マーク対象列: DT / CMNo / 枠コード(ADDR) のみ
# - 色: 一致=青(ユーザー指定可) / 不一致=赤 / 不明=黄
# - 透明度: 半透明（デフォルト0.20）
# - マーカーは「細い」= 行の文字高さに追従（v7方式）
# - 尺は秒換算して合算判定（例: LT60" + PT60" = 120秒）
# - 枠コード: PR/PT/LT/NT 対応。CMの直前の0は塗らない（0CM→CMから）
# - NF: フォーマット連絡表が発行された番組に対し、番組枠情報内に "NF" を表示
#       NF色は常に青（またはユーザー指定色）。黄/赤にはしない。

import io
import re
import csv
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
import unicodedata

import chardet
import pandas as pd
import pdfplumber
import streamlit as st

# === S2 Experimental ===
# S2 機能: 検出ログのみ (stage1) / フル描画 (full)
ENABLE_S2_STAGE1_LOG_ONLY = False  # ログのみ無効
ENABLE_S2_FULL_IMPLEMENTATION = True  # フル実装を有効化

def _norm_token(t: str) -> str:
    """Normalize token text for robust matching (e.g., full-width, trailing spaces)."""
    try:
        t2 = unicodedata.normalize('NFKC', t)
    except Exception:
        t2 = t
    return (t2 or '').strip().upper()


def _extract_teikyou_key(name: str) -> str:
    """フォーマット連絡表またはオペログPDFの提供名からキー（アルファベット+プライム）を抽出。

    プライム記号の正規化:
        U+00B4 ACUTE ACCENT (´) → NKFCで U+0020+U+0301 に分解 → "B ́" → "B'"
        U+0027 APOSTROPHE (') → そのまま
        U+2019 RIGHT SINGLE QUOTATION MARK → "'"
        U+2032 PRIME (′) → "'"
    PDFスペース除去:
        "提 供 A'" → "提供A'" → キー "A'"
    例:
        '提供A'                  -> 'A'
        '提供B\''                -> "B'"
        '提供Ｂ´' (U+00B4)       -> "B'"
        'サイドE\''              -> "E'"
        'クッション CX=サイドB\'' -> "B'"
        '提 供 B\'' (PDFスペース) -> "B'"
        'CX=提供C'               -> 'C'
    """
    if not name:
        return ""
    n = unicodedata.normalize('NFKC', name)
    # U+0301 (COMBINING ACUTE ACCENT) → アポストロフィ '
    # NFKC("Ｂ´") = "B " + U+0301  → "B'"
    n = re.sub(r'\s*\u0301', "'", n)
    # PDFでスペースが挿入される「提 供」「サ イ ド」を正規化
    n = re.sub(r'提\s*供', '提供', n)
    n = re.sub(r'サ\s*イ\s*ド', 'サイド', n)
    # CX= (ＣＸ= 等) を除去
    n2 = re.sub(r'C[Xx]X?[=＝]', '', n, flags=re.IGNORECASE)
    # プライム系文字を統一 → '
    n2 = re.sub(r"[\u2019\u02bc\u2032\u00b4´′]", "'", n2)
    # 「提供」「サイド」直後のアルファベット+プライムを優先
    m = re.search(r'(?:提供|サイド)([A-Z][\']*)', n2)
    if m:
        return m.group(1)
    # 単独アルファベット+プライム
    m = re.search(r"[A-Z][']*", n2)
    return m.group(0) if m else ""



S2_DEBUG_LOG: list[dict] = []

def _extract_words_fallback(page):
    """
    pdfplumberのextract_wordsは細い列（TR列）で取りこぼすことがあるため、
    まず通常設定で取得し、S2が見当たらない場合のみ許容幅を小さくした再抽出を試みる。
    """
    w1 = page.extract_words(use_text_flow=True, keep_blank_chars=False)
    if any(_norm_token(w.get("text","")) == "S2" for w in w1):
        return w1, "primary"
    # fallback: 小さめ tolerance + blank chars 許容（TR列の細文字対策）
    try:
        w2 = page.extract_words(
            use_text_flow=True,
            keep_blank_chars=True,
            x_tolerance=1,
            y_tolerance=1,
        )
        if len(w2) >= len(w1) and (any(_norm_token(w.get("text","")) in ("S2","S","2") for w in w2)):
            return w2, "fallback"
    except Exception:
        pass
    return w1, "primary"

def _find_s2_spans(words, tr_x0: float, tr_x1: float):
    """
    TR列内の 'S2' を検出する。pdfplumberが 'S' と '2' に分割するケースも考慮し、
    近接している 'S'+'2' を合成して S2 とみなす。
    戻り値: list[dict] (x0,x1,top,bottom,text,method)
    """
    # TR列に入っている単語だけ
    tr_words = [w for w in words if (w.get("x0", 0) >= tr_x0 and w.get("x1", 0) <= tr_x1)]
    tr_words = sorted(tr_words, key=lambda x: (x.get("top", 0), x.get("x0", 0)))

    spans = []
    used = set()

    def _y_overlap(a, b):
        top = max(a["top"], b["top"])
        bottom = min(a["bottom"], b["bottom"])
        return max(0.0, bottom - top)

    # 1) まずは完全一致 'S2'
    for i, w in enumerate(tr_words):
        t = _norm_token(w.get("text", ""))
        if t == "S2":
            spans.append({**w, "text": "S2", "method": "token"})
            used.add(i)

    # 2) 'S' + '2' の合成（同一行近接）
    for i, w in enumerate(tr_words):
        if i in used:
            continue
        t = _norm_token(w.get("text", ""))
        if t != "S":
            continue
        # 近い右隣の '2' を探す
        for j in range(i+1, min(i+4, len(tr_words))):
            if j in used:
                continue
            w2 = tr_words[j]
            t2 = _norm_token(w2.get("text", ""))
            if t2 != "2":
                continue
            # 横方向の距離が近く、縦方向の重なりが十分
            if (w2["x0"] - w["x1"]) <= 6 and _y_overlap(w, w2) >= (0.5 * min(w["bottom"]-w["top"], w2["bottom"]-w2["top"])):
                spans.append({
                    "x0": min(w["x0"], w2["x0"]),
                    "x1": max(w["x1"], w2["x1"]),
                    "top": min(w["top"], w2["top"]),
                    "bottom": max(w["bottom"], w2["bottom"]),
                    "text": "S2",
                    "method": "merge(S+2)",
                })
                used.add(i); used.add(j)
                break

    return spans


from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont


# ----------------------------
# Utilities
# ----------------------------

def norm_tr_char(s: str) -> str | None:
    """Normalize TR cell text to 'R','U','D' or None.
    Handles full-width and minor OCR-ish variants."""
    if s is None:
        return None
    t = s.strip()
    t = t.upper()
    if not t:
        return None
    # full-width -> half-width
    t = t.translate(str.maketrans({
        "Ｒ": "R", "Ｕ": "U", "Ｄ": "D",
        "ｒ": "R", "ｕ": "U", "ｄ": "D",
    }))
    # common confusions
    t = t.replace("∪", "U").replace("∨", "U")
    # take leading char if like 'R.' etc
    m = re.match(r"^[RUD]", t)
    return m.group(0) if m else None

def parse_hhmmss(s: str) -> Optional[int]:
    m = re.search(r"(\d{1,2}):(\d{2}):(\d{2})", s or "")
    if not m:
        return None
    h, mi, se = map(int, m.groups())
    return h * 3600 + mi * 60 + se


def parse_duration_to_seconds(s: str) -> Optional[int]:
    if not s:
        return None
    t = (s.strip()
         .replace("’", "'").replace("′", "'")
         .replace("”", '"').replace("″", '"'))
    m = re.search(r"(\d+)\s*'\s*(\d+)\s*\"", t)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    m = re.search(r"(\d+)\s*\"", t)
    if m:
        return int(m.group(1))
    return None


def decode_csv_bytes(raw: bytes) -> str:
    enc = chardet.detect(raw).get("encoding") or "shift_jis"
    return raw.decode(enc, errors="replace")


# ----------------------------
# Format parsing (CSV / Excel)
# ----------------------------
@dataclass
class FormatProgram:
    program: str
    time_window: Tuple[int, int]          # 放送時刻（start,end）秒
    planned: Dict[str, int]               # CMNo -> planned seconds（合算後）
    q_flags: Dict[str, bool]              # CMNo -> True if format row has 'Q'
    teikyou_flags: Dict[str, bool]        # CMNo -> True if format row has '提供'（提供行はS2と対応）
    teikyou_names: Dict[str, str]         # CMNo -> '提供A', '提供B'', etc.（S2行用）
    left_teikyou_names: Dict[str, str]    # CMNo -> '提供C', '提供H', etc.（D行用・左端独立記載の提供名）
    all_teikyou_names: List[str]          # すべての提供名（順番付き、CM番号と関連付けなし含む）
    teikyou_seconds: Dict[str, int]       # CMNo -> 提供の尺（秒）
    s2_flags: Dict[str, bool]            # CMNo -> True if format row has 独立した 'S2' セル
    net_uke_seconds: List[int]           # ネット受け提供の尺（秒）リスト（出現順）
    source_name: str


def parse_format_csv_bytes(filename: str, raw: bytes) -> Optional[FormatProgram]:
    text = decode_csv_bytes(raw)
    rdr = csv.reader(io.StringIO(text))
    program: Optional[str] = None
    window: Optional[Tuple[int, int]] = None
    planned: Dict[str, int] = {}
    q_flags: Dict[str, bool] = {}
    teikyou_flags: Dict[str, bool] = {}
    teikyou_names: Dict[str, str] = {}
    s2_flags: Dict[str, bool] = {}

    def _parse_window_from_row(row: List[str]) -> Optional[Tuple[int, int]]:
        joined = " ".join(row)
        m = re.search(r"(\d{1,2}:\d{2}:\d{2})\s*[-－〜~]\s*(\d{1,2}:\d{2}:\d{2})", joined)
        if not m:
            return None
        s = parse_hhmmss(m.group(1))
        e = parse_hhmmss(m.group(2))
        if s is None or e is None:
            return None
        return (s, e)

    for row in rdr:
        if not row:
            continue
        row = [c.strip() for c in row]

        if row[0] == "番組名" and len(row) > 1 and row[1]:
            program = row[1]

        # CSV側は「放送時刻/編成時間/編成時刻」など揺れうるので、時刻範囲があれば採用
        if any(k in (row[0] or "") for k in ["放送時刻", "編成時間", "編成時刻"]) and len(row) > 1:
            w = _parse_window_from_row(row)
            if w:
                window = w

        # CM lines (full-width "Ｃ　Ｍ" が含まれる行)
        if any("Ｃ　Ｍ" in c for c in row):
            cmcell = next((c for c in row if re.fullmatch(r"\d{3}", c or "")), None)
            if cmcell:
                cmno = f"CM{cmcell}"
                dur_candidates = [
                    c for c in row
                    if re.search(r"\d+\s*[\"″”]", c) or re.search(r"\d+\s*['’′]\s*\d+\s*[\"″”]", c)
                ]
                dur_candidates = [c for c in dur_candidates if not re.search(r"\d{1,2}:\d{2}:\d{2}", c)]
                dur = sorted(dur_candidates, key=len)[0] if dur_candidates else None
                sec = parse_duration_to_seconds(dur) if dur else None
                if sec is not None:
                    planned[cmno] = planned.get(cmno, 0) + sec
                    # Qフラグ（行内に単独の 'Q' があれば True）
                    if any((c or '').strip() == 'Q' for c in row):
                        q_flags[cmno] = True
                    # 提供フラグ（行内に '提供' があれば True）
                    # 提供がある場合、対応するオペログ行のTR列には「S2」がある
                    teikyou_cells = [c for c in row if '提供' in (c or '')]
                    if teikyou_cells:
                        teikyou_flags[cmno] = True
                        # 提供名を抽出
                        # パターン1: 「提供A」のような単純な形式
                        # パターン2: 「CX=提供B」のような形式
                        for cell in teikyou_cells:
                            # セル全体を取得（「CX=提供B」などを含む）
                            if '提供' in cell:
                                teikyou_names[cmno] = cell.strip()
                                break
                    # S2フラグ（行内に独立したセルとして 'S2' があれば True）
                    # フォーマット連絡表に「S2」と記載がある場合、オペログの提供チャンス欄のS2をマーク
                    if any(_norm_token(c) == 'S2' for c in row):
                        s2_flags[cmno] = True

    if not (program and window and planned):
        return None
    return FormatProgram(program=program, time_window=window, planned=planned, q_flags=q_flags, teikyou_flags=teikyou_flags, teikyou_names=teikyou_names, left_teikyou_names={}, all_teikyou_names=[], teikyou_seconds={}, s2_flags=s2_flags, net_uke_seconds=[], source_name=filename)


def parse_format_excel_bytes(filename: str, raw: bytes) -> List[FormatProgram]:
    """
    添付Excel想定のパーサ。
    - 番組名: U3 を優先。空なら「上部(例:1〜8行)×右側(例:T〜AD列)」から最も“番組名っぽい”文字列を拾う
    - 放送時刻: 「放送時刻:」ラベル行から HH:MM:SS-HH:MM:SS を抽出（なければ「編成時刻:」）
    - CM: 行内に 'Ｃ　Ｍ' があり、3桁のCM番号と尺（60″ or 2'00" 等）がある行を拾い秒換算して合算
    """
    out: List[FormatProgram] = []

    bio = io.BytesIO(raw)
    xls = pd.ExcelFile(bio)

    def _cell(df, r, c) -> str:
        try:
            v = df.iat[r, c]
        except Exception:
            return ""
        return str(v).strip() if v is not None else ""

    def _is_jp_text(s: str) -> bool:
        return bool(re.search(r"[ぁ-んァ-ヶ一-龥]", s or ""))

    def _looks_like_program_name(s: str) -> bool:
        s = (s or "").strip()
        if not s:
            return False
        # 時刻/数値だけは除外
        if re.fullmatch(r"[0-9:\-]+", s):
            return False
        # 典型ラベルは除外
        ng = ["放送時刻", "編成時刻", "編成時間", "番組コード", "枠ID", "入力MUX", "CM尺", "映像モード", "枠音声", "サービス区分", "ﾌｫｰﾏｯﾄNo"]
        if any(k in s for k in ng):
            return False
        # 日本語または英字等の文字が含まれ、適度な長さなら番組名候補とみなす
        has_letters = any(ch.isalpha() for ch in s)
        return has_letters and (2 <= len(s) <= 60)

    def _find_program_name(df) -> Optional[str]:
        # 1) U3 優先（A1表記で U3 = row=2, col=20 (0-index)）
        u3 = _cell(df, 2, 20)
        if _looks_like_program_name(u3):
            return u3

        # 2) フォールバック：上部(0〜7行)×右側(T〜AD列)あたりを探索
        best = ""
        for r in range(0, min(8, df.shape[0])):
            for c in range(19, min(30, df.shape[1])):  # T(19)〜AD(29)
                s = _cell(df, r, c)
                if _looks_like_program_name(s) and len(s) > len(best):
                    best = s
        return best or None

    def _find_time_window(df) -> Optional[Tuple[int, int]]:
        labels = ["放送時刻", "編成時刻", "編成時間"]
        for lab in labels:
            for r in range(df.shape[0]):
                row = [str(x).strip() for x in df.iloc[r].tolist()]
                if any(lab in x for x in row):
                    joined = " ".join(row)
                    m = re.search(r"(\d{1,2}:\d{2}:\d{2})\s*[-－〜~]\s*(\d{1,2}:\d{2}:\d{2})", joined)
                    if m:
                        s = parse_hhmmss(m.group(1))
                        e = parse_hhmmss(m.group(2))
                        if s is not None and e is not None:
                            return (s, e)
        return None

    def _extract_planned(df) -> Tuple[Dict[str, int], Dict[str, bool], Dict[str, bool], Dict[str, str], List[str], Dict[str, int], Dict[str, bool], List[int]]:
        planned: Dict[str, int] = {}
        q_flags: Dict[str, bool] = {}
        teikyou_flags: Dict[str, bool] = {}
        teikyou_names: Dict[str, str] = {}
        left_teikyou_names: Dict[str, str] = {}  # 左端独立記載の提供名（D行用）: CMNo -> 提供名
        all_teikyou_names: List[str] = []  # すべての提供名（順番付き）
        teikyou_seconds: Dict[str, int] = {}  # CMNo -> 提供の尺（秒）
        s2_flags: Dict[str, bool] = {}  # CMNo -> True if format row has 独立した 'S2' セル
        net_uke_seconds: List[int] = []  # ネット受け提供の尺（秒）リスト（出現順）

        max_r, max_c = df.shape[0], df.shape[1]

        # 全角→半角（数字・英字）
        fw_digits = str.maketrans("０１２３４５６７８９", "0123456789")
        fw_alpha = str.maketrans("ＣＭ", "CM")  # 最低限

        def norm(s: str) -> str:
            if s is None:
                return ""
            s = str(s).strip()
            return s.translate(fw_digits).translate(fw_alpha)

        def get_cell(r, c) -> str:
            if r < 0 or c < 0 or r >= max_r or c >= max_c:
                return ""
            v = df.iat[r, c]
            return norm(v)

        cm_pat = re.compile(r"(?:CM)0*([0-9]{3})")  # norm後

        def has_q_near_cm(r, c) -> bool:
            # CMセル(r,c)の左上近傍に単独 'Q' がある想定（上4行×左2列＋同列）
            for rr in range(r - 4, r + 1):
                for cc in range(c - 2, c + 1):
                    v = get_cell(rr, cc)
                    if v == "Q":
                        return True
            return False

        def has_s2_near_cm(r, c) -> bool:
            # CMセル(r,c)の近傍に独立したセルとして 'S2' がある想定（上4行×左右10列）
            # フォーマット連絡表に「S2」と記載がある場合、オペログの提供チャンス欄のS2をマーク
            for rr in range(r - 4, r + 1):
                for cc in range(max(0, c - 2), min(max_c, c + 10)):
                    v = get_cell(rr, cc)
                    if _norm_token(v) == 'S2':
                        return True
            return False

        def has_teikyou_near_cm(r, c) -> bool:
            # CMセル(r,c)の近傍に '提供' / 'サイド' / 'クッション' がある想定
            for rr in range(r - 4, r + 1):
                for cc in range(c - 2, min(c + 10, max_c)):
                    v = get_cell(rr, cc)
                    if '提供' in v or 'サイド' in v or 'クッション' in v:
                        return True
            return False

        def get_teikyou_name_near_cm(r, c) -> Optional[str]:
            # CMセル(r,c)の近傍から提供名を取得
            # 対応パターン:
            #   パターン1: 「提供A」「サイドA'」など（アルファベット付き）
            #   パターン2: 「クッション」単独 → 同じ列の直下2行に「ＣＸ＝サイドB」等がある
            #   パターン3: 「クッション ＣＸ＝サイドB」のように同一セルに含まれる場合

            def _get_full_teikyou_name(rr, cc, v) -> Optional[str]:
                """セル(rr,cc)の値vから完全な提供名を返す。
                「クッション」単独の場合は直下2行のCX=...と結合する。"""
                v = v.strip()
                # クッション単独の場合
                if v == 'クッション':
                    for dr in [1, 2, 3]:
                        if rr + dr < max_r:
                            v2 = get_cell(rr + dr, cc).strip()
                            if v2 and ('ＣＸ' in v2 or 'CX' in v2 or '提供' in v2 or 'サイド' in v2):
                                return f"クッション {v2}"
                    # CX=が見つからなくてもクッションとして返す
                    return 'クッション'
                # 提供/サイドを含む（そのまま返す）
                if '提供' in v or 'サイド' in v or 'ＣＸ' in v or 'CX' in v:
                    return v
                return None

            # 同じ行を優先的に探索
            for cc in range(max(0, c - 5), min(max_c, c + 10)):
                v = get_cell(r, cc)
                if '提供' in v or 'サイド' in v or 'クッション' in v:
                    result = _get_full_teikyou_name(r, cc, v)
                    if result:
                        return result

            # 見つからなければ上下数行も探索
            for rr in range(r - 2, r + 3):
                for cc in range(max(0, c - 5), min(max_c, c + 10)):
                    v = get_cell(rr, cc)
                    if '提供' in v or 'サイド' in v or 'クッション' in v:
                        result = _get_full_teikyou_name(rr, cc, v)
                        if result:
                            return result

            # それでも見つからなければ、提供パターンのみ抽出
            for rr in range(r - 2, r + 3):
                for cc in range(max(0, c - 5), min(max_c, c + 10)):
                    v = get_cell(rr, cc)
                    m = re.search(r'提供[A-ZＡ-Ｚ][\'\'′]*', v)
                    if m:
                        return m.group(0)

            return None

        def find_duration_near(r, c) -> Optional[str]:
            # まず「CMの直上（2〜3行上）」で、右隣〜数列にあるケースが多い（あなたの帳票）
            for rr in range(r - 4, r + 1):
                for cc in range(c - 1, min(c + 10, max_c)):
                    cand = get_cell(rr, cc)
                    if re.search(r"\d+\s*[\"″”]", cand) or re.search(r"\d+\s*['’′]\s*\d+\s*[\"″”]", cand):
                        return cand
            # ダメなら周囲を広めに探索
            for rr in range(max(0, r - 6), min(max_r, r + 6)):
                for cc in range(max(0, c - 4), min(max_c, c + 14)):
                    cand = get_cell(rr, cc)
                    if re.search(r"\d+\s*[\"″”]", cand) or re.search(r"\d+\s*['’′]\s*\d+\s*[\"″”]", cand):
                        return cand
            return None

        # CMセル（例: 'ＣＭ００１'）を起点に planned / q_flags を作る
        for r in range(max_r):
            for c in range(max_c):
                v = get_cell(r, c)
                m = cm_pat.search(v)
                if not m:
                    continue
                cmno = f"CM{m.group(1)}"

                dur = find_duration_near(r, c)
                sec = parse_duration_to_seconds(dur) if dur else None
                if sec is not None:
                    planned[cmno] = planned.get(cmno, 0) + sec

                if has_q_near_cm(r, c):
                    q_flags[cmno] = True

                if has_s2_near_cm(r, c):
                    s2_flags[cmno] = True

                if has_teikyou_near_cm(r, c):
                    teikyou_flags[cmno] = True
                    teikyou_name = get_teikyou_name_near_cm(r, c)
                    if teikyou_name:
                        teikyou_names[cmno] = teikyou_name

                        # 提供の尺を抽出
                        # 実データ調査: 提供セルから2〜5行上・+1列（または同列）に秒数がある
                        teikyou_cell_pos = None
                        for rr in range(max(0, r - 6), r + 1):
                            for cc in range(max(0, c - 3), min(max_c, c + 10)):
                                vv = get_cell(rr, cc)
                                if '提供' in vv and len(vv) < 12:
                                    teikyou_cell_pos = (rr, cc)
                                    break
                            if teikyou_cell_pos:
                                break

                        if teikyou_cell_pos:
                            tr, tc = teikyou_cell_pos
                            found_sec = None
                            for search_r in range(max(0, tr - 6), tr):
                                for search_c in range(max(0, tc - 1), min(max_c, tc + 3)):
                                    dur_cell = get_cell(search_r, search_c)
                                    if dur_cell and ':' not in dur_cell:
                                        sec = parse_duration_to_seconds(dur_cell)
                                        if sec is not None:
                                            found_sec = sec
                                            break
                                if found_sec is not None:
                                    break
                            if found_sec is not None:
                                teikyou_seconds[cmno] = found_sec
        # すべての提供名を抽出（CM番号と関連付けなし含む）
        # 行ごとに「提供」「サイド」「クッション」を含むセルを探し、順番に追加
        for r in range(max_r):
            for c in range(max_c):
                v = get_cell(r, c).strip()
                if not v:
                    continue
                # 「クッション」単独 → 直下2行のCX=...と結合
                if v == 'クッション':
                    teikyou_name = 'クッション'
                    for dr in [1, 2, 3]:
                        if r + dr < max_r:
                            v2 = get_cell(r + dr, c).strip()
                            if v2 and ('ＣＸ' in v2 or 'CX' in v2 or '提供' in v2 or 'サイド' in v2):
                                teikyou_name = f"クッション {v2}"
                                break
                    if teikyou_name not in all_teikyou_names:
                        all_teikyou_names.append(teikyou_name)
                # 「提供」「サイド」「CX=...」を含む場合
                elif '提供' in v or 'サイド' in v:
                    # CX=...単独セルは上のクッションと結合済みなので無視
                    if v.startswith('ＣＸ') or v.startswith('CX'):
                        continue
                    teikyou_name = v
                    if teikyou_name not in all_teikyou_names:
                        all_teikyou_names.append(teikyou_name)

        # ネット受け提供の尺を抽出（「ネット受」セルと同じ列の上方向にある尺）
        for r in range(max_r):
            for c in range(max_c):
                v = get_cell(r, c)
                if 'ネット受' in v and v != 'ネット受け':
                    # 同じ列・近傍列の上方向から尺を探す
                    for search_r in range(r - 1, max(0, r - 8), -1):
                        for dc in [-1, 0, 1, 2]:
                            d = get_cell(search_r, c + dc)
                            s = parse_duration_to_seconds(d) if d else None
                            if s is not None:
                                net_uke_seconds.append(s)
                                break
                        else:
                            continue
                        break

        # 左端（col=0〜4）の独立した提供名を抽出（D行用）
        # フォーマット連絡表でD行トリガー提供の提供名は、CM番号列の左端に単独記載される
        # 例: 「提供Ｃ」(col=2) → 同行右の「ＣＭ００６」に紐付け
        for r in range(max_r):
            for c in range(min(5, max_c)):
                v = get_cell(r, c).strip()
                if not v:
                    continue
                if '提供' in v or 'サイド' in v:
                    # 同行の右側からCM番号を探す
                    for cc in range(c + 1, max_c):
                        vv = get_cell(r, cc).strip()
                        m = cm_pat.search(vv)
                        if m:
                            cmno = f"CM{m.group(1).zfill(3)}"
                            if cmno not in left_teikyou_names:
                                left_teikyou_names[cmno] = v
                                print(f"[FORMAT] left_teikyou_names[{cmno}] = {repr(v)}")
                            break

        return planned, q_flags, teikyou_flags, teikyou_names, left_teikyou_names, all_teikyou_names, teikyou_seconds, s2_flags, net_uke_seconds

    for sheet in xls.sheet_names:
        df = xls.parse(sheet_name=sheet, header=None, dtype=str).fillna("")
        program = _find_program_name(df)
        window = _find_time_window(df)
        planned, q_flags, teikyou_flags, teikyou_names, left_teikyou_names, all_teikyou_names, teikyou_seconds, s2_flags, net_uke_seconds = _extract_planned(df)

        if program and window and planned:
            out.append(FormatProgram(
                program=program,
                time_window=window,
                planned=planned,
                q_flags=q_flags,
                teikyou_flags=teikyou_flags,
                teikyou_names=teikyou_names,
                left_teikyou_names=left_teikyou_names,
                all_teikyou_names=all_teikyou_names,
                teikyou_seconds=teikyou_seconds,
                s2_flags=s2_flags,
                net_uke_seconds=net_uke_seconds,
                source_name=f"{filename}#{sheet}"
            ))
    return out


def load_formats(uploaded_files) -> List[FormatProgram]:
    formats: List[FormatProgram] = []
    for uf in uploaded_files:
        name = uf.name
        raw = uf.getvalue()
        if name.lower().endswith(".csv"):
            fp = parse_format_csv_bytes(name, raw)
            if fp:
                formats.append(fp)
        elif name.lower().endswith((".xlsx", ".xls")):
            formats.extend(parse_format_excel_bytes(name, raw))
        else:
            st.warning(f"未対応形式: {name}（CSV / Excel を入力してください）")
    formats = [f for f in formats if f.time_window]
    formats.sort(key=lambda f: f.time_window[0])
    return formats


# ----------------------------
# OpeLog parsing + marker targets
# ----------------------------
def header_word(words, text):
    return next((w for w in words if w["text"] == text), None)


def status_for(planned_map: Dict[str, int], cmno: str, actual: Optional[int]) -> str:
    planned = planned_map.get(cmno)
    if planned is None or actual is None:
        return "unknown"
    return "match" if planned == actual else "mismatch"


def find_header_boxes(page):
    words, _words_mode = _extract_words_fallback(page)
    lab_id = [w for w in words if w["text"] == "枠ID"]
    boxes = []
    for lab in lab_id:
        tr = next((w for w in words if w["text"] == "TR" and w["top"] > lab["top"]), None)
        y_top = lab["top"] - 10
        y_bottom = (tr["top"] + 5) if tr else (lab["top"] + 260)
        region = [w for w in words if y_top <= w["top"] <= y_bottom]

        time_vals = [w for w in region if re.fullmatch(r"\d{1,2}:\d{2}:\d{2}", w["text"])]
        start = end = None
        if len(time_vals) >= 2:
            tv = sorted(time_vals, key=lambda w: w["x0"])
            start = parse_hhmmss(tv[0]["text"])
            end = parse_hhmmss(tv[1]["text"])

        lab_cm = next((w for w in region if w["text"] == "CM"), None)
        lab_vid = next((w for w in region if w["text"].startswith("映像モード")), None)
        if lab_cm and lab_vid:
            x_nf = lab_cm["x1"] + (lab_vid["x0"] - lab_cm["x1"]) * 0.45
            y_nf_top = lab_cm["top"] + 16
        else:
            x_nf, y_nf_top = 250, lab["top"] + 25

        boxes.append({
            "y_top": y_top, "y_bottom": y_bottom,
            "start": start, "end": end,
            "x_nf": x_nf, "top_nf": y_nf_top,
        })
    if not boxes:
        boxes.append({"y_top": 0, "y_bottom": 250, "start": None, "end": None, "x_nf": 250, "top_nf": 90})
    return boxes


def match_format_for_box(box, formats: List[FormatProgram]) -> Optional[FormatProgram]:
    """
    ヘッダ枠（枠IDブロック）の開始/終了時刻と、フォーマット連絡表の開始/終了時刻を突合して
    「この番組はNF（=フォーマット連絡表あり）」かどうかを判断する。

    ★重要: R/U ロジックやマーカー仕様に影響を与えないため、この関数は NF 判定のみに使用する。
    """
    if box.get("start") is None or box.get("end") is None:
        return None

    bs = int(box["start"])
    be = int(box["end"])

    for f in formats:
        fs, fe = f.time_window
        fs = int(fs)
        fe = int(fe)

        # まずは従来通りの厳密一致（±60秒）
        if abs(fs - bs) <= 60 and abs(fe - be) <= 60:
            return f

        # フォールバック（final23運用を崩さないため、条件は厳しめ）
        # - 開始時刻は一致している（±60秒）
        # - ただし、フォーマット側の終了が「かなり長く」取られているケースがある
        #   （例: 24:17-27:50 のように、番組枠より大きい“編成帯”が入っている）
        #   → この場合でも、枠IDブロックの [開始-終了] がフォーマット窓の中に入っていれば NF とみなす。
        if abs(fs - bs) <= 60:
            # 枠の開始がフォーマット開始より前に見えるのはNG（誤マッチ防止）
            if bs < fs - 60:
                continue

            # 枠の終了がフォーマット終了より後ろに大きくはみ出すのはNG
            if be > fe + 60:
                continue

            # ただし「フォーマットの方が番組枠より極端に長い」場合だけ許容する
            fmt_len = fe - fs
            box_len = be - bs
            # 1時間以上の差があり、かつフォーマットが2時間以上の長さを持つ場合に限定
            if fmt_len >= 2 * 3600 and (fmt_len - box_len) >= 3600:
                return f

    return None
    for f in formats:
        s, e = f.time_window
        if abs(s - box["start"]) <= 60 and abs(e - box["end"]) <= 60:
            return f
    return None


def extract_targets_and_nf(opelog_bytes: bytes, formats: List[FormatProgram]):
    targets = []
    nf_labels = []

    with pdfplumber.open(io.BytesIO(opelog_bytes)) as pdf:
        for pi, page in enumerate(pdf.pages):
            words, _words_mode = _extract_words_fallback(page)  # primary/fallback; for S2 debug

            h_tr = header_word(words, "TR")
            h_dt = header_word(words, "DT")
            h_start = header_word(words, "開始時刻")
            h_cmno = header_word(words, "CMNo")
            h_waku = header_word(words, "枠コード")
            h_addr = header_word(words, "ADDR")
            # S2列とスポンサー列のヘッダーを検出（存在しない場合もあるのでオプショナル）
            h_s2 = header_word(words, "S2")
            h_sponsor = next((w for w in words if "スポンサー" in w["text"] or "備考" in w["text"] or "VTR" in w["text"]), None)
            if not (h_tr and h_dt and h_cmno and h_waku and h_addr):
                continue

            tr_x0 = h_tr["x0"] - 40
            tr_x1 = h_dt["x0"] - 10

            # 開始時刻列の範囲（めざましはヘッダ位置がズレることがあるため推定も用意）
            if h_start:
                start_x0 = h_start["x0"] - 10
                start_x1 = h_dt["x0"] - 10
                # TR列は開始時刻の左までに絞る
                tr_x1 = start_x0 - 5
            else:
                # ヘッダが取れない場合：本文の HH:MM:SS の x0 から開始時刻列を推定（中央値）
                time_words = [w for w in words if re.fullmatch(r"\d{2}:\d{2}:\d{2}", (w.get("text") or "").strip())]
                if time_words:
                    xs = sorted(w["x0"] for w in time_words)
                    start_x0 = xs[len(xs)//2] - 5
                else:
                    # 最終フォールバック：TRヘッダの右端の少し右を開始時刻とみなす
                    start_x0 = h_tr["x1"] + 2
                start_x1 = h_dt["x0"] - 10
                tr_x1 = start_x0 - 5

            dt_x0 = h_dt["x0"] - 10
            dt_x1 = h_cmno["x0"] - 10
            waku_x0 = h_waku["x0"] - 160
            waku_x1 = h_addr["x1"] + 80

            # S2列の範囲（ヘッダーがある場合）
            s2_x0 = h_s2["x0"] - 10 if h_s2 else None
            s2_x1 = h_s2["x1"] + 10 if h_s2 else None

            # スポンサー/備考列の範囲（ヘッダーがある場合、右端まで広めに取る）
            sponsor_x0 = h_sponsor["x0"] - 10 if h_sponsor else None
            sponsor_x1 = h_sponsor["x0"] + 200 if h_sponsor else None  # 右方向に広めに取る

            # NF: per header box, match by start/end times
            for box in find_header_boxes(page):
                fmt = match_format_for_box(box, formats)
                if fmt:
                    nf_labels.append({"page": pi, "x": box["x_nf"], "top": box["top_nf"], "program": fmt.program})

            cm_words = [w for w in words if re.fullmatch(r"CM\d{3}", w["text"])]

            # --- 追加: CMNoが無い「トリガー行」（R/Uなど）も拾う ---
            # 例: めざましテレビのように R がCM行の直上に単独で出る、また「とれたてっ！」のU行のように
            # トリガー行自体にはCMNoが無いケースがあるため、ページ全体から「開始時刻の左側にあるR/U/D」を抽出して、
            #  (1) その行のTRセル自体をマーキング対象にする（R/Uのみ）
            #  (2) 同一開始時刻のCM行で「Dのみ」になっている場合に R/U を参照できるよう辞書化する
            triggers_by_time: Dict[int, str] = {}

            # 本文の開始時刻ワード一覧（行の特定用）
            all_time_words = [w for w in words if re.fullmatch(r"\d{2}:\d{2}:\d{2}", (w.get("text") or "").strip())]

            def _nearest_time_for_y(y_center: float, y_top: float, y_bottom: float) -> Optional[dict]:
                # 同一行判定（縦方向が近い開始時刻を探す）
                best_tw = None
                best_dy = 1e9
                for tw in all_time_words:
                    twc = (tw["top"] + tw["bottom"]) / 2.0
                    dy = abs(twc - y_center)
                    if dy < best_dy:
                        best_dy = dy
                        best_tw = tw
                # まずは中心距離で許可（少し広め：行間ズレ対策）
                if best_tw and best_dy <= 22.0:
                    return best_tw
                # 次に、縦方向の重なりで許可（中心がズレても同一行のことがある）
                for tw in all_time_words:
                    if not (tw["bottom"] < y_top - 2 or tw["top"] > y_bottom + 2):
                        return tw
                return None

            for w in words:
                c = norm_tr_char(w.get("text") or "")
                if c is None:
                    continue
                if c == "D":
                    continue
                # TR列っぽい位置（開始時刻の左側にある単独文字）
                y_center = (w["top"] + w["bottom"]) / 2.0
                tw = _nearest_time_for_y(y_center, w["top"], w["bottom"])
                if not tw:
                    continue
                tsec2 = parse_hhmmss(tw["text"])
                if tsec2 is None:
                    continue

                # 同一開始時刻で R/U を優先（Dより強い）
                prev = triggers_by_time.get(tsec2)
                if prev in ("R", "U"):
                    pass
                else:
                    triggers_by_time[tsec2] = c

                # R/U は「トリガー行」自体もマーキング（D単体は意味を持たないので塗らない）
                if c in ("R", "U"):
                    fmt2 = None
                    for f in formats:
                        s, e = f.time_window
                        if s <= tsec2 <= e:
                            fmt2 = f
                            break
                    # programが取れない場合でも描画はできるが、CSV用に空文字を入れる
                    prog2 = fmt2.program if fmt2 else ""

                    y_pad2 = 3
                    row_top2 = tw["top"] - y_pad2
                    row_bottom2 = tw["bottom"] + y_pad2

                    targets.append({
                        "page": pi,
                        "status": "match",          # トリガーは青で塗りたいのでmatch扱い
                        "row_bbox": (row_top2, row_bottom2),
                        "dt_span": None,
                        "cm_span": None,
                        "kind": "tr",
                        "tr_span": (w["x0"] - 2, w["x1"] + 2),
                        "tr_y_span": (w["top"] - 2, w["bottom"] + 2),
                        "start_span_candidate": None,
                        "waku_span": None,
                        "program": prog2,
                        "cmno": "",
                        "planned_sec": None,
                        "actual_sec": None,
                    })

            # TRの直前状態（R/U→D判定用）
            prev_trigger_char = None
            prev_cmno = None
            prev_top = None
            prev_tsec = None
            prev_tsec = None
            for cmw in cm_words:
                y_pad = 3  # v7 thin
                row_top = cmw["top"] - y_pad
                row_bottom = cmw["bottom"] + y_pad
                row_words = [w for w in words if not (w["bottom"] < row_top or w["top"] > row_bottom)]
                row_words_sorted = sorted(row_words, key=lambda x: x["x0"])
                row_text = " ".join(w["text"] for w in row_words_sorted)

                tsec = parse_hhmmss(row_text)
                if tsec is None:
                    continue

                fmt = None
                for f in formats:
                    s, e = f.time_window
                    if s <= tsec <= e:
                        fmt = f
                        break
                if fmt is None:
                    continue

                cmno = cmw["text"]

                # Qフラグ（フォーマット連絡表にQがあるCMのみ、TR列の 'R' を追加マーキング）
                q_flag = bool(getattr(fmt, 'q_flags', {}).get(cmno, False))
                
                # 提供フラグ（フォーマット連絡表に提供があるCMのみ）
                # 提供がある場合、TR列には「S2」があるはず
                teikyou_flag = bool(getattr(fmt, 'teikyou_flags', {}).get(cmno, False))

                # フォーマット連絡表S2フラグ（独立したS2セルがある場合）
                # このフラグがTrueの場合、オペログの提供チャンス欄の「S2」文字列をマークする
                format_s2_flag = bool(getattr(fmt, 's2_flags', {}).get(cmno, False))

                dt_words = [
                    w for w in row_words_sorted
                    if dt_x0 <= w["x0"] <= dt_x1 and (
                        ("'" in w["text"]) or ('"' in w["text"]) or ("″" in w["text"]) or ('”' in w["text"])
                    )
                ]
                dt_str = " ".join(w["text"] for w in dt_words)
                actual = parse_duration_to_seconds(dt_str)

                # statusはCMNo単位（同一CMNoの合算）で後段で確定
                status = None

                dt_span = (min(w["x0"] for w in dt_words) - 1, max(w["x1"] for w in dt_words) + 1) if dt_words else None
                cm_span = (cmw["x0"] - 1, cmw["x1"] + 1)

                tr_span = None

                tr_y_span = None

                tr_char = None  # 'R'/'U'/'D'/None

                # TR列のトリガー文字（R/U/D）を検出
                # めざましのように列幅がズレたり、R/UがCMNo行の直上に出るケースに強い方式：
                #  - その行の「開始時刻(HH:MM:SS)」ワードのx0を境界にして、左側の単独文字をTRとして扱う
                #  - 同一行で見つからなければ、上方向（最大2行分程度）を探索して最も近いものを採用

                # その行の開始時刻ワード bbox を拾う（これが一番安定）
                time_words = [w for w in row_words_sorted if re.fullmatch(r"\d{2}:\d{2}:\d{2}", w["text"].strip())]
                row_start_x0 = time_words[0]["x0"] if time_words else None

                candidates = []

                def _push_if_tr(w):
                    c = norm_tr_char(w["text"])
                    if c is None:
                        return
                    # 開始時刻の左側にあること（TR欄付近）
                    if row_start_x0 is not None and w["x1"] >= row_start_x0 - 2:
                        return
                    candidates.append((c, w))

                # まず同一行（row_words）から拾う
                for w in row_words_sorted:
                    _push_if_tr(w)

                # 見つからなければ上方向を探索（R/U/D がCMNo行の直上に出るケース）
                if not candidates:
                    y_center = (row_top + row_bottom) / 2.0
                    for w in words:
                        c = norm_tr_char(w["text"])
                        if c is None:
                            continue
                        if row_start_x0 is not None and w["x1"] >= row_start_x0 - 2:
                            continue
                        w_center = (w["top"] + w["bottom"]) / 2.0
                        # CMNo行より上〜同じあたり（上方向に最大140pt程度）
                        if -140 <= (w_center - y_center) <= 25:
                            candidates.append((c, w))

                best = None
                if candidates:
                    y_center = (row_top + row_bottom) / 2.0
                    # 最も近い（縦距離）を採用
                    c, wbest = sorted(candidates, key=lambda cw: abs(((cw[1]["top"] + cw[1]["bottom"]) / 2.0) - y_center))[0]
                    tr_char = c
                    best = wbest

# --- マーキング仕様 ---

                # - R/U は必ずマーク（トリガー）

                # - D単体はマークしない

                # - ただし直前が同一CMNoの R/U のときだけ（R/U→D連続）Dもマーク

                if tr_char in ("R", "U") and best is not None:

                    tr_span = (best["x0"] - 2, (row_start_x0 - 2) if row_start_x0 is not None else (best["x1"] + 2))

                    tr_y_span = (best["top"], best["bottom"])

                elif tr_char == "D" and best is not None:
                    # D行が「CM行」の場合は、TRセルにもマークを付けたい。
                    # 同一開始時刻にR/Uが別行で存在する場合でも、D側も併せてマークしてOK。
                    # また、提供フラグがある場合（CM直後の提供）もDをマークする。
                    tr_span = (best["x0"] - 2, (row_start_x0 - 2) if row_start_x0 is not None else (best["x1"] + 2))
                    tr_y_span = (best["top"], best["bottom"])


                tr_really_empty = (tr_char is None)


                # 開始時刻（HH:MM:SS）のbboxは「時刻文字そのもの」から拾う（列ズレ耐性）

                start_span_candidate = None

                for w in row_words_sorted:

                    if w["x0"] >= dt_x0:

                        continue

                    if re.fullmatch(r"\d{2}:\d{2}:\d{2}", w["text"].strip()):

                        start_span_candidate = (w["x0"] - 1, w["x1"] + 1)

                        break


                waku_span = None


                for w in row_words_sorted:
                    if not (waku_x0 <= w["x0"] <= waku_x1):
                        continue
                    m = re.search(r"(CM\d{3}(?:PR|PT|LT|NT)(?:/[A-Z]{2})?)", w["text"])
                    if m:
                        text = w["text"]
                        x0, x1 = w["x0"], w["x1"]
                        cw = (x1 - x0) / max(len(text), 1)
                        sel_x0 = x0 + cw * m.start(1)
                        sel_x1 = x0 + cw * m.end(1)
                        waku_span = (sel_x0 - 1, sel_x1 + 1)
                        break

                # S2列のspan（提供フラグがある場合のみ）
                # 提供がある場合、TR列（開始時刻の左側）に「S2」がある
                # ただし、S2はCM行ではなく別のS2行にあるため、CM行ではs2_spanは設定しない
                s2_span = None
                teikyou_name = None
                teikyou_span = None
                
                # 提供名を取得（S2行で使用するため保存）
                if teikyou_flag:
                    teikyou_name = getattr(fmt, 'teikyou_names', {}).get(cmno)

                
                # 次行のためにTR状態を更新

                effective_tr = triggers_by_time.get(tsec) if triggers_by_time.get(tsec) in ("R","U") else tr_char
                prev_trigger_char = effective_tr
                prev_cmno = cmno
                prev_top = cmw["top"]
                prev_tsec = tsec
                targets.append({
                    "page": pi,
                    "status": status,
                    "row_bbox": (row_top, row_bottom),
                    "dt_span": dt_span,
                    "cm_span": cm_span,
                    "tr_span": tr_span,
                    "start_span_candidate": start_span_candidate,
                    "tr_char": tr_char,
                    "tr_y_span": tr_y_span,
                    "waku_span": waku_span,
                    "s2_span": s2_span,
                    "teikyou_span": teikyou_span,
                    "teikyou_name": teikyou_name,
                    "teikyou_flag": teikyou_flag,
                    "kind": "cm",
                    "program": fmt.program,
                    "cmno": cmno,
                    "planned_sec": fmt.planned.get(cmno),
                    "actual_sec": actual,
                })

            # --- S2行（提供行）とD行（提供）の処理 ---
            # TR列の範囲を常に定義（STAGE1/FULL両方で使用）
            def _sec_to_hms(_t: int) -> str:
                h = _t // 3600
                m = (_t % 3600) // 60
                s = _t % 60
                return f"{h:02d}:{m:02d}:{s:02d}"
            # TR列の範囲（ヘッダ「TR」と「開始時刻」から推定）
            _h_tr = next((w for w in words if _norm_token(w.get("text","")) == "TR"), None)
            _h_start_hdr = next((w for w in words if _norm_token(w.get("text","")) == "開始時刻"), None)
            _tr_x0 = (_h_tr["x0"] - 2) if _h_tr else 0
            _tr_x1 = (_h_start_hdr["x0"] - 2) if (_h_tr and _h_start_hdr and _h_tr.get("x0") < _h_start_hdr.get("x0")) else (_h_tr["x1"] + 40 if _h_tr else 60)

            # --- S2 stage1 (log only; no marking/drawing) ---
            if ENABLE_S2_STAGE1_LOG_ONLY:
                _s2_spans = _find_s2_spans(words, _tr_x0, _tr_x1)
                # Try to locate the Start Time ("開始時刻") column span from the header row,
                # so we can robustly parse time even when row text is noisy.
                _start_time_x0 = None
                _start_time_x1 = None
                _h_start = next((w for w in words if _norm_token(w.get("text","")) == "開始時刻"), None)
                _h_dt = next((w for w in words if _norm_token(w.get("text","")) == "DT"), None)
                if _h_start and _h_dt and _h_start.get("x0") < _h_dt.get("x0"):
                    _start_time_x0 = _h_start["x0"] - 2
                    _start_time_x1 = _h_dt["x0"] - 2

                for _s2w in _s2_spans:
                    _y_pad = 8
                    _row_top = _s2w["top"] - _y_pad
                    _row_bottom = _s2w["bottom"] + _y_pad
                    _row_words = [w for w in words if not (w["bottom"] < _row_top or w["top"] > _row_bottom)]
                    _row_words_sorted = sorted(_row_words, key=lambda x: x["x0"])
                    _row_text = " ".join(w["text"] for w in _row_words_sorted)
                    _tsec = parse_hhmmss(_row_text)
                    if _tsec is None:
                        continue
                    _fmt = None
                    for _f in formats:
                        _s, _e = _f.time_window
                        if _s <= _tsec <= _e:
                            _fmt = _f
                            break
                    S2_DEBUG_LOG.append({
                        "page": pi + 1,
                        "time": _sec_to_hms(_tsec),
                        "program": getattr(_fmt, "program", None),
                        "raw": _row_text[:120],
                        "_words_mode": _words_mode,
                        "_s2_method": _s2w.get("method"),
                    })

            # --- S2 full implementation (disabled in stage1) ---
            if ENABLE_S2_FULL_IMPLEMENTATION:
                print("[DEBUG_S2] BUILD 2026-02-21 A", flush=True)
                # ページレベルで変数を事前初期化（複数ループ間の安全性確保）
                d_span = None
                s2_span = None
                teikyou_span = None
                fmt_d = None  # D行ループ外でも参照されるため前もって初期化
                
                # D行提供の出現順インデックス管理（CM番号との紐付けなし）
                # フォーマット連絡表のall_teikyou_namesの何番目まで使ったか（番組→インデックス）
                # ページをまたがるため、ページの外側（枠レベル）で初期化が必要だが
                # ここではページ内で完結する番組ごとに管理
                # ※ targets追加後に後処理で割り当てるため、ここではD行提供カウントのみ管理
                
                # 同じCM番号で複数の提供行がある場合、最初の1行のみに提供名を描画
                processed_cm_for_teikyou = set()  # 既に提供名を描画したCM番号を記録
                
                # TR列に「S2」がある行を抽出し、提供名を描画またはマーク
                s2_words = _find_s2_spans(words, _tr_x0, _tr_x1)  # normalized + merge
                print(f"[DEBUG_S2] Page {pi+1}: s2_words found: {len(s2_words)}", flush=True)
                
                # スポンサー列の範囲を再定義（S2行処理用）
                # スポンサー列の範囲を制限（右端列の「提」を除外するため）
                h_sponsor_s2 = next((w for w in words if "スポンサー" in w["text"] or "備考" in w["text"] or "VTR" in w["text"]), None)
                sponsor_x0_s2 = h_sponsor_s2["x0"] - 10 if h_sponsor_s2 else 593
                sponsor_x1_s2 = 720  # スポンサー列の右端（右端列の「提」を除外）
                
                for s2w in s2_words:
                    y_pad = 3
                    row_top = s2w["top"] - y_pad
                    row_bottom = s2w["bottom"] + y_pad
                    row_words = [w for w in words if not (w["bottom"] < row_top or w["top"] > row_bottom)]
                    row_words_sorted = sorted(row_words, key=lambda x: x["x0"])
                    
                    # 開始時刻を取得
                    row_text = " ".join(w["text"] for w in row_words_sorted)
                    tsec = parse_hhmmss(row_text)
                    if tsec is None:
                        print(f"[DEBUG_S2] Page {pi+1}: 時刻解析失敗 - row_text='{row_text}'→スキップ", flush=True)
                        continue
                    
                    # フォーマット連絡表と突合（時刻範囲で番組を特定）
                    fmt_s2 = None
                    for f in formats:
                        s, e = f.time_window
                        if s <= tsec <= e:
                            fmt_s2 = f
                            break
                    
                    if fmt_s2 is None:
                        print(f"[DEBUG_S2] Page {pi+1}: 時刻{tsec}名(秒:{tsec})がフォーマット連絡表にない→スキップ", flush=True)
                        continue
                    
                    # このS2行の直前のCM行を探す（同じページ内で、このS2行より上にある最も近いCM行）
                    preceding_cm = None
                    min_distance = float('inf')
                    
                    for t in targets:
                        if t.get("kind") == "cm" and t["page"] == pi and t.get("program") == fmt_s2.program:
                            # CM行の下端とS2行の上端の距離を計算
                            cm_bottom = t["row_bbox"][1]
                            if cm_bottom < row_top:  # CM行がS2行より上にある
                                distance = row_top - cm_bottom
                                if distance < min_distance:
                                    min_distance = distance
                                    preceding_cm = t["cmno"]
                    
                    # 直前のCMに対応する提供名を取得
                    matching_cm = None
                    matching_teikyou_name = None
                    if preceding_cm and preceding_cm in fmt_s2.teikyou_flags:
                        if fmt_s2.teikyou_flags[preceding_cm]:
                            matching_cm = preceding_cm
                            matching_teikyou_name = fmt_s2.teikyou_names.get(preceding_cm)
                            
                            # 同じCM番号で既に提供名を描画済みの場合、このS2行では描画しない
                            if matching_cm in processed_cm_for_teikyou:
                                matching_teikyou_name = None

                            else:
                                processed_cm_for_teikyou.add(matching_cm)

                    # フォーマット連絡表に独立したS2セルがある場合もマーク対象にする
                    # （teikyou_flagsに該当がなくても、s2_flagsがあればS2行をマークする）
                    if matching_cm is None and preceding_cm and getattr(fmt_s2, 's2_flags', {}).get(preceding_cm, False):
                        matching_cm = preceding_cm
                        print(f"[DEBUG_S2] フォーマット連絡表のS2フラグによりマーク: CM={preceding_cm}", flush=True)
                    
                    # S2のspan
                    s2_span = (s2w["x0"] - 1, s2w["x1"] + 1)
                    print(f"[DEBUG_S2] Line 1091実行: s2w={s2w}, s2_span={s2_span}", flush=True)
                    
                    # VOL列をチェック（BLまたはTKの場合、提供の尺を取得）
                    # VOL列はDT列の右側（x0が約150-200の範囲）
                    vol_x0, vol_x1 = 150, 200
                    vol_words = [w for w in row_words_sorted if vol_x0 <= w["x0"] <= vol_x1]
                    has_bl_or_tk = any(w["text"].startswith("BL") or w["text"].startswith("TK") for w in vol_words)
                    print(f"[DEBUG_S2] VOL check: has_bl_or_tk={has_bl_or_tk}, vol_words={len(vol_words)}", flush=True)
                    print(f"[DEBUG_S2] after_VOL_1: About to process teikyou_actual_sec", flush=True)
                    
                    teikyou_actual_sec = None
                    teikyou_planned_sec = None
                    teikyou_status = "match"  # デフォルトは青
                    teikyou_span = None
                    try:
                        teikyou_dt_span = None
                        # 提供の尺を取得（VOL列が「BL」または「TK」の場合、DT列から尺を取得）
                        if has_bl_or_tk:
                            # DT列から尺を取得
                            dt_x0 = h_dt["x0"] - 10 if h_dt else 100
                            dt_x1 = h_cmno["x0"] - 10 if h_cmno else 140
                            dt_words = [
                                w for w in row_words_sorted
                                if dt_x0 <= w["x0"] <= dt_x1 and (
                                    ("'" in w["text"]) or ('"'  in w["text"])
                                    or ("″" in w["text"]) or ('”' in w["text"])
                                )
                            ]
                            dt_str = " ".join(w["text"] for w in dt_words)
                            teikyou_actual_sec = parse_duration_to_seconds(dt_str)
                            teikyou_dt_span = (min(w["x0"] for w in dt_words) - 1, max(w["x1"] for w in dt_words) + 1) if dt_words else None

                        # フォーマット連絡表の提供尺を取得
                        print(f"[DEBUG_S2] teikyou_actual_sec取得後: teikyou_actual_sec={teikyou_actual_sec}", flush=True)
                        if matching_cm and fmt_s2 is not None:
                            teikyou_planned_sec = fmt_s2.teikyou_seconds.get(matching_cm)

                        # 提供尺の一致判定
                        if teikyou_planned_sec is not None and teikyou_actual_sec is not None:
                            if teikyou_planned_sec != teikyou_actual_sec:
                                teikyou_status = "mismatch"  # 不一致なら赤
                                print(f"[DEBUG] S2行: 番組「{fmt_s2.program}」, CM{matching_cm}, 尺不一致: 予定={teikyou_planned_sec}秒, 実際={teikyou_actual_sec}秒")
                            else:
                                print(f"[DEBUG] S2行: 番組「{fmt_s2.program}」, CM{matching_cm}, 尺一致: {teikyou_planned_sec}秒")
                        elif teikyou_planned_sec is None and matching_cm:
                            print(f"[DEBUG] S2行: 番組「{fmt_s2.program}」, CM{matching_cm}, フォーマット連絡表に提供尺なし")
                        elif matching_cm is None:
                            print(f"[DEBUG] S2行: 番組「{fmt_s2.program}」, CM(なし), フォーマット連絡表に提供尺なし")

                        print(f"[DEBUG_S2] teikyou_status判定後: status={teikyou_status}", flush=True)

                        # 提供名のspan（既に記載されている場合）またはNone（描画が必要）
                        print(f"[DEBUG_S2] teikyou_span初期化直後", flush=True)
                        print(f"[DEBUG_S2] About to check sponsor column (sponsor_x1_s2={sponsor_x1_s2})", flush=True)

                        # スポンサー列から提供名を探す
                        # 「提」から「）」までの全単語を含める（括弧内の説明も含めるため）
                        # スポンサー列の範囲を少し広げて「提」の文字も含める（x0=570から、x1=720まで）
                        sponsor_words_in_range = [w for w in row_words_sorted 
                                                 if 570 <= w["x0"] <= sponsor_x1_s2 and w["x1"] <= sponsor_x1_s2]

                        teikyou_words = []
                        in_teikyou = False
                        for w in sponsor_words_in_range:
                            # 「提」「サイド」で開始
                            if '提' in w["text"] or 'サイド' in w["text"]:
                                in_teikyou = True

                            if in_teikyou:
                                teikyou_words.append(w)

                                # 「）」で終わる場合、終了
                                if '）' in w["text"] or ')' in w["text"]:
                                    break

                        if teikyou_words:
                            # 提供名の文字列全体を取得
                            teikyou_text = " ".join(w["text"] for w in teikyou_words)
                            # オペログから提供キー（アルファベット+記号）を抽出
                            opelog_key = _extract_teikyou_key(teikyou_text)
                            # フォーマット連絡表の提供名キーを抽出
                            fmt_key = _extract_teikyou_key(matching_teikyou_name or "")
                            if opelog_key:
                                # 提供名の文字列全体をマーク
                                teikyou_span = (min(w["x0"] for w in teikyou_words) - 1,
                                               max(w["x1"] for w in teikyou_words) + 1)
                                # フォーマット連絡表のキーと照合
                                if fmt_key and opelog_key != fmt_key:
                                    teikyou_status = "mismatch"
                                    print(f"[DEBUG] S2行: 提供名不一致 opelog={opelog_key!r} fmt={fmt_key!r}")
                                else:
                                    # 一致または比較不能（フォーマット連絡表に提供名なし）→ 青
                                    if teikyou_status != "mismatch":
                                        teikyou_status = "match"
                                # オペログに提供名がある場合は描画しない
                                matching_teikyou_name = None
                    except Exception as e:
                        teikyou_status = "unknown"
                        print(f"[DEBUG_S2] S2処理例外: {e}", flush=True)
                    
                    # S2行をtargetsに追加
                    print(f"[DEBUG_S2] Line 1169前: s2_span={s2_span}, matching_cm={matching_cm}", flush=True)
                    targets.append({
                        "page": pi,
                        "status": teikyou_status,  # 提供尺の一致判定結果
                        "row_bbox": (row_top, row_bottom),
                        "s2_span": s2_span,
                        "teikyou_dt_span": teikyou_dt_span,
                        "teikyou_span": teikyou_span,
                        "teikyou_name": matching_teikyou_name,
                        "teikyou_flag": bool(matching_teikyou_name),
                        "kind": "s2",  # S2行として識別
                        "program": fmt_s2.program if fmt_s2 else "",
                        "cmno": matching_cm or "",
                        "planned_sec": teikyou_planned_sec,
                        "actual_sec": teikyou_actual_sec,
                    })
                    
                    # --- S2行の直後の行（TR列が空でも提供名がある行）の処理 ---
                    # トリガーがない番組では、S2行の直後に提供名が続く場合がある
                    # S2行の下50ポイント以内で、スポンサー列に提供名がある行を探す
                    next_row_range = 50
                    next_rows_words = [w for w in words if s2w["bottom"] < w["top"] < s2w["bottom"] + next_row_range]
                    
                    # y座標でグループ化（同じ行の単語をまとめる）
                    next_rows = {}
                    for w in next_rows_words:
                        y_key = round(w["top"])
                        if y_key not in next_rows:
                            next_rows[y_key] = []
                        next_rows[y_key].append(w)
                    
                    for y_key in sorted(next_rows.keys()):
                        next_row_words = sorted(next_rows[y_key], key=lambda w: w["x0"])
                        
                        # スポンサー列に提供名があるか確認
                        sponsor_words_in_range = [w for w in next_row_words 
                                                 if 570 <= w["x0"] <= sponsor_x1_s2 and w["x1"] <= sponsor_x1_s2]
                        
                        has_teikyou = any('提' in w["text"] or 'サイド' in w["text"] for w in sponsor_words_in_range)
                        
                        if not has_teikyou:
                            continue
                        
                        # 提供名を検出
                        teikyou_words_next = []
                        in_teikyou = False
                        for w in sponsor_words_in_range:
                            if '提' in w["text"] or 'サイド' in w["text"]:
                                in_teikyou = True
                            
                            if in_teikyou:
                                teikyou_words_next.append(w)
                                
                                if '）' in w["text"] or ')' in w["text"]:
                                    break
                        
                        if teikyou_words_next:
                            teikyou_text_next = " ".join(w["text"] for w in teikyou_words_next)
                            has_alphabet = bool(re.search(r'[A-ZＡ-Ｚ][\'\'′]*', teikyou_text_next))
                            
                            if has_alphabet:
                                # 提供名の文字列全体をマーク
                                teikyou_span_next = (min(w["x0"] for w in teikyou_words_next) - 1, 
                                                   max(w["x1"] for w in teikyou_words_next) + 1)
                                
                                # 行の範囲を計算
                                next_row_top = min(w["top"] for w in next_row_words) - 3
                                next_row_bottom = max(w["bottom"] for w in next_row_words) + 3
                                
                                # この行をtargetsに追加（S2行の続きとして）
                                targets.append({
                                    "page": pi,
                                    "status": "match",
                                    "row_bbox": (next_row_top, next_row_bottom),
                                    "teikyou_span": teikyou_span_next,
                                    "teikyou_name": None,  # 既に記載されているのでNone
                                    "teikyou_flag": False,
                                    "kind": "s2_next",  # S2行の続きとして識別
                                    "program": fmt_s2.program if fmt_s2 else "",
                                    "cmno": matching_cm or "",
                                    "planned_sec": None,
                                    "actual_sec": None,
                                })
    
                # --- BL行（ネット受け提供行）の処理 ---
                # 条件: TR=D、DT列に数字あり、VOL=BL、かつAS列にH1/QRX等のネット受け信号あり
                # ※自局提供BL行（AS列が空またはCM1）は d_teikyou として別途処理する
                # マーク対象: D/DT/BL と直下OFセル
                # 尺チェック: フォーマット連絡表のnet_uke_secondsと照合 → 一致:青、不一致:赤
                vol_x0_bl, vol_x1_bl = 140, 170  # VOL列の範囲（BLはx0=151.9）
                d_bl_words = [w for w in words if w["text"] == "D"]
                bl_net_idx = 0  # ネット受け提供の出現インデックス（フォーマット連絡表と対応）
                # ネット受け信号の判定キーワード
                NET_SIGNALS = {"H1", "H2", "QRX", "LN", "LA"}

                for dw_bl in d_bl_words:
                    y_pad = 3
                    row_top_bl = dw_bl["top"] - y_pad
                    row_bottom_bl = dw_bl["bottom"] + y_pad
                    row_words_bl = [w for w in words if not (w["bottom"] < row_top_bl or w["top"] > row_bottom_bl)]

                    # VOL列にBLがあるか確認
                    vol_words_bl = [w for w in row_words_bl if vol_x0_bl <= w["x0"] <= vol_x1_bl]
                    bl_word = next((w for w in vol_words_bl if w["text"] == "BL"), None)
                    if bl_word is None:
                        continue  # BLがなければスキップ

                    # AS列（x0=230〜400）にネット受け信号（H1/QRX等）があるか確認
                    # ない場合は自局提供BL行なので d_teikyou として処理（ここではスキップ）
                    as_words_bl = [w for w in row_words_bl if 230 <= w["x0"] <= 400]
                    has_net_signal = any(w["text"] in NET_SIGNALS for w in as_words_bl)
                    if not has_net_signal:
                        print(f"[DEBUG] bl_net skip（自局提供BL）: top={dw_bl['top']:.1f} as_words={[w['text'] for w in as_words_bl]}", flush=True)
                        continue  # 自局提供BLはd_teikyouループで処理

                    # DT列に数字（尺）があるか確認
                    dt_col_x0 = h_dt["x0"] - 10 if h_dt else 85
                    dt_col_x1 = h_cmno["x0"] - 10 if h_cmno else 140
                    dt_words_bl = [
                        w for w in row_words_bl
                        if dt_col_x0 <= w["x0"] <= dt_col_x1
                        and ("'" in w["text"] or '"' in w["text"] or "″" in w["text"] or "’" in w["text"])
                    ]
                    if not dt_words_bl:
                        continue  # DT列に数字がなければスキップ

                    dt_str_bl = " ".join(w["text"] for w in dt_words_bl)
                    actual_sec_bl = parse_duration_to_seconds(dt_str_bl)
                    if actual_sec_bl is None:
                        continue

                    # フォーマット連絡表のネット受け尺と照合
                    # 時刻から番組を特定
                    tw_bl = _nearest_time_for_y(
                        (dw_bl["top"] + dw_bl["bottom"]) / 2,
                        dw_bl["top"], dw_bl["bottom"]
                    )
                    tsec_bl = parse_hhmmss(tw_bl["text"]) if tw_bl else None
                    fmt_bl = None
                    if tsec_bl is not None:
                        for f in formats:
                            s, e = f.time_window
                            if s <= tsec_bl <= e:
                                fmt_bl = f
                                break

                    planned_sec_bl = None
                    bl_status = "match"  # デフォルト: 青
                    if fmt_bl is not None and fmt_bl.net_uke_seconds:
                        # 出現順でマッチング
                        if bl_net_idx < len(fmt_bl.net_uke_seconds):
                            planned_sec_bl = fmt_bl.net_uke_seconds[bl_net_idx]
                        bl_net_idx += 1
                        if planned_sec_bl is not None and actual_sec_bl != planned_sec_bl:
                            bl_status = "mismatch"

                    # D span
                    d_span_bl = (dw_bl["x0"] - 1, dw_bl["x1"] + 1)
                    # DT span
                    dt_span_bl = (
                        min(w["x0"] for w in dt_words_bl) - 1,
                        max(w["x1"] for w in dt_words_bl) + 1
                    ) if dt_words_bl else None
                    # BL span
                    bl_span = (bl_word["x0"] - 1, bl_word["x1"] + 1)

                    # 直下OFを探す（50pt以内）
                    of_span = None
                    of_row_bbox = None
                    of_words_nearby = [
                        w for w in words
                        if w["text"] == "OF"
                        and vol_x0_bl <= w["x0"] <= vol_x1_bl
                        and 0 < w["top"] - dw_bl["bottom"] < 60
                    ]
                    if of_words_nearby:
                        ow = min(of_words_nearby, key=lambda w: w["top"])
                        of_span = (ow["x0"] - 1, ow["x1"] + 1)
                        of_row_bbox = (ow["top"] - y_pad, ow["bottom"] + y_pad)

                    targets.append({
                        "page": pi,
                        "status": bl_status,
                        "row_bbox": (row_top_bl, row_bottom_bl),
                        "kind": "bl_net",
                        "d_span_bl": d_span_bl,
                        "dt_span_bl": dt_span_bl,
                        "bl_span": bl_span,
                        "of_span": of_span,
                        "of_row_bbox": of_row_bbox,
                        "program": fmt_bl.program if fmt_bl else "",
                        "cmno": "",
                        "planned_sec": planned_sec_bl,
                        "actual_sec": actual_sec_bl,
                    })
                    print(f"[DEBUG] BL行: P{pi+1} top={dw_bl['top']:.1f} 尺={dt_str_bl} ({actual_sec_bl}秒) 予定={planned_sec_bl}秒 status={bl_status}", flush=True)

                # --- D行（CM直後の提供行）の処理 ---
                # TR列に「D」があり、スポンサー列に「提」または「サイド」がある行を抽出
                # これらはCM直後の提供行で、直前のCMに対応する提供名を描画またはマーク
                # 同じCM番号で複数のD行がある場合、最初の1行のみに提供名を描画（processed_cm_for_teikyouで管理）
                d_words = [w for w in words if w["text"] == "D"]
                d_span = None  # ← ループ直前で再度初期化
                print(f"[DEBUG] ページ{pi+1}: D行の数={len(d_words)}", flush=True)
                
                for dw in d_words:
                    print(f"[DEBUG] D行ループ開始: dw={dw}", flush=True)
                    # 初期化
                    d_span = None
                    teikyou_status = "match"
                    teikyou_planned_sec = None
                    teikyou_actual_sec = None
                    matching_cm = None
                    matching_teikyou_name = None
                    teikyou_span = None
                    
                    y_pad = 3
                    row_top = dw["top"] - y_pad
                    row_bottom = dw["bottom"] + y_pad
                    row_words = [w for w in words if not (w["bottom"] < row_top or w["top"] > row_bottom)]
                    row_words_sorted = sorted(row_words, key=lambda x: x["x0"])
                    
                    # VOL列（x0=150-220）に「BL」または「TK」があるか確認
                    # ※C/提列の「提」ではなくVOL列のBLで提供Dトリガー行を判定する
                    vol_check_words = [w for w in row_words_sorted if 150 <= w["x0"] <= 220]
                    has_bl_vol = any(w["text"].startswith("BL") or w["text"].startswith("TK") for w in vol_check_words)
                    print(f"[DEBUG] D行: has_bl_vol={has_bl_vol}, vol_words={[w['text'] for w in vol_check_words]}", flush=True)
                    
                    if not has_bl_vol:
                        continue
                    
                    # 開始時刻を取得
                    row_text = " ".join(w["text"] for w in row_words_sorted)
                    tsec = parse_hhmmss(row_text)
                    if tsec is None:
                        continue
                    
                    # フォーマット連絡表と突合（時刻範囲で番組を特定）
                    fmt_d = None
                    for f in formats:
                        s, e = f.time_window
                        if s <= tsec <= e:
                            fmt_d = f
                            break
                    
                    if fmt_d is None:
                        continue
                    
                    # このD行の直前のCM行を探す（同じページ内で、このD行より上にある最も近いCM行）
                    preceding_cm = None
                    min_distance = float('inf')
                    
                    for t in targets:
                        if t.get("kind") == "cm" and t["page"] == pi and t.get("program") == getattr(fmt_d, "program", ""):
                            # CM行の下端とD行の上端の距離を計算
                            cm_bottom = t["row_bbox"][1]
                            if cm_bottom < row_top:  # CM行がD行より上にある
                                distance = row_top - cm_bottom
                                if distance < min_distance:
                                    min_distance = distance
                                    preceding_cm = t["cmno"]
                    
                    # 直前のCMに対応する提供名を取得
                    if preceding_cm and preceding_cm in fmt_d.teikyou_flags:
                        if fmt_d.teikyou_flags[preceding_cm]:
                            matching_cm = preceding_cm
                            matching_teikyou_name = fmt_d.teikyou_names.get(preceding_cm)
                            
                            # 同じCM番号で既に提供名を描画済みの場合、このD行では描画しない
                            if matching_cm in processed_cm_for_teikyou:
                                matching_teikyou_name = None
                            else:
                                processed_cm_for_teikyou.add(matching_cm)
                    
                    # Dのspan
                    print(f"[DEBUG] D行処理: dw={dw}")
                    d_span = (dw["x0"] - 1, dw["x1"] + 1)
                    print(f"[DEBUG] d_span定義完了: {d_span}")
                    
                    # VOL列をチェック（BLまたはTKの場合、提供の尺を取得）
                    # VOL列はDT列の右側（x0が約150-200の範囲）
                    vol_x0, vol_x1 = 150, 200
                    vol_words = [w for w in row_words_sorted if vol_x0 <= w["x0"] <= vol_x1]
                    has_bl_or_tk = any(w["text"].startswith("BL") or w["text"].startswith("TK") for w in vol_words)
                    
                    # 提供の尺を取得（VOL列が「BL」または「TK」の場合、DT列から尺を取得）
                    teikyou_actual_sec = None
                    teikyou_dt_span = None
                    if has_bl_or_tk:
                        # DT列から尺を取得
                        dt_x0 = h_dt["x0"] - 10 if h_dt else 100
                        dt_x1 = h_cmno["x0"] - 10 if h_cmno else 140
                        dt_words = [
                            w for w in row_words_sorted
                            if dt_x0 <= w["x0"] <= dt_x1 and (
                                ("'" in w["text"]) or ('"'  in w["text"])
                                or ("″" in w["text"]) or ('”' in w["text"])
                            )
                        ]
                        dt_str = " ".join(w["text"] for w in dt_words)
                        teikyou_actual_sec = parse_duration_to_seconds(dt_str)
                        teikyou_dt_span = (min(w["x0"] for w in dt_words) - 1, max(w["x1"] for w in dt_words) + 1) if dt_words else None

                    # フォーマット連絡表の提供尺を取得
                    teikyou_planned_sec = None
                    if matching_cm and fmt_d is not None:
                        teikyou_planned_sec = fmt_d.teikyou_seconds.get(matching_cm)
                    
                    # 提供尺の一致判定
                    teikyou_status = "match"  # デフォルトは青
                    if teikyou_planned_sec is not None and teikyou_actual_sec is not None:
                        if teikyou_planned_sec != teikyou_actual_sec:
                            teikyou_status = "mismatch"  # 不一致なら赤
                            print(f"[DEBUG] D行: 番組「{getattr(fmt_d, 'program', '')}」, CM{matching_cm}, 尺不一致: 予定={teikyou_planned_sec}秒, 実際={teikyou_actual_sec}秒")
                        else:
                            print(f"[DEBUG] D行: 番組「{getattr(fmt_d, 'program', '')}」, CM{matching_cm}, 尺一致: {teikyou_planned_sec}秒")
                    elif teikyou_planned_sec is None and matching_cm:
                        print(f"[DEBUG] D行: 番組「{getattr(fmt_d, 'program', '')}」, CM{matching_cm}, フォーマット連絡表に提供尺なし")
                    
                    # 提供名のspan（既に記載されている場合）またはNone（描画が必要）
                    teikyou_span = None
                    
                    # スポンサー列から提供名を探す
                    # 「提」から「）」までの全単語を含める（括弧内の説明も含めるため）
                    # スポンサー列の範囲を少し広げて「提」の文字も含める（x0=570から、x1=720まで）
                    sponsor_words_in_range = [w for w in row_words_sorted 
                                             if 570 <= w["x0"] <= sponsor_x1_s2 and w["x1"] <= sponsor_x1_s2]
                    
                    teikyou_words = []
                    in_teikyou = False
                    for w in sponsor_words_in_range:
                        # 「提」「サイド」で開始
                        if '提' in w["text"] or 'サイド' in w["text"]:
                            in_teikyou = True
                        
                        if in_teikyou:
                            teikyou_words.append(w)
                            
                            # 「）」で終わる場合、終了
                            if '）' in w["text"] or ')' in w["text"]:
                                break
                    
                    # teikyou_wordsに「提供X」「サイドX」形式の提供名があるか確認
                    # 単独の「提」（C/提 列から拾ったもの）は無視する
                    valid_teikyou_words = []
                    if teikyou_words:
                        teikyou_text_check = " ".join(w["text"] for w in teikyou_words)
                        # 「提供」または「サイド」＋英字が含まれているか確認
                        if _extract_teikyou_key(teikyou_text_check):
                            valid_teikyou_words = teikyou_words
                            print(f"[DEBUG] D行: スポンサー列に提供名あり: {repr(teikyou_text_check)}")
                        else:
                            print(f"[DEBUG] D行: スポンサー列の「提」は単独のため無視: {repr(teikyou_text_check)}")
                    
                    if valid_teikyou_words:
                        # 提供名の文字列全体を取得
                        teikyou_text = " ".join(w["text"] for w in valid_teikyou_words)
                        opelog_key = _extract_teikyou_key(teikyou_text)
                        
                        # フォーマット連絡表のall_teikyou_namesから出現順に対応する提供名を取得
                        current_d_teikyou_count_all = sum(
                            1 for t in targets
                            if t.get("kind") in ("d_teikyou", "s2")
                            and t.get("program") == (fmt_d.program if fmt_d else "")
                        )
                        fmt_teikyou_name = None
                        if fmt_d and fmt_d.all_teikyou_names:
                            idx = current_d_teikyou_count_all
                            if idx < len(fmt_d.all_teikyou_names):
                                fmt_teikyou_name = fmt_d.all_teikyou_names[idx]
                        
                        fmt_key = _extract_teikyou_key(fmt_teikyou_name or "")
                        print(f"[DEBUG] D行: 提供名照合 opelog={repr(teikyou_text)} opelog_key={repr(opelog_key)} fmt[{current_d_teikyou_count_all}]={repr(fmt_teikyou_name)} fmt_key={repr(fmt_key)}")
                        
                        if opelog_key:
                            teikyou_span = (min(w["x0"] for w in valid_teikyou_words) - 1,
                                           max(w["x1"] for w in valid_teikyou_words) + 1)
                            if fmt_key and opelog_key != fmt_key:
                                teikyou_status = "mismatch"
                                print(f"[DEBUG] D行: 提供名不一致 opelog={opelog_key!r} fmt={fmt_key!r}")
                            else:
                                if teikyou_status != "mismatch":
                                    teikyou_status = "match"
                            matching_teikyou_name = fmt_teikyou_name  # スポンサー列表示用
                    
                    # D行（提供）をtargetsに追加
                    # fmt_d は continue で抜けない限り必ず定義されている
                    # teikyou_spanがある = オペログに提供名が既記載
                    # → teikyou_nameはスポンサー列に描画するフォーマット連絡表の対応提供名
                    # teikyou_spanがない = オペログに提供名なし
                    # → teikyou_nameを後処理で割り当て（all_teikyou_namesから順番に）
                    targets.append({
                        "page": pi,
                        "status": teikyou_status,  # 提供尺の一致判定結果
                        "row_bbox": (row_top, row_bottom),
                        "d_span": d_span,  # D列のspan
                        "teikyou_dt_span": teikyou_dt_span,
                        "teikyou_span": teikyou_span,
                        "teikyou_name": matching_teikyou_name,  # スポンサー列描画用（フォーマット連絡表対応名）
                        "teikyou_flag": bool(matching_teikyou_name),
                        "opelog_teikyou_found": bool(teikyou_span),  # オペログに既記載の提供名があったか
                        "kind": "d_teikyou",  # D行（提供）として識別
                        "program": fmt_d.program if fmt_d else "",
                        "cmno": matching_cm or "",
                        "planned_sec": teikyou_planned_sec,
                        "actual_sec": teikyou_actual_sec,
                    })
                
                # --- TR列が空で、スポンサー列に提供名がある行の処理 ---
                # トリガーがない番組では、TR列が空でも提供名が記載されている行がある
                # これらの行を検出してマークする
                # 開始時刻がある行（時刻列に「:」がある）で、TR列が空で、スポンサー列に提供名がある行を探す
                time_words_all = [w for w in words if ":" in w["text"] and w["x0"] < 100]
                
                for tw in time_words_all:
                    y_pad = 3
                    row_top = tw["top"] - y_pad
                    row_bottom = tw["bottom"] + y_pad
                    row_words = [w for w in words if not (w["bottom"] < row_top or w["top"] > row_bottom)]
                    row_words_sorted = sorted(row_words, key=lambda w: w["x0"])
                    
                    # TR列が空か確認（S2、D、R、U、CMなどがない）
                    tr_words = [w for w in row_words_sorted if 340 <= w["x0"] <= 360]
                    has_tr = any(w["text"] in ("S2", "D", "R", "U", "CM", "S", "DT", "ON") for w in tr_words)
                    
                    if has_tr:
                        continue  # TR列に何かある場合はスキップ（既に処理済み）
                    
                    # 開始時刻を取得してフォーマット連絡表と突合
                    row_text = " ".join(w["text"] for w in row_words_sorted)
                    tsec = parse_hhmmss(row_text)
                    if tsec is None:
                        continue
                    
                    # フォーマット連絡表と突合（時刻範囲で番組を特定）
                    fmt_tr_empty = None
                    for f in formats:
                        s, e = f.time_window
                        if s <= tsec <= e:
                            fmt_tr_empty = f
                            break
                    
                    if fmt_tr_empty is None:
                        print(f"[DEBUG] ページ{pi+1}: 時刻{tw['text']}はフォーマット連絡表にない → スキップ")
                        continue  # フォーマット連絡表にない番組はスキップ
                    
                    # この番組に提供があるかチェック（all_teikyou_namesが空でない）
                    if not fmt_tr_empty.all_teikyou_names:
                        print(f"[DEBUG] ページ{pi+1}: 番組「{fmt_tr_empty.program}」は提供なし → スキップ")
                        continue  # 提供がない番組はスキップ
                    
                    # スポンサー列に提供名があるか確認
                    sponsor_words_in_range = [w for w in row_words_sorted 
                                             if 570 <= w["x0"] <= sponsor_x1_s2 and w["x1"] <= sponsor_x1_s2]
                    
                    has_teikyou = any('提' in w["text"] or 'サイド' in w["text"] for w in sponsor_words_in_range)
                    
                    if not has_teikyou:
                        continue
                    
                    # 提供名を検出
                    teikyou_words_tr_empty = []
                    in_teikyou = False
                    for w in sponsor_words_in_range:
                        if '提' in w["text"] or 'サイド' in w["text"]:
                            in_teikyou = True
                        
                        if in_teikyou:
                            teikyou_words_tr_empty.append(w)
                            
                            if '）' in w["text"] or ')' in w["text"]:
                                break
                    
                    if teikyou_words_tr_empty:
                        teikyou_text_tr_empty = " ".join(w["text"] for w in teikyou_words_tr_empty)
                        has_alphabet = bool(re.search(r'[A-ZＡ-Ｚ][\'\'′]*', teikyou_text_tr_empty))
                        
                        print(f"[DEBUG] ページ{pi+1}: 時刻{tw['text']}, 番組「{fmt_tr_empty.program}」, 提供名「{teikyou_text_tr_empty}」, アルファベット={has_alphabet}")
                        
                        if has_alphabet:
                            # 提供名の文字列全体をマーク
                            teikyou_span_tr_empty = (min(w["x0"] for w in teikyou_words_tr_empty) - 1, 
                                                   max(w["x1"] for w in teikyou_words_tr_empty) + 1)
                            
                            # この行をtargetsに追加（TR列が空の提供行として）
                            targets.append({
                                "page": pi,
                                "status": "match",
                                "row_bbox": (row_top, row_bottom),
                                "teikyou_span": teikyou_span_tr_empty,
                                "teikyou_name": None,  # 既に記載されているのでNone
                                "teikyou_flag": False,
                                "kind": "teikyou_no_tr",  # TR列が空の提供行として識別
                                "program": fmt_tr_empty.program if fmt_tr_empty else "",
                                "cmno": "",
                                "planned_sec": None,
                                "actual_sec": None,
                            })
    
    
        # --- 提供名の割り当て（all_teikyou_namesから順番に） ---
        # S2行とD行（提供あり）に対して、フォーマット連絡表のall_teikyou_namesから順番に提供名を割り当てる
        # オペログに既記載の提供名がある行も含めて順番でカウントする（CM番号との紐付けなし）
        for fmt in formats:
            # この番組の全提供行を時刻順に抽出（S2行＋D行）
            program_teikyou_targets_all = [
                t for t in targets 
                if t.get("kind") in ("s2", "d_teikyou") 
                and t.get("program") == fmt.program
            ]
            
            # 時刻順にソート
            program_teikyou_targets_sorted = sorted(program_teikyou_targets_all, key=lambda x: (x["page"], x["row_bbox"][0]))
            
            # all_teikyou_namesから順番に割り当て
            for i, t in enumerate(program_teikyou_targets_sorted):
                fmt_name = fmt.all_teikyou_names[i] if i < len(fmt.all_teikyou_names) else None
                
                if t.get("opelog_teikyou_found"):
                    # オペログに既記載 → teikyou_nameはスポンサー列表示用（既に設定済みのfmt対応名を使う）
                    # ただし未設定の場合はfmt_nameを使う
                    if not t.get("teikyou_name") and fmt_name:
                        t["teikyou_name"] = fmt_name
                        t["teikyou_flag"] = True
                elif not t.get("teikyou_name") and not t.get("teikyou_span"):
                    # オペログに提供名なし → フォーマット連絡表から描画
                    if fmt_name:
                        t["teikyou_name"] = fmt_name
                        t["teikyou_flag"] = True
    
    
        # --- CMNo単位で合算して一致/不一致を確定（例: LT+PT の合計） ---
        # program + cmno ごとに actual_sec を合算し、planned と比較して各行の status に反映
        # ただし、S2行、D行、TR列が空の提供行、S2行の直後の行は除外（これらは提供の尺チェック用）
        sums: Dict[Tuple[str, str], int] = {}
        planned_map: Dict[Tuple[str, str], Optional[int]] = {}
        for t in targets:
            # S2行、D行、TR列が空の提供行、S2行の直後の行は除外
            if t.get("kind") in ("s2", "d_teikyou", "teikyou_no_tr", "s2_next", "bl_net"):
                continue
            
            key = (t["program"], t["cmno"])
            if t.get("actual_sec") is not None:
                sums[key] = sums.get(key, 0) + int(t["actual_sec"])
            # planned_secがNoneでない場合のみplanned_mapに設定（S2行やD行でNoneに上書きされないように）
            if t.get("planned_sec") is not None:
                planned_map[key] = t.get("planned_sec")
    
        for t in targets:
            # TR欄(R/U/D)のマーカーは「突合(尺)」とは無関係なので、色判定で上書きしない
            if t.get("kind") == "tr":
                t["status"] = "match"
                t["actual_total_sec"] = None
                continue
            
            # S2行、D行、TR列が空の提供行、S2行の直後の行は提供尺チェック済みなので、色判定で上書きしない
            if t.get("kind") in ("s2", "d_teikyou", "teikyou_no_tr", "s2_next", "bl_net"):
                t["actual_total_sec"] = t.get("actual_sec")
                continue
    
            key = (t["program"], t["cmno"])
            planned = planned_map.get(key)
            actual_total = sums.get(key)
            # 黄色（unknown）を廃止：plannedまたはactual_totalがNoneの場合は赤（mismatch）とする
            if planned is None or actual_total is None:
                t["status"] = "mismatch"  # unknownからmismatchに変更
            else:
                t["status"] = "match" if int(planned) == int(actual_total) else "mismatch"
            t["actual_total_sec"] = actual_total
    
    
        # --- 同一番組内にトリガー(R/U)が混在している場合のみ、TR空欄行の開始時刻をマーク ---
        program_has_trigger: Dict[str, bool] = {}
        for t in targets:
            prog = t.get("program") or ""
            if t.get("tr_char") in ("R", "U"):
                program_has_trigger[prog] = True
    
        for t in targets:
            prog = t.get("program") or ""
            if t.get("tr_char") is None and program_has_trigger.get(prog):
                t["start_span"] = t.get("start_span_candidate")
            else:
                t["start_span"] = None
    
        # de-dup NF per page+program
        seen = set()
        nf_unique = []
        for n in nf_labels:
            key = (n["page"], n["program"])
            if key in seen:
                continue
            seen.add(key)
            nf_unique.append(n)
    
    return targets, nf_unique


# ----------------------------
# PDF generation
# ----------------------------
def hex_to_rgb01(h: str) -> Tuple[float, float, float]:
    h = (h or "").strip().lstrip("#")
    if len(h) != 6:
        return (0.1, 0.3, 1.0)
    r = int(h[0:2], 16) / 255.0
    g = int(h[2:4], 16) / 255.0
    b = int(h[4:6], 16) / 255.0
    return (r, g, b)


def make_overlay_pdf(
    reader: PdfReader,
    targets,
    nf_labels,
    s2_targets=None,
    alpha=0.20,
    match_rgb=(0.1, 0.3, 1.0),
    mismatch_rgb=(1.0, 0.2, 0.2),
    unknown_rgb=(1.0, 1.0, 0.2),
    nf_rgb=(0.1, 0.3, 1.0),
    nf_font_size=12,
    s2_rgb=None,
    teikyou_match_rgb=None,
):
    buf = io.BytesIO()
    c = canvas.Canvas(buf)

    # 日本語フォントを登録（reportlab標準のCIDフォント）
    try:
        pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))
        japanese_font = 'HeiseiMin-W3'
    except:
        try:
            pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
            japanese_font = 'HeiseiKakuGo-W5'
        except:
            japanese_font = 'Helvetica'  # フォールバック

    # 「NF」が付与された（＝フォーマット連絡表が発行されている）番組だけをマーク対象にする
    nf_programs = {n.get("program") for n in nf_labels if n.get("program")}

    # s2_rgb が未指定なら match_rgb を使う（S2 は既定で一致色）
    if s2_rgb is None:
        s2_rgb = match_rgb

    for pi, page in enumerate(reader.pages):
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        c.setPageSize((w, h))

        # NF番組以外はマークしない（※NF表示そのものは別処理）
        # ただし、kind="teikyou_no_tr"（TR列が空の提供行）は常にマークする
        for t in [x for x in targets if x["page"] == pi and (x.get("program") in nf_programs or x.get("kind") == "teikyou_no_tr" or x.get("kind") == "s2_next" or x.get("kind") == "bl_net")]:
            top, bottom = t["row_bbox"]
            y1 = h - top
            y0 = h - bottom

            # TR列（R/U/D）は常に指定色（デフォルト青）で塗る。
            # それ以外（DT/CMNo/ADDR）は判定色（一致=青、不一致=赤、不明=黄）を維持。
            if t.get("status") == "match":
                col_main = Color(match_rgb[0], match_rgb[1], match_rgb[2], alpha=alpha)
            elif t.get("status") == "mismatch":
                col_main = Color(mismatch_rgb[0], mismatch_rgb[1], mismatch_rgb[2], alpha=alpha)
            else:
                col_main = Color(unknown_rgb[0], unknown_rgb[1], unknown_rgb[2], alpha=alpha)
            col_tr = Color(s2_rgb[0], s2_rgb[1], s2_rgb[2], alpha=alpha)
            # bl_net（ネット受け提供）専用の描画
            if t.get("kind") == "bl_net":
                bl_color = Color(teikyou_match_rgb[0], teikyou_match_rgb[1], teikyou_match_rgb[2], alpha=alpha) if t.get("status") != "mismatch" else Color(mismatch_rgb[0], mismatch_rgb[1], mismatch_rgb[2], alpha=alpha)
                # D・DT・BLをまとめて同じ色でマーク
                for bl_key in ("d_span_bl", "dt_span_bl", "bl_span"):
                    span = t.get(bl_key)
                    if span:
                        x0, x1 = span
                        c.setFillColor(bl_color)
                        c.rect(x0, y0, x1 - x0, y1 - y0, fill=1, stroke=0)
                # OFは同行ではなく別行なので別途描画
                of_span = t.get("of_span")
                of_row_bbox = t.get("of_row_bbox")
                if of_span and of_row_bbox:
                    of_top, of_bottom = of_row_bbox
                    of_y1 = h - of_top
                    of_y0 = h - of_bottom
                    ox0, ox1 = of_span
                    c.setFillColor(bl_color)
                    c.rect(ox0, of_y0, ox1 - ox0, of_y1 - of_y0, fill=1, stroke=0)
                continue  # bl_net は以下の汎用spanループをスキップ

            for key in ("dt_span", "cm_span", "waku_span", "tr_span", "start_span", "s2_span", "d_span", "teikyou_span", "teikyou_dt_span"):
                span = t.get(key)
                if not span:
                    continue
                x0, x1 = span

                # TR列、S2列は指定色固定
                # 提供列（s2_span, d_span, teikyou_span）は一致/不一致で色を変える
                if key in ("tr_span", "start_span"):
                    c.setFillColor(col_tr)
                elif key in ("s2_span", "d_span", "teikyou_span", "teikyou_dt_span"):
                    if t.get("status") == "mismatch":
                        c.setFillColor(Color(mismatch_rgb[0], mismatch_rgb[1], mismatch_rgb[2], alpha=alpha))
                    else:
                        c.setFillColor(Color(teikyou_match_rgb[0], teikyou_match_rgb[1], teikyou_match_rgb[2], alpha=alpha))
                else:
                    c.setFillColor(col_main)

                # ★TR列のRマーク：めざましテレビだけは「R文字の高さ」に合わせて塗る（Dまで塗らない）
                if key == "tr_span" and "めざましテレビ" in (t.get("program") or ""):
                    yspan = t.get("tr_y_span")
                    if yspan:
                        top, bottom = yspan
                        _y0 = h - bottom
                        _y1 = h - top
                        c.rect(x0, _y0, x1 - x0, _y1 - _y0, fill=1, stroke=0)
                        continue

                c.rect(x0, y0, x1 - x0, y1 - y0, fill=1, stroke=0)

            # 提供名の描画
            # ケース1: オペログに提供名の記載がない（teikyou_spanなし）→ フォーマット連絡表の文言を描画
            # ケース2: オペログに提供名の記載がある（teikyou_spanあり）→ スポンサー列にフォーマット連絡表の対応提供名を描画
            if t.get("kind") in ("s2", "d_teikyou") and t.get("teikyou_flag") and t.get("teikyou_name"):
                teikyou_name = t.get("teikyou_name")
                display_name = teikyou_name
                
                x_teikyou = 593
                y_teikyou = (y0 + y1) / 2.0 - 3
                c.setFillColor(Color(match_rgb[0], match_rgb[1], match_rgb[2], alpha=1.0))
                c.setFont(japanese_font, 10)
                c.drawString(x_teikyou, y_teikyou, display_name)

        # NF (always NF color)
        c.setFillColor(Color(nf_rgb[0], nf_rgb[1], nf_rgb[2], alpha=1.0))
        c.setFont("Helvetica-Bold", nf_font_size)
        for n in [x for x in nf_labels if x["page"] == pi]:
            x = n["x"]
            y = h - n["top"] + 10
            c.drawString(x, y, "NF")

        c.showPage()

    c.save()
    buf.seek(0)
    return buf


def generate_marked_pdf(
    opelog_bytes: bytes,
    targets,
    nf_labels,
    s2_targets=None,
    alpha=0.20,
    match_rgb=(0.1, 0.3, 1.0),
    nf_rgb=(0.1, 0.3, 1.0),
    nf_font_size=12,
    s2_rgb=None,
    teikyou_match_rgb=None
) -> bytes:
    """
    マーキング済みPDFを生成する。
    - 既存: CM行マーカー（targets）, NF表示（nf_labels）
    - 追加: S2マーカー（TR列の"S2"文字そのもの, NF番組のみ）
    """

    if s2_targets is None:
        s2_targets = []
    if s2_rgb is None:
        s2_rgb = match_rgb
    if teikyou_match_rgb is None:
        teikyou_match_rgb = match_rgb

    reader = PdfReader(io.BytesIO(opelog_bytes))

    # 既存 overlay 生成処理を拡張
    overlay_buf = make_overlay_pdf(
        reader,
        targets,
        nf_labels,
        s2_targets=s2_targets,
        alpha=alpha,
        match_rgb=match_rgb,
        nf_rgb=nf_rgb,
        nf_font_size=nf_font_size,
        s2_rgb=s2_rgb,
        teikyou_match_rgb=teikyou_match_rgb
    )

    overlay_reader = PdfReader(overlay_buf)

    writer = PdfWriter()
    for i, p in enumerate(reader.pages):
        # S2対応 overlay を重ねる
        p.merge_page(overlay_reader.pages[i])
        writer.add_page(p)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()



def build_result_csv(targets) -> bytes:
    df = pd.DataFrame([{
        "program": t["program"],
        "cmno": t["cmno"],
        "tr_marked": bool(t.get("tr_span")),
        "planned_sec": t["planned_sec"],
        "actual_sec": t["actual_sec"],
        "status": t["status"],
        "page": t["page"] + 1,
    } for t in targets if t.get("kind") != "tr"])
    bio = io.BytesIO()
    df.to_csv(bio, index=False, encoding="utf-8-sig")
    return bio.getvalue()


# ----------------------------
# Streamlit UI
# ----------------------------
st.set_page_config(page_title="CMチェックツール", layout="wide")
st.title("CMチェックツール（オペログPDF × フォーマット連絡表CSV/Excel）")
st.caption("現在日付（JST）: 2025-12-22")

with st.sidebar:
    st.header("設定")
    match_hex = st.color_picker("一致（青）の色", value="#1A4DFF")
    teikyou_match_hex = st.color_picker("提供尺一致の色", value="#1A4DFF")
    nf_hex = st.color_picker("NF表示色（デフォルト青）", value="#1A4DFF")
    alpha = st.slider("マーカー透明度", min_value=0.05, max_value=0.40, value=0.20, step=0.01)
    nf_font_size = st.slider("NF文字サイズ", min_value=8, max_value=18, value=12, step=1)

st.subheader("入力")
opelog_file = st.file_uploader("オペログPDF（1日分）", type=["pdf"], accept_multiple_files=False)
fmt_files = st.file_uploader("フォーマット連絡表（CSV/Excel）※複数可", type=["csv", "xlsx", "xls"], accept_multiple_files=True)

if opelog_file and fmt_files:
    formats = load_formats(fmt_files)
    if not formats:
        st.error("フォーマット連絡表の解析に失敗しました。CSV/Excel様式を確認してください。")
        st.stop()

    st.write("✅ 解析できたフォーマット連絡表（発行番組）")
    st.dataframe(pd.DataFrame([{
        "番組名": f.program,
        "放送時刻": f"{f.time_window[0]//3600:02d}:{(f.time_window[0]%3600)//60:02d}:{f.time_window[0]%60:02d}"
                 f" - {f.time_window[1]//3600:02d}:{(f.time_window[1]%3600)//60:02d}:{f.time_window[1]%60:02d}",
        "CM件数（予定）": len(f.planned),
        "入力ファイル": f.source_name,
    } for f in formats]), use_container_width=True)

    if st.button("✅ 生成（マーキングPDF + 結果CSV）", type="primary"):
        opelog_bytes = opelog_file.getvalue()
        targets, nf_labels = extract_targets_and_nf(opelog_bytes, formats)
        # --- S2 stage1 debug view (log only) ---
        if ENABLE_S2_STAGE1_LOG_ONLY and S2_DEBUG_LOG:
            with st.expander('S2検出ログ（テスト・描画なし）', expanded=False):
                st.write(f'検出件数: {len(S2_DEBUG_LOG)}')
                st.dataframe(S2_DEBUG_LOG, use_container_width=True)


        match_rgb = hex_to_rgb01(match_hex)
        teikyou_match_rgb = hex_to_rgb01(teikyou_match_hex)
        nf_rgb = hex_to_rgb01(nf_hex)

        marked_pdf = generate_marked_pdf(
            opelog_bytes, targets, nf_labels,
            alpha=alpha, match_rgb=match_rgb, nf_rgb=nf_rgb, nf_font_size=nf_font_size,
            teikyou_match_rgb=teikyou_match_rgb
        )
        result_csv = build_result_csv(targets)

        st.success(f"生成完了: 対象CM行={len(targets)} / NF表示={len(nf_labels)}")

        st.download_button("📄 マーキング済みPDFをダウンロード", data=marked_pdf, file_name="opelog_marked.pdf", mime="application/pdf")
        st.download_button("🧾 判定結果CSVをダウンロード", data=result_csv, file_name="result.csv", mime="text/csv")

        st.subheader("簡易プレビュー（先頭20件）")
        st.dataframe(pd.DataFrame(targets).head(20), use_container_width=True)
else:
    st.info("オペログPDF（1日分）と、フォーマット連絡表（CSV/Excel 複数）をアップロードしてください。")