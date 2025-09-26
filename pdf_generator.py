## `pdf_generator.py`"weight":     data_no_idx[DATA_COLUMNS.index("重量(kg)")],   # ← 保留原字串（可能是空字串）
from fpdf import FPDF
import sys
import os
import re
from typing import List, Dict, Tuple

# ======================== 設定區 ========================
MAX_ROWS_PER_PDF = 20           # 每頁最多顯示筆數
FONT_PATH = "NotoSansTC-Regular.ttf"   # CJK 主字型
MAIN_FONT_NAME = "NotoSansTC"
FRACTION_FONT_PATH = "DejaVuSans.ttf"  # 類別欄專用（支援 ⅜、¼…）
FRACTION_FONT_NAME = "DejaVuSans"

PER_COL_FONT_SIZE: Dict[str, int] = {
    "日期": 10, "訂單號碼": 10, "類別": 10, "顏色(組)": 9,
    "數量(片)": 10, "單價": 10, "重量(kg)": 10, "金額": 10, "備註": 9,
}
LINE_H = 7
HEADER_H = 10
COL_WIDTHS: List[int] = [24, 26, 18, 28, 18, 18, 18, 20, 20]  # 總寬需為 190

REPORT_TITLE_LEFT = "捷盛針織企業社"
REPORT_ADDR = "地址：新北市樹林區田尾街211-2號"
REPORT_TEL = "電話：8970-2937 / 8970-3534    傳真：8970-2936"

last_saved_dir = os.getcwd()

# ======================== 公用工具 ========================
def set_last_saved_dir(path: str) -> None:
    global last_saved_dir
    last_saved_dir = path

def resource_path(relative_path: str) -> str:
    try:
        return os.path.join(sys._MEIPASS, relative_path)
    except Exception:
        return os.path.abspath(relative_path)

def ensure_fonts(pdf: FPDF) -> None:
    """確保兩套字型已加入（重複加入會自動忽略）。"""
    try:
        pdf.add_font(MAIN_FONT_NAME, '', resource_path(FONT_PATH), uni=True)
    except Exception:
        pass
    try:
        pdf.add_font(FRACTION_FONT_NAME, '', resource_path(FRACTION_FONT_PATH), uni=True)
    except Exception:
        pass

# ====================== 金額中文大寫 =====================
def number_to_chinese(n: int | str) -> str:
    """整數金額轉中文大寫，結尾加『元整』。支援萬/億/兆/京。"""
    if int(n) == 0:
        return "零元整"
    digits = "零壹貳參肆伍陸柒捌玖"
    unit1 = ["", "拾", "佰", "仟"]
    unit2 = ["", "萬", "億", "兆", "京"]
    s = str(int(n))
    groups: List[str] = []
    while s:
        groups.insert(0, s[-4:].rjust(4, '0'))
        s = s[:-4]
    parts: List[str] = []
    for gi, g in enumerate(groups):
        seg = ""; zero = False
        for i, ch in enumerate(g):
            d = int(ch); pos = 3 - i
            if d == 0:
                zero = True
            else:
                if zero and seg:
                    seg += "零"
                seg += digits[d] + (unit1[pos] if pos > 0 else "")
                zero = False
        if seg:
            seg += unit2[len(groups) - 1 - gi]
        parts.append(seg)
    res = "".join(parts).rstrip("零")
    return res + " 元整"

# ============= 判斷中文字 / 分數 ASCII 化 =============
_CJK_RE = re.compile(r"[\u3400-\u9FFF\uF900-\uFAFF]")
_SUP_TO_NORM = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹", "0123456789")
_SUB_TO_NORM = str.maketrans("₀₁₂₃₄₅₆₇₈₉", "0123456789")
_VULGAR_TO_PAIR: Dict[str, Tuple[int, int]] = {
    "½": (1,2), "⅓": (1,3), "⅔": (2,3),
    "¼": (1,4), "¾": (3,4),
    "⅕": (1,5), "⅖": (2,5), "⅗": (3,5), "⅘": (4,5),
    "⅙": (1,6), "⅚": (5,6),
    "⅐": (1,7),
    "⅛": (1,8), "⅜": (3,8), "⅝": (5,8), "⅞": (7,8),
    "⅑": (1,9), "⅒": (1,10),
}

def contains_cjk(s: str) -> bool:
    return bool(_CJK_RE.search(s or ""))

def to_ascii_fractions(s: str) -> str:
    """將任意分數樣式轉為 ASCII：⅜→3/8，²⁄₉→2/9，⁄→/，以及將 ″/′ 轉為通用引號。"""
    if not s:
        return ""
    for ch, (num, den) in _VULGAR_TO_PAIR.items():
        s = s.replace(ch, f"{num}/{den}")
    s = s.translate(_SUP_TO_NORM).translate(_SUB_TO_NORM)
    s = s.replace("⁄", "/")
    s = s.replace("″", '"').replace("′", "'")
    return s

# =================== 一般欄位繪製 ===================
def _wrap_lines(pdf: FPDF, text: str, max_w: float, padding: float = 1.5) -> List[str]:
    text = "" if text is None else str(text)
    if not text:
        return [""]
    lines, line = [], ""
    limit = max_w - padding
    for ch in text:
        if pdf.get_string_width(line + ch) <= limit:
            line += ch
        else:
            lines.append(line); line = ch
    lines.append(line)
    return lines

def _draw_wrapped_cell(pdf: FPDF, w: float, h: float, text: str, line_h: float,
                       align: str = "C", font_size: int = 10, font_family: str = MAIN_FONT_NAME) -> None:
    x0, y0 = pdf.get_x(), pdf.get_y()
    pdf.cell(w, h, "", border=1)
    pdf.set_xy(x0, y0)
    pdf.set_font(font_family, '', font_size)
    lines = _wrap_lines(pdf, text or "", w)
    total_text_h = max(line_h * len(lines), line_h)
    y_text = y0 + max((h - total_text_h) / 2, 0)
    pdf.set_xy(x0, y_text)
    for i, ln in enumerate(lines):
        pdf.multi_cell(w, line_h, ln, border=0, align=align)
        if i < len(lines) - 1:
            pdf.set_x(x0)
    pdf.set_xy(x0 + w, y0)
def _fit_font_size(pdf: FPDF, text: str, max_w: float,
                   font_family: str, base_size: float, min_size: float = 8.0) -> float:
    """把 font size 從 base_size 往下調，直到 text 寬度塞進 max_w 或到 min_size。"""
    s = text or ""
    size = float(base_size)
    pdf.set_font(font_family, '', size)
    # 留一點左右內距
    limit = max_w - 1.5
    while size > min_size and pdf.get_string_width(s) > limit:
        size -= 0.5
        pdf.set_font(font_family, '', size)
    return size

def _draw_fit_cell(pdf: FPDF, w: float, h: float, text: str,
                   align: str = "C", base_size: int = 10, font_family: str = "NotoSansTC") -> None:
    """單行顯示、必要時縮小字級；過長仍超寬時用省略號。"""
    x0, y0 = pdf.get_x(), pdf.get_y()
    pdf.cell(w, h, "", border=1)

    s = "" if text is None else str(text)
    size = _fit_font_size(pdf, s, w, font_family, base_size)
    pdf.set_font(font_family, '', size)
    limit = w - 1.5
    text_w = pdf.get_string_width(s)

    # 若到最小字級仍太長，做省略（加 …）
    if text_w > limit:
        while s and pdf.get_string_width(s + "…") > limit:
            s = s[:-1]
        s = (s + "…") if s else s
        text_w = pdf.get_string_width(s)

    # 水平對齊
    if align == "R":
        x = x0 + max(w - text_w - 1.0, 0)
    elif align == "C":
        x = x0 + max((w - text_w) / 2, 0)
    else:
        x = x0 + 1.0

    # 垂直置中（FPDF 的 text 以基線為準，估一個舒適係數）
    y = y0 + (h + size * 0.35) / 2.0
    pdf.text(x, y, s)

    pdf.set_xy(x0 + w, y0)

# ====================== 行高估算 ======================
def _measure_row_height(pdf: FPDF, headers, col_widths, row, line_h):
    max_lines = 1
    for i, hname in enumerate(headers):
        fs = PER_COL_FONT_SIZE.get(hname, 10)
        cell_text = row[i] or ""
        if hname == "類別":
            # 類別：單行顯示（自動縮字），因此列高以 1 行為準
            # 仍照你的字型策略：含中文→用主字型；無中文→用分數字型
            if contains_cjk(cell_text):
                pdf.set_font(MAIN_FONT_NAME, '', fs)
            else:
                pdf.set_font(FRACTION_FONT_NAME, '', fs)
            lines = 1
        else:
            pdf.set_font(MAIN_FONT_NAME, '', fs)
            lines = len(_wrap_lines(pdf, cell_text, col_widths[i]))
        max_lines = max(max_lines, lines)
    return max(max_lines * line_h, HEADER_H)


# ======================== 主渲染 ========================
def _render_one_pdf_page(pdf: FPDF, customer: str, year: str, title_month: str,
                         rows: List[Dict[str, str]], overall_totals=None, is_last: bool = False) -> None:
    headers = ["日期", "訂單號碼", "類別", "顏色(組)", "數量(片)", "單價", "重量(kg)", "金額", "備註"]
    col_widths = COL_WIDTHS[:]

    pdf.add_page(); ensure_fonts(pdf)

    # 抬頭
    pdf.set_font(MAIN_FONT_NAME, '', 20)
    pdf.cell(0, 16, REPORT_TITLE_LEFT, ln=True, align='C')
    pdf.set_font(MAIN_FONT_NAME, '', 10)
    pdf.cell(0, 6, REPORT_ADDR, ln=True, align='C')
    pdf.cell(0, 6, REPORT_TEL, ln=True, align='C')
    pdf.ln(2)
    pdf.set_font(MAIN_FONT_NAME, '', 16)
    pdf.cell(0, 14, f"{year}年{title_month}月份工繳請款明細表", ln=True, align='C')
    pdf.set_font(MAIN_FONT_NAME, '', 12)
    pdf.cell(0, 10, f"客戶：{customer}", ln=True)
    pdf.ln(2)

    # 表頭
    pdf.set_font(MAIN_FONT_NAME, '', 10)
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], HEADER_H, h, border=1, align='C')
    pdf.ln()

    # 內容
    for rr in rows:
        row = [
            rr["date_str"], rr["order"], rr["type"], rr["color"], rr["quantity"],
            rr["unit_price"], rr["weight"], rr["amount"], rr["remark"]
        ]
        row_h = _measure_row_height(pdf, headers, col_widths, row, LINE_H)

        # 分頁控制
        bottom_limit = pdf.h - pdf.b_margin
        if pdf.get_y() + row_h > bottom_limit:
            pdf.add_page(); ensure_fonts(pdf)
            pdf.set_font(MAIN_FONT_NAME, '', 20)
            pdf.cell(0, 16, REPORT_TITLE_LEFT, ln=True, align='C')
            pdf.set_font(MAIN_FONT_NAME, '', 10)
            pdf.cell(0, 6, REPORT_ADDR, ln=True, align='C')
            pdf.cell(0, 6, REPORT_TEL, ln=True, align='C')
            pdf.ln(2)
            pdf.set_font(MAIN_FONT_NAME, '', 16)
            pdf.cell(0, 14, f"{year}年{title_month}月份工繳請款明細表", ln=True, align='C')
            pdf.set_font(MAIN_FONT_NAME, '', 12)
            pdf.cell(0, 10, f"客戶：{customer}", ln=True)
            pdf.ln(2)
            pdf.set_font(MAIN_FONT_NAME, '', 10)
            for i, h in enumerate(headers):
                pdf.cell(col_widths[i], HEADER_H, h, border=1, align='C')
            pdf.ln()

        # 畫列
        x0, y0 = pdf.get_x(), pdf.get_y()
        for i, text in enumerate(row):
            w = col_widths[i]
            fs = PER_COL_FONT_SIZE.get(headers[i], 10)
            if headers[i] == "類別":
                if contains_cjk(text or ""):
                    safe_text = to_ascii_fractions(text or "")
                    _draw_fit_cell(pdf, w, row_h, safe_text, align="C", base_size=fs, font_family=MAIN_FONT_NAME)
                else:
                    _draw_fit_cell(pdf, w, row_h, text or "", align="C", base_size=fs, font_family=FRACTION_FONT_NAME)
            else:
                align = "R" if headers[i] in ["數量(片)", "單價", "重量(kg)", "金額"] else "C"
                _draw_wrapped_cell(pdf, w, row_h, text or "", LINE_H, align=align, font_size=fs, font_family=MAIN_FONT_NAME)
        pdf.set_xy(x0, y0 + row_h)

    # 結尾合計（最後一頁）
    if is_last and overall_totals is not None:
        subtotal, tax, total = overall_totals
        pdf.set_font(MAIN_FONT_NAME, '', 12)
        pdf.multi_cell(0, 10, f"小計：{subtotal} 元\n稅(5%)：{tax} 元\n合計：{total} 元", align='R')
        pdf.set_x(10)
        pdf.multi_cell(190, 10, f"新臺幣：{number_to_chinese(total)}", align='R')

# ======================== 產出流程 ========================
def generate_pdf(customer: str, year: str, month: str, records: List[Dict[str, str]]) -> None:
    from tkinter import filedialog, messagebox

    # 計算金額
    for r in records:
        qty = int(r.get("quantity") or 0)
        price = float(r.get("unit_price") or 0)
        r["amount"] = str(round(qty * price))
    subtotal = sum(int(r["amount"]) for r in records)
    tax = round(subtotal * 0.05)
    total = subtotal + tax
    totals_tuple = (subtotal, tax, total)

    # 整理列資料（日期合併；金額加「元」）
    normalized_rows: List[Dict[str, str]] = []
    for r in records:
        date_str = f"{str(r.get('month','')).strip()}/{str(r.get('date','')).strip()}"
        w_raw = r.get("weight", "")
        w_text = "" if (w_raw is None or str(w_raw).strip() == "") else f"{float(w_raw):.2f}"
        normalized_rows.append({
            "date_str": date_str,
            "order": str(r.get("order", "")),
            "type":  str(r.get("type", "")),
            "color": str(r.get("color", "")),
            "quantity": str(r.get("quantity", "")),
            "unit_price": f"{float(r.get('unit_price', 0)):.2f}",
            "weight": w_text,
            "amount": str(r.get("amount", "")) + "元",
            "remark": str(r.get("remark", "")),
        })

    # 存檔對話框
    default_name = f"{year}年{month}月份_{customer}_工繳明細.pdf"
    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        initialdir=last_saved_dir,
        initialfile=default_name,
        filetypes=[("PDF files", "*.pdf")]
    )
    if not save_path:
        return
    set_last_saved_dir(os.path.dirname(save_path))

    # 分頁
    chunks = [normalized_rows[i:i + MAX_ROWS_PER_PDF] for i in range(0, len(normalized_rows), MAX_ROWS_PER_PDF)]
    num_pages = max(1, len(chunks))

    pdf = FPDF(format="A4", unit="mm")
    pdf.set_auto_page_break(auto=False)

    # 繪製
    try:
        for idx, rows_part in enumerate(chunks, start=1):
            _render_one_pdf_page(
                pdf=pdf,
                customer=customer,
                year=year,
                title_month=month,
                rows=rows_part,
                overall_totals=totals_tuple,
                is_last=(idx == num_pages)
            )
    except Exception as e:
        messagebox.showerror("產生失敗", f"產生 PDF 過程發生錯誤：\n{e}")
        return

    # 輸出
    try:
        pdf.output(save_path)
    except Exception as e:
        messagebox.showerror("輸出失敗", f"無法寫入 PDF：\n{e}")
        return

    try:
        messagebox.showinfo("成功", f"PDF 已輸出：\n{save_path}")
        os.startfile(save_path)
    except Exception:
        pass