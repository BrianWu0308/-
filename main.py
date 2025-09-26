
## `main.py`
import tkinter as tk
from tkinter import ttk, messagebox
from pdf_generator import generate_pdf, set_last_saved_dir
import os
import re
from typing import Dict, List, Tuple

# ======================== 常數區 ========================
CUSTOMERS: Tuple[str, ...] = ("廣銘", "傑展", "儒鴻", "慧聚", "陞勇", "合一", "昌鴻", "其他")
ORDERNO_OPTIONS: Tuple[str, ...] = ("銷樣",)
CATEGORY_OPTIONS: Tuple[str, ...] = (
    "鍵盤", "領片", "袖口", "下擺", "門襟", "電腦領", "波浪領", "腰頭",
    "總針", "魚骨領", "雙面領", "其他"
)
REMARK_OPTIONS: Tuple[str, ...] = ("勾1次", "勾2次", "勾3次", "勾4次", "勾5次", "冷凍", "大尺寸", "立彬倒紗", "以下為冷凍", "其他")
MONTHS: Tuple[str, ...] = tuple(str(i) for i in range(1, 13))
DATES: Tuple[str, ...] = tuple(str(i) for i in range(1, 32))

DATA_COLUMNS: Tuple[str, ...] = (
    "月份", "日期", "訂單號碼", "類別", "顏色(組)", "數量(片)", "單價(元)", "重量(kg)", "備註"
)
DISPLAY_COLUMNS: Tuple[str, ...] = ("序號",) + DATA_COLUMNS

INPUT_WIDTHS: Dict[str, int] = {
    "月份": 6, "日期": 6, "訂單號碼": 12, "類別": 10, "顏色(組)": 16,
    "數量(片)": 8, "單價(元)": 8, "重量(kg)": 8, "備註": 16,
}

# 單字分數對照表（能對上就用它；對不上 fallback 成 ASCII num/den）
_VULGAR = {
    (1, 2): "½", (1, 3): "⅓", (2, 3): "⅔",
    (1, 4): "¼", (3, 4): "¾",
    (1, 5): "⅕", (2, 5): "⅖", (3, 5): "⅗", (4, 5): "⅘",
    (1, 6): "⅙", (5, 6): "⅚",
    (1, 7): "⅐",
    (1, 8): "⅛", (3, 8): "⅜", (5, 8): "⅝", (7, 8): "⅞",
    (1, 9): "⅑",
    (1, 10): "⅒",
}
_VULGAR_CHARS = "".join(_VULGAR.values())

# 預編譯正則
_RE_SPACE = re.compile(r"\s+")
_RE_CN_FRACTION = re.compile(r"(\d+)(?:分之)(\d+)")  # A分之B -> (B,A)
_RE_ASCII_FRACTION = re.compile(r"(\d+)\s*/\s*(\d+)") # B/A
_RE_MIXED_VULGAR = re.compile(rf"(\d+)又([{_VULGAR_CHARS}])")
_RE_MIXED_ASCII = re.compile(r"(\d+)又(\d+/\d+)")

# ===================== 轉換工具 =====================
def pretty_fraction_text(expr: str) -> str:
    """將『A分之B』『B/A』轉成單字分數或 ASCII 分數；處理混合數與常見符號。
    - 能對應者用單字分數（⅛、¼…）；否則 fallback 為 `B/A`
    - `乘/乘以/*/x/X` → `×`，`英吋` → `"`，`公分` → `cm`，小數點正規為 `.`
    - 混合數：整數 + 單字分數相連（16⅛），整數 + ASCII 分數有空格（16 1/16）
    """
    if not expr:
        return expr

    t = _RE_SPACE.sub("", expr)
    t = (
        t.replace("乘以", "×").replace("乘", "×")
         .replace("*", "×").replace("x", "×").replace("X", "×")
         .replace("．", ".").replace("。", ".")
         .replace("英吋", '"')
         .replace("公分", "cm")
    )

    def _repl_cn(m: re.Match) -> str:
        den, num = m.group(1), m.group(2)
        try:
            key = (int(num), int(den))
            return _VULGAR.get(key, f"{num}/{den}")
        except Exception:
            return f"{num}/{den}"

    def _repl_ascii(m: re.Match) -> str:
        num, den = m.group(1), m.group(2)
        try:
            key = (int(num), int(den))
            return _VULGAR.get(key, f"{num}/{den}")
        except Exception:
            return f"{num}/{den}"

    t = _RE_CN_FRACTION.sub(_repl_cn, t)
    t = _RE_ASCII_FRACTION.sub(_repl_ascii, t)
    t = _RE_MIXED_VULGAR.sub(r"\1\2", t)
    t = _RE_MIXED_ASCII.sub(r"\1 \2", t)
    return t

# ======================= 主 App =======================
class WageApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("工繳明細自動產生器")
        root.geometry("1400x900")

        # 全域字體
        default_font = ("Microsoft JhengHei", 16)
        root.option_add("*Font", default_font)
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Microsoft JhengHei", 16))
        style.configure("Treeview", font=("Microsoft JhengHei", 16), rowheight=40)

        self.last_saved_dir = os.getcwd()
        self.inputs: Dict[str, tk.Widget] = {}
        self.color_mode = tk.StringVar(value="輸入數量")

        self._build_top()
        self._build_table()
        self._build_inputs()
        self._build_buttons()
        self._bind_shortcuts()

    # ---------- UI Blocks ----------
    def _build_top(self) -> None:
        frame_top = tk.Frame(self.root)
        frame_top.pack(pady=5)

        tk.Label(frame_top, text="客戶名稱").grid(row=0, column=0, padx=4)
        tk.Label(frame_top, text="年份（民國）").grid(row=0, column=2, padx=4)
        tk.Label(frame_top, text="標題月份").grid(row=0, column=4, padx=4)

        self.customer_entry = ttk.Combobox(frame_top, width=18, state="normal", values=CUSTOMERS)
        self.year_entry = tk.Entry(frame_top, width=10)
        self.month_combobox = ttk.Combobox(frame_top, values=MONTHS, width=5, state="readonly")

        self.customer_entry.grid(row=0, column=1)
        self.year_entry.grid(row=0, column=3)
        self.month_combobox.grid(row=0, column=5)
        self.month_combobox.set("1")

    def _build_table(self) -> None:
        table_frame = tk.Frame(self.root)
        table_frame.pack(padx=5, pady=5, fill="both", expand=True)

        self.table = ttk.Treeview(table_frame, columns=DISPLAY_COLUMNS, show='headings')
        ysb = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        xsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)

        self.table.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        for col in DISPLAY_COLUMNS:
            self.table.heading(col, text=col)
            self.table.column(col, width=(70 if col == "序號" else 130), anchor="center")

        # 滾輪
        def _on_mousewheel(event):
            if event.num == 4:
                delta = 1
            elif event.num == 5:
                delta = -1
            else:
                delta = int(event.delta / 120)
            self.table.yview_scroll(-delta, "units")
        self.table.bind("<MouseWheel>", _on_mousewheel)
        self.table.bind("<Button-4>", _on_mousewheel)
        self.table.bind("<Button-5>", _on_mousewheel)
        self.table.bind("<Shift-MouseWheel>", lambda e: self.table.xview_scroll(int(-e.delta/120), "units"))

    def _build_inputs(self) -> None:
        frame_input = tk.Frame(self.root)
        frame_input.pack(pady=6)

        tk.Label(frame_input, text="顏色輸入模式").grid(row=0, column=0, padx=4, sticky="w")
        mode_cb = ttk.Combobox(frame_input, textvariable=self.color_mode, values=("輸入數量", "輸入顏色"), width=10, state="readonly")
        mode_cb.grid(row=1, column=0, padx=4, sticky="w")

        # 欄位
        for i, col in enumerate(DATA_COLUMNS, start=1):
            tk.Label(frame_input, text=col).grid(row=0, column=i, padx=4, sticky="n", pady=(0, 2))
            w = INPUT_WIDTHS.get(col, 12)

            if col == "類別":
                cb = ttk.Combobox(frame_input, values=CATEGORY_OPTIONS, width=w, state="normal")
                cb.grid(row=1, column=i, padx=4, sticky="w")
                self.inputs[col] = cb

                def _on_cat_selected(_evt=None, this_cb=cb):
                    if this_cb.get() == "鍵盤":
                        this_cb.set("")
                        self.open_category_keypad(target_cb=this_cb)
                cb.bind("<<ComboboxSelected>>", _on_cat_selected)

            elif col == "月份":
                cb = ttk.Combobox(frame_input, values=MONTHS, width=w, state="readonly")
                cb.grid(row=1, column=i, padx=4, sticky="w")
                cb.set(self.month_combobox.get())
                self.inputs[col] = cb

            elif col == "日期":
                cb = ttk.Combobox(frame_input, values=DATES, width=w)
                cb.grid(row=1, column=i, padx=4, sticky="w")
                self.inputs[col] = cb

            elif col == "訂單號碼":
                cb = ttk.Combobox(frame_input, width=w, state="normal", values=ORDERNO_OPTIONS)
                cb.grid(row=1, column=i, padx=4, sticky="w")
                self.inputs[col] = cb

            elif col == "顏色(組)":
                e = tk.Entry(frame_input, width=w)
                e.grid(row=1, column=i, padx=4, sticky="w")
                self.inputs[col] = e

                self.color_hint = tk.Label(frame_input, fg="#666666", text="（模式：輸入數量）")
                self.color_hint.grid(row=2, column=i, padx=4, pady=(2, 0), sticky="w")

                def on_mode_changed(_evt=None):
                    self.color_hint.config(text=f"（模式：{self.color_mode.get()}）")
                mode_cb.bind("<<ComboboxSelected>>", on_mode_changed)
                on_mode_changed()

            elif col == "備註":
                cb = ttk.Combobox(frame_input, width=w, state="normal", values=REMARK_OPTIONS)
                cb.grid(row=1, column=i, padx=4, sticky="w")
                self.inputs[col] = cb

            else:
                e = tk.Entry(frame_input, width=w)
                e.grid(row=1, column=i, padx=4, sticky="w")
                self.inputs[col] = e

    def _build_buttons(self) -> None:
        frame_button = tk.Frame(self.root)
        frame_button.pack(pady=8)
        tk.Button(frame_button, text="新增資料列", command=self.add_row).pack(side='left', padx=10)
        tk.Button(frame_button, text="複製到輸入欄", command=self.copy_to_inputs).pack(side='left', padx=10)
        tk.Button(frame_button, text="刪除選取列", command=self.delete_row, bg='red', fg='white').pack(side='left', padx=10)
        tk.Button(frame_button, text="產生 PDF", command=self.export_pdf, bg='green', fg='white').pack(side='left', padx=10)

    def _bind_shortcuts(self) -> None:
        self.root.bind('<Return>', lambda event: self.add_row())
        self.root.bind('<Delete>', lambda event: self.delete_row())

    # ---------- 類別鍵盤 ----------
    def open_category_keypad(self, target_cb: ttk.Combobox) -> None:
        top = tk.Toplevel(self.root)
        top.title("類別輸入鍵盤")
        top.grab_set()

        init_val = target_cb.get()
        expr_var = tk.StringVar(value="" if init_val == "鍵盤" else init_val)

        row0 = tk.Frame(top); row0.pack(padx=8, pady=6, fill="x")
        tk.Label(row0, text="輸入：").pack(side="left")
        entry = tk.Entry(row0, textvariable=expr_var, width=40)
        entry.pack(side="left", fill="x", expand=True)

        row1 = tk.Frame(top); row1.pack(padx=8, pady=(0,8), fill="x")
        tk.Label(row1, text="預覽：").pack(side="left")
        preview = tk.Label(row1, text=pretty_fraction_text(expr_var.get()), fg="#444")
        preview.pack(side="left", fill="x", expand=True)

        expr_var.trace_add("write", lambda *_: preview.config(text=pretty_fraction_text(expr_var.get())))

        keys = [
            ["7","8","9","乘","英吋"],
            ["4","5","6","分之","公分"],
            ["1","2","3","又","小數點"],
            ["0","←","清除","空白","確定"]
        ]

        def put(tok: str) -> None:
            if tok == "小數點": tok = "."
            if tok == "乘": tok = "×"
            if tok == "英吋": tok = '"'
            if tok == "公分": tok = "cm"
            if tok == "空白": tok = " "
            if tok == "清除":
                expr_var.set(""); return
            if tok == "←":
                s = expr_var.get()
                expr_var.set(s[:-1] if s else s); return
            if tok == "確定":
                nice = pretty_fraction_text(expr_var.get())
                target_cb.delete(0, tk.END)
                target_cb.insert(0, nice)
                top.destroy(); return
            pos = entry.index(tk.INSERT)
            s = expr_var.get()
            expr_var.set(s[:pos] + tok + s[pos:])
            entry.icursor(pos + len(tok))
            entry.focus_set()

        grid = tk.Frame(top); grid.pack(padx=8, pady=8)
        for r, row in enumerate(keys):
            for c, label in enumerate(row):
                ttk.Button(grid, text=label, width=8, command=lambda t=label: put(t)).grid(row=r, column=c, padx=4, pady=4)
        try:
            top.option_add("*Font", ("Microsoft JhengHei", 14))
        except Exception:
            pass
        entry.focus_set()

    # ---------- 表格操作 ----------
    def add_row(self) -> None:
        values = [self.inputs[col].get().strip() for col in DATA_COLUMNS]

        if not values[DATA_COLUMNS.index("月份")]:
            messagebox.showwarning("缺少資料", "請選擇【月份】。"); return
        if not values[DATA_COLUMNS.index("日期")] or not values[DATA_COLUMNS.index("訂單號碼")]:
            messagebox.showwarning("缺少資料", "請至少填寫【日期】與【訂單號碼】。"); return

        # 顏色處理：兩種模式
        color_raw = self.inputs["顏色(組)"].get().strip()
        if self.color_mode.get() == "輸入數量":
            try:
                n = int(color_raw)
                if n < 0: raise ValueError
                color_display = str(n)
            except ValueError:
                messagebox.showwarning("格式錯誤", "顏色輸入模式為【輸入數量】時，請輸入整數（例：3）。"); return
        else:
            parts = [p for p in re.split(r"[,\s、]+", color_raw) if p]
            if not parts:
                messagebox.showwarning("格式錯誤", "請輸入至少一個顏色名稱（例：黑, 白, 紅）。"); return
            color_display = "、".join(parts)
        values[DATA_COLUMNS.index("顏色(組)")] = color_display

        # 數字欄驗證（可空白）
        def _num_ok(val: str, caster) -> bool:
            if val == "": return True
            try:
                caster(val); return True
            except Exception:
                return False
        if not _num_ok(values[DATA_COLUMNS.index("數量(片)")], int):
            messagebox.showwarning("格式錯誤", "【數量(片)】需為整數。"); return
        if not _num_ok(values[DATA_COLUMNS.index("單價(元)")], float):
            messagebox.showwarning("格式錯誤", "【單價(元)】需為數字。"); return
        if not _num_ok(values[DATA_COLUMNS.index("重量(kg)")], float):
            messagebox.showwarning("格式錯誤", "【重量(kg)】需為數字。"); return

        # 插入表格
        next_idx = len(self.table.get_children()) + 1
        display_values = [str(next_idx)] + values
        self.table.insert('', 'end', values=display_values)

        # 清空輸入欄，月份回填標題月份
        for widget in self.inputs.values():
            try:
                widget.delete(0, tk.END)
            except Exception:
                pass
        self.inputs["月份"].set(self.month_combobox.get())

    def delete_row(self) -> None:
        selected = self.table.selection()
        if not selected:
            messagebox.showwarning("未選取", "請先選取要刪除的資料列"); return

        lines: List[str] = []
        for iid in selected:
            vals = self.table.item(iid)['values']
            row = {"序號": vals[0]}
            for i, col in enumerate(DATA_COLUMNS, start=1):
                row[col] = vals[i]
            lines.append(
                f"#{row['序號']}  月:{row.get('月份','')} 日:{row.get('日期','')}  "
                f"訂單:{row.get('訂單號碼','')}  類別:{row.get('類別','')}  "
                f"顏色:{row.get('顏色(組)','')}  數量:{row.get('數量(片)','')}  "
                f"單價:{row.get('單價(元)','')}  重量:{row.get('重量(kg)','')}  "
                f"備註:{row.get('備註','')}"
            )
        preview_text = "\n".join(lines[:10]) + (f"\n...（共 {len(lines)} 筆）" if len(lines) > 10 else "")
        if not messagebox.askyesno("確認刪除", f"即將刪除以下 {len(selected)} 筆資料：\n\n{preview_text}\n\n是否確定刪除？"):
            return
        for iid in selected:
            self.table.delete(iid)
        for i, iid in enumerate(self.table.get_children(), start=1):
            self.table.set(iid, "序號", str(i))
    def _set_widget_text(self, widget: tk.Widget, text: str) -> None:
        """通用：把文字塞進 Entry / Combobox。"""
        s = "" if text is None else str(text)
        if isinstance(widget, ttk.Combobox):
            widget.set(s)
        else:
            try:
                widget.delete(0, tk.END)
                widget.insert(0, s)
            except Exception:
                pass

    def copy_to_inputs(self) -> None:
        """將目前表格所選『單一列』的資料填回下方輸入欄。"""
        sel = self.table.selection()
        if not sel:
            messagebox.showwarning("未選取", "請先在表格選取要複製的一列")
            return
        if len(sel) > 1:
            messagebox.showwarning("一次僅支援一列", "請只選取一列再執行『複製到輸入欄』")
            return

        iid = sel[0]
        values = self.table.item(iid)["values"]  # ["序號", 月份, 日期, 訂單號碼, 類別, 顏色(組), 數量(片), 單價(元), 重量(kg), 備註]
        if not values or len(values) < 1 + len(self.inputs):
            messagebox.showerror("資料錯誤", "選取列的資料格式不完整，無法複製。")
            return

        # 跳過 index=0 的「序號」，其餘依 DATA_COLUMNS 順序寫回
        data_values = values[1:]
        for col_name, col_value in zip(DATA_COLUMNS, data_values):
            widget = self.inputs.get(col_name)
            if widget is not None:
                self._set_widget_text(widget, col_value)

        # 讓使用者一眼看到可編輯：把游標移到第一個欄位
        first = self.inputs.get("月份") or self.inputs.get(DATA_COLUMNS[0])
        try:
            first.focus_set()
        except Exception:
            pass    
    def export_pdf(self) -> None:
        customer = self.customer_entry.get().strip()
        year = self.year_entry.get().strip()
        month = self.month_combobox.get().strip()
        if not (customer and year and month):
            messagebox.showwarning("錯誤", "請填寫客戶名稱、年份與標題月份"); return

        records = []
        for row in self.table.get_children():
            data = self.table.item(row)['values']
            data_no_idx = data[1:]
            try:
                rec = {
                    "month":      data_no_idx[DATA_COLUMNS.index("月份")],
                    "date":       data_no_idx[DATA_COLUMNS.index("日期")],
                    "order":      data_no_idx[DATA_COLUMNS.index("訂單號碼")],
                    "type":       data_no_idx[DATA_COLUMNS.index("類別")],
                    "color":      data_no_idx[DATA_COLUMNS.index("顏色(組)")],
                    "quantity":   int(data_no_idx[DATA_COLUMNS.index("數量(片)")] or 0),
                    "unit_price": float(data_no_idx[DATA_COLUMNS.index("單價(元)")] or 0),
                    "weight":     data_no_idx[DATA_COLUMNS.index("重量(kg)")],
                    "remark":     data_no_idx[DATA_COLUMNS.index("備註")],
                }
                records.append(rec)
            except Exception:
                messagebox.showerror("格式錯誤", f"有資料無法解析：{data}\n請檢查【數量/單價/重量】是否為數字。")
                return

        if not records:
            messagebox.showwarning("沒有資料", "請先新增至少一筆資料再產生 PDF。"); return

        set_last_saved_dir(self.last_saved_dir)
        generate_pdf(customer, year, month, records)


if __name__ == '__main__':
    root = tk.Tk()
    app = WageApp(root)
    root.mainloop()
