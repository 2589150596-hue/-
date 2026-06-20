import csv
import json
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

try:
    import openpyxl
except ImportError:
    openpyxl = None

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_FILE = os.path.join(BASE_DIR, "inventory.json")


class InventoryApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("设备库存管理 - 手机")
        self.root.geometry("900x620")
        self.root.resizable(False, False)

        # 数据
        self.total = 0
        self.available = 0
        self.borrowed = 0
        self.records = []
        self._load()

        # 样式
        self.style = ttk.Style()
        self.style.configure("Title.TLabel", font=("Microsoft YaHei", 14, "bold"))
        self.style.configure("Card.TFrame", background="#f8f9fa")
        self.style.configure("CardLabel.TLabel", background="#f8f9fa", font=("Microsoft YaHei", 10))
        self.style.configure("CardValue.TLabel", background="#f8f9fa", font=("Microsoft YaHei", 18, "bold"))

        self._build_ui()
        self._refresh_status()
        self._refresh_table()

    def _load(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.total = data.get("total", 0)
                self.available = data.get("available", 0)
                self.borrowed = data.get("borrowed", 0)
                self.records = data.get("records", [])
            except Exception:
                pass

    def _save(self):
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump({
                "total": self.total,
                "available": self.available,
                "borrowed": self.borrowed,
                "records": self.records
            }, f, ensure_ascii=False, indent=2)

    def _log(self, action: str, qty: int, model: str, amount: float, note: str):
        rec = {
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "action": action,
            "qty": qty,
            "model": model,
            "amount": amount,
            "note": note,
            "total": self.total,
            "available": self.available,
            "borrowed": self.borrowed
        }
        self.records.append(rec)
        self._save()

    def _build_ui(self):
        # 顶部标题
        title = ttk.Label(self.root, text="设备库存管理系统", style="Title.TLabel")
        title.pack(pady=(16, 8))

        # 库存卡片区
        card_frame = ttk.Frame(self.root)
        card_frame.pack(pady=8)

        self.card_total = self._create_card(card_frame, "总库存", "#0d6efd", 0)
        self.card_avail = self._create_card(card_frame, "可用库存", "#198754", 1)
        self.card_borrow = self._create_card(card_frame, "借出", "#fd7e14", 2)

        # 操作区
        op_frame = ttk.LabelFrame(self.root, text="操作", padding=12)
        op_frame.pack(fill="x", padx=20, pady=10)

        ttk.Label(op_frame, text="操作类型：").grid(row=0, column=0, sticky="w")
        self.action_var = tk.StringVar(value="采购")
        action_combo = ttk.Combobox(op_frame, textvariable=self.action_var, values=["采购", "退货", "借用", "归还"], state="readonly", width=12)
        action_combo.grid(row=0, column=1, sticky="w", padx=(0, 16))

        ttk.Label(op_frame, text="数量：").grid(row=0, column=2, sticky="w")
        self.qty_var = tk.StringVar()
        ttk.Entry(op_frame, textvariable=self.qty_var, width=10).grid(row=0, column=3, sticky="w", padx=(0, 16))

        ttk.Label(op_frame, text="手机型号：").grid(row=0, column=4, sticky="w")
        self.model_var = tk.StringVar()
        ttk.Entry(op_frame, textvariable=self.model_var, width=16).grid(row=0, column=5, sticky="w", padx=(0, 16))

        ttk.Label(op_frame, text="金额(元)：").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.amount_var = tk.StringVar()
        ttk.Entry(op_frame, textvariable=self.amount_var, width=12).grid(row=1, column=1, sticky="w", padx=(0, 16), pady=(8, 0))

        ttk.Label(op_frame, text="备注：").grid(row=1, column=2, sticky="w", pady=(8, 0))
        self.note_var = tk.StringVar()
        ttk.Entry(op_frame, textvariable=self.note_var, width=28).grid(row=1, column=3, columnspan=3, sticky="w", pady=(8, 0))

        btn_frame = ttk.Frame(op_frame)
        btn_frame.grid(row=2, column=0, columnspan=6, pady=(12, 0))

        ttk.Button(btn_frame, text="执行", command=self._on_execute).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="初始化库存", command=self._on_init).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="撤销最后一条", command=self._on_undo).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="导出 CSV", command=self._export_csv).pack(side="left", padx=4)
        if openpyxl:
            ttk.Button(btn_frame, text="导出 Excel", command=self._export_excel).pack(side="left", padx=4)

        # 记录表格
        table_frame = ttk.LabelFrame(self.root, text="操作记录", padding=(8, 8))
        table_frame.pack(fill="both", expand=True, padx=20, pady=(4, 16))

        cols = ("time", "action", "qty", "model", "amount", "total", "available", "borrowed", "note")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=12)
        self.tree.heading("time", text="时间")
        self.tree.heading("action", text="操作")
        self.tree.heading("qty", text="数量")
        self.tree.heading("model", text="手机型号")
        self.tree.heading("amount", text="金额")
        self.tree.heading("total", text="总库存")
        self.tree.heading("available", text="可用")
        self.tree.heading("borrowed", text="借出")
        self.tree.heading("note", text="备注")

        self.tree.column("time", width=130, anchor="center")
        self.tree.column("action", width=60, anchor="center")
        self.tree.column("qty", width=50, anchor="center")
        self.tree.column("model", width=100, anchor="center")
        self.tree.column("amount", width=80, anchor="center")
        self.tree.column("total", width=60, anchor="center")
        self.tree.column("available", width=50, anchor="center")
        self.tree.column("borrowed", width=50, anchor="center")
        self.tree.column("note", width=180)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _create_card(self, parent, label, color, col):
        card = tk.Frame(parent, bg="#f8f9fa", bd=1, relief="solid", highlightbackground="#dee2e6", highlightthickness=1)
        card.grid(row=0, column=col, padx=8, pady=4, ipadx=20, ipady=10)

        lbl = tk.Label(card, text=label, bg="#f8f9fa", fg="#6c757d", font=("Microsoft YaHei", 10))
        lbl.pack()

        val = tk.Label(card, text="0", bg="#f8f9fa", fg=color, font=("Microsoft YaHei", 20, "bold"))
        val.pack()
        return val

    def _refresh_status(self):
        self.card_total.config(text=str(self.total))
        self.card_avail.config(text=str(self.available))
        self.card_borrow.config(text=str(self.borrowed))

    def _refresh_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for r in reversed(self.records):
            self.tree.insert("", "end", values=(
                r["time"], r["action"], r["qty"],
                r.get("model", ""), r.get("amount", 0),
                r["total"], r["available"], r["borrowed"], r["note"]
            ))

    def _get_qty(self) -> int | None:
        try:
            q = int(self.qty_var.get().strip())
            if q <= 0:
                messagebox.showwarning("输入错误", "数量必须大于 0")
                return None
            return q
        except ValueError:
            messagebox.showwarning("输入错误", "数量必须是整数")
            return None

    def _get_amount(self) -> float:
        try:
            return float(self.amount_var.get().strip())
        except ValueError:
            return 0.0

    def _on_execute(self):
        action = self.action_var.get()
        qty = self._get_qty()
        if qty is None:
            return
        model = self.model_var.get().strip()
        amount = self._get_amount()
        note = self.note_var.get().strip()

        if action == "采购":
            self.total += qty
            self.available += qty
            self._log("采购", qty, model, amount, note)
        elif action == "退货":
            if qty > self.available:
                messagebox.showerror("库存不足", f"可用库存只有 {self.available} 台")
                return
            self.total -= qty
            self.available -= qty
            self._log("退货", qty, model, amount, note)
        elif action == "借用":
            if qty > self.available:
                messagebox.showerror("库存不足", f"可用库存只有 {self.available} 台")
                return
            self.available -= qty
            self.borrowed += qty
            self._log("借用", qty, model, amount, note)
        elif action == "归还":
            if qty > self.borrowed:
                messagebox.showerror("数量错误", f"当前借出只有 {self.borrowed} 台")
                return
            self.available += qty
            self.borrowed -= qty
            self._log("归还", qty, model, amount, note)

        self.qty_var.set("")
        self.model_var.set("")
        self.amount_var.set("")
        self.note_var.set("")
        self._refresh_status()
        self._refresh_table()

    def _on_init(self):
        ans = messagebox.askyesno("初始化", "初始化会清空所有记录并重新设定库存，确定吗？")
        if not ans:
            return
        qty = self._get_qty()
        if qty is None:
            return
        model = self.model_var.get().strip()
        amount = self._get_amount()
        self.total = qty
        self.available = qty
        self.borrowed = 0
        self.records = []
        self._log("初始化", qty, model, amount, f"初始库存 {qty} 台")
        self.qty_var.set("")
        self.model_var.set("")
        self.amount_var.set("")
        self.note_var.set("")
        self._refresh_status()
        self._refresh_table()
        messagebox.showinfo("完成", f"库存已初始化为 {qty} 台")

    def _on_undo(self):
        if not self.records:
            messagebox.showinfo("提示", "没有可撤销的记录")
            return
        ans = messagebox.askyesno("撤销", "撤销最后一条操作记录？")
        if not ans:
            return

        rec = self.records.pop()
        action = rec["action"]
        qty = rec["qty"]

        # 回滚状态：恢复到记录之前的状态（记录里存的是操作后的状态，所以直接用记录里的前一项状态更稳妥）
        # 简单做法：用记录里保存的状态值直接回退
        if self.records:
            prev = self.records[-1]
            self.total = prev["total"]
            self.available = prev["available"]
            self.borrowed = prev["borrowed"]
        else:
            # 只有一条记录，撤销后归零
            self.total = 0
            self.available = 0
            self.borrowed = 0

        self._save()
        self._refresh_status()
        self._refresh_table()
        messagebox.showinfo("完成", f"已撤销：{rec['time']} {action} {qty} 台")

    def _export_csv(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")],
            initialfile=f"inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["时间", "操作", "数量", "手机型号", "金额", "总库存", "可用", "借出", "备注"])
                for r in self.records:
                    writer.writerow([
                        r["time"], r["action"], r["qty"],
                        r.get("model", ""), r.get("amount", 0),
                        r["total"], r["available"], r["borrowed"], r["note"]
                    ])
            messagebox.showinfo("导出成功", f"已保存到：{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _export_excel(self):
        if openpyxl is None:
            messagebox.showwarning("提示", "未安装 openpyxl，无法导出 Excel")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
            initialfile=f"inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not path:
            return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "库存记录"
            headers = ["时间", "操作", "数量", "手机型号", "金额", "总库存", "可用", "借出", "备注"]
            ws.append(headers)
            for r in self.records:
                ws.append([
                    r["time"], r["action"], r["qty"],
                    r.get("model", ""), r.get("amount", 0),
                    r["total"], r["available"], r["borrowed"], r["note"]
                ])
            wb.save(path)
            messagebox.showinfo("导出成功", f"已保存到：{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))


def main():
    root = tk.Tk()
    app = InventoryApp(root)

    if app.total == 0 and app.available == 0 and not app.records:
        messagebox.showinfo("首次使用", "当前库存为 0，请点击【初始化库存】设置初始数量（如 222）")

    root.mainloop()


if __name__ == "__main__":
    main()
