import csv
import json
import os
import sys
from datetime import datetime
from dataclasses import dataclass, asdict
from typing import List

try:
    import openpyxl
except ImportError:
    openpyxl = None

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_FILE = os.path.join(BASE_DIR, "inventory.json")


@dataclass
class Record:
    time: str
    action: str
    qty: int
    model: str
    amount: float
    note: str
    total: int
    available: int
    borrowed: int


class Inventory:
    def __init__(self, data_file: str = DATA_FILE):
        self.data_file = data_file
        self.records: List[dict] = []
        self.total = 0
        self.available = 0
        self.borrowed = 0
        self._load()

    def _load(self):
        if os.path.exists(self.data_file):
            with open(self.data_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.total = data.get("total", 0)
            self.available = data.get("available", 0)
            self.borrowed = data.get("borrowed", 0)
            self.records = data.get("records", [])
        else:
            print("未找到数据文件，请先用 'init' 命令初始化库存。")

    def _save(self):
        with open(self.data_file, "w", encoding="utf-8") as f:
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

    def init(self, qty: int, model: str = "", amount: float = 0.0, note: str = ""):
        self.total = qty
        self.available = qty
        self.borrowed = 0
        self.records = []
        self._log("初始化", qty, model, amount, note or f"初始库存 {qty} 台")
        print(f"库存已初始化：总库存 {qty}，可用 {qty}，借出 0")

    def purchase(self, qty: int, model: str = "", amount: float = 0.0, note: str = ""):
        if qty <= 0:
            print("采购数量必须大于 0")
            return
        self.total += qty
        self.available += qty
        self._log("采购", qty, model, amount, note)
        print(f"采购入库 {qty} 台，当前可用：{self.available}，总库存：{self.total}")

    def return_device(self, qty: int, model: str = "", amount: float = 0.0, note: str = ""):
        if qty <= 0:
            print("退货数量必须大于 0")
            return
        if qty > self.available:
            print(f"退货失败：可用库存只有 {self.available} 台")
            return
        self.total -= qty
        self.available -= qty
        self._log("退货", qty, model, amount, note)
        print(f"退货出库 {qty} 台，当前可用：{self.available}，总库存：{self.total}")

    def borrow(self, qty: int, model: str = "", amount: float = 0.0, note: str = ""):
        if qty <= 0:
            print("借用数量必须大于 0")
            return
        if qty > self.available:
            print(f"借用失败：可用库存只有 {self.available} 台")
            return
        self.available -= qty
        self.borrowed += qty
        self._log("借用", qty, model, amount, note)
        print(f"借出 {qty} 台，当前可用：{self.available}，借出：{self.borrowed}")

    def give_back(self, qty: int, model: str = "", amount: float = 0.0, note: str = ""):
        if qty <= 0:
            print("归还数量必须大于 0")
            return
        if qty > self.borrowed:
            print(f"归还失败：当前借出只有 {self.borrowed} 台")
            return
        self.available += qty
        self.borrowed -= qty
        self._log("归还", qty, model, amount, note)
        print(f"归还入库 {qty} 台，当前可用：{self.available}，借出：{self.borrowed}")

    def status(self):
        print(f"\n======== 库存状态 ========")
        print(f"总库存  ：{self.total}")
        print(f"可用库存：{self.available}")
        print(f"借出    ：{self.borrowed}")
        print(f"========================\n")

    def history(self, limit: int = 20):
        if not self.records:
            print("暂无记录")
            return
        print(f"\n{'时间':<20} {'操作':<6} {'数量':>4} {'型号':<12} {'金额':>8} {'总库存':>5} {'可用':>5} {'借出':>5} 备注")
        print("-" * 100)
        for r in self.records[-limit:]:
            model = r.get("model", "")
            amount = r.get("amount", 0)
            print(f"{r['time']:<20} {r['action']:<6} {r['qty']:>4} {model:<12} {amount:>8.2f} {r['total']:>5} {r['available']:>5} {r['borrowed']:>5}  {r['note']}")
        print()

    def export_csv(self, path: str):
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["时间", "操作", "数量", "手机型号", "金额", "总库存", "可用", "借出", "备注"])
            for r in self.records:
                writer.writerow([
                    r["time"], r["action"], r["qty"],
                    r.get("model", ""), r.get("amount", 0),
                    r["total"], r["available"], r["borrowed"], r["note"]
                ])
        print(f"已导出 CSV：{path}")

    def export_excel(self, path: str):
        if openpyxl is None:
            print("未安装 openpyxl，无法导出 Excel。请使用 pip install openpyxl 安装。")
            return
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
        print(f"已导出 Excel：{path}")


def _parse_args(parts: list) -> tuple:
    """解析命令参数，提取 -m 型号 -p 金额，返回 (数量, 型号, 金额, 备注)"""
    qty = int(parts[1])
    model = ""
    amount = 0.0
    note_parts = []
    i = 2
    while i < len(parts):
        p = parts[i]
        if p == "-m" and i + 1 < len(parts):
            model = parts[i + 1]
            i += 2
        elif p == "-p" and i + 1 < len(parts):
            try:
                amount = float(parts[i + 1])
            except ValueError:
                pass
            i += 2
        else:
            note_parts.append(p)
            i += 1
    return qty, model, amount, " ".join(note_parts)


def print_help():
    print("""
命令格式：
  init <数量> [-m 型号] [-p 金额] [备注]      初始化库存（首次使用）
  purchase <数量> [-m 型号] [-p 金额] [备注]  采购入库
  return <数量> [-m 型号] [-p 金额] [备注]    退货出库（从可用库存中扣减）
  borrow <数量> [-m 型号] [-p 金额] [备注]    借用出库（可用减少，借出增加）
  back <数量> [-m 型号] [-p 金额] [备注]      归还入库（可用增加，借出减少）
  status                                      查看当前库存
  history [条数]                              查看最近操作记录（默认20条）
  export <文件名.csv>                         导出记录为 CSV 表格
  export --excel <文件名.xlsx>                导出记录为 Excel 表格
  help                                        显示帮助
  exit / quit                                 退出
""")


def main():
    inv = Inventory()
    print("设备库存管理系统（手机）")
    print("输入 help 查看命令，exit 退出")

    while True:
        try:
            cmd = input("> ").strip()
        except (EOFError, KeyboardInterrupt):
            break

        if not cmd:
            continue

        parts = cmd.split()
        action = parts[0].lower()

        if action in ("exit", "quit", "q"):
            break
        elif action == "help":
            print_help()
        elif action == "init":
            if len(parts) < 2:
                print("用法：init <数量> [-m 型号] [-p 金额] [备注]")
                continue
            try:
                qty, model, amount, note = _parse_args(parts)
                inv.init(qty, model, amount, note)
            except ValueError:
                print("数量必须是整数")
        elif action == "purchase":
            if len(parts) < 2:
                print("用法：purchase <数量> [-m 型号] [-p 金额] [备注]")
                continue
            try:
                qty, model, amount, note = _parse_args(parts)
                inv.purchase(qty, model, amount, note)
            except ValueError:
                print("数量必须是整数")
        elif action == "return":
            if len(parts) < 2:
                print("用法：return <数量> [-m 型号] [-p 金额] [备注]")
                continue
            try:
                qty, model, amount, note = _parse_args(parts)
                inv.return_device(qty, model, amount, note)
            except ValueError:
                print("数量必须是整数")
        elif action == "borrow":
            if len(parts) < 2:
                print("用法：borrow <数量> [-m 型号] [-p 金额] [备注]")
                continue
            try:
                qty, model, amount, note = _parse_args(parts)
                inv.borrow(qty, model, amount, note)
            except ValueError:
                print("数量必须是整数")
        elif action == "back":
            if len(parts) < 2:
                print("用法：back <数量> [-m 型号] [-p 金额] [备注]")
                continue
            try:
                qty, model, amount, note = _parse_args(parts)
                inv.give_back(qty, model, amount, note)
            except ValueError:
                print("数量必须是整数")
        elif action == "status":
            inv.status()
        elif action == "history":
            limit = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 20
            inv.history(limit)
        elif action == "export":
            if len(parts) < 2:
                print("用法：export <文件名.csv> 或 export --excel <文件名.xlsx>")
                continue
            if parts[1] == "--excel":
                if len(parts) < 3:
                    print("用法：export --excel <文件名.xlsx>")
                    continue
                inv.export_excel(parts[2])
            else:
                inv.export_csv(parts[1])
        else:
            print(f"未知命令：{action}，输入 help 查看帮助")

    print("已退出，数据已自动保存。")


if __name__ == "__main__":
    main()
