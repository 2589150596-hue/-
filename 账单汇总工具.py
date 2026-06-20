import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import datetime
import re
import os

try:
    import xlwt
except ImportError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "xlwt"])
    import xlwt


def parse_date(date_str):
    """把 '6月5日' 转成 datetime(2026, 6, 5)"""
    m = re.match(r'(\d{1,2})月(\d{1,2})日', date_str.strip())
    if m:
        month, day = int(m.group(1)), int(m.group(2))
        return datetime(2026, month, day)
    try:
        return datetime.strptime(date_str.strip(), '%Y-%m-%d')
    except ValueError:
        pass
    try:
        return datetime.strptime(date_str.strip(), '%Y/%m/%d')
    except ValueError:
        pass
    return None


def parse_line(line):
    """解析一行数据，支持空格/制表符分隔"""
    parts = re.split(r'\s+', line.strip())
    if len(parts) == 4:
        date_str, desc, amount_str, account = parts
    elif len(parts) > 4:
        date_str = parts[0]
        account = parts[-1]
        amount_str = parts[-2]
        desc = ' '.join(parts[1:-2])
    else:
        return None

    dt = parse_date(date_str)
    if dt is None:
        return None

    try:
        amount = float(amount_str)
    except ValueError:
        return None

    return {
        'date': dt,
        'desc': desc.strip(),
        'amount': amount,
        'account': account.strip()
    }


def split_remark(desc):
    """提取括号内容到备注，并删除摘要中的括号内容"""
    m = re.search(r'（(.+?)）', desc)
    if m:
        remark = m.group(1)
        desc_clean = re.sub(r'（.+?）', '', desc).strip()
        return desc_clean, remark
    return desc, ''


def generate_excel(records, save_path):
    """生成精简格式的 .xls 文件"""
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('支出明细')

    # 设置列宽
    col_widths = [12, 45, 20, 14, 12, 20]
    for i, w in enumerate(col_widths):
        sheet.col(i).width = 256 * w

    # 样式定义
    thin_border = xlwt.Borders()
    thin_border.left = xlwt.Borders.THIN
    thin_border.right = xlwt.Borders.THIN
    thin_border.top = xlwt.Borders.THIN
    thin_border.bottom = xlwt.Borders.THIN

    # 表头样式 - 浅灰背景+白字+加粗
    header_bg = xlwt.Pattern()
    header_bg.pattern = xlwt.Pattern.SOLID_PATTERN
    header_bg.pattern_fore_colour = xlwt.Style.colour_map['gray25']

    header_font = xlwt.Font()
    header_font.bold = True
    header_font.colour_index = xlwt.Style.colour_map['white']
    header_font.height = 220

    header_style = xlwt.XFStyle()
    header_style.font = header_font
    header_style.pattern = header_bg
    header_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
    header_style.alignment.vert = xlwt.Alignment.VERT_CENTER
    header_style.borders = thin_border

    # 数据行样式
    data_font = xlwt.Font()
    data_font.height = 200

    center_style = xlwt.XFStyle()
    center_style.font = data_font
    center_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
    center_style.alignment.vert = xlwt.Alignment.VERT_CENTER
    center_style.borders = thin_border

    left_style = xlwt.XFStyle()
    left_style.font = data_font
    left_style.alignment.horz = xlwt.Alignment.HORZ_LEFT
    left_style.alignment.vert = xlwt.Alignment.VERT_CENTER
    left_style.borders = thin_border

    amount_style = xlwt.XFStyle()
    amount_style.font = data_font
    amount_style.num_format_str = '#,##0.00'
    amount_style.alignment.horz = xlwt.Alignment.HORZ_RIGHT
    amount_style.alignment.vert = xlwt.Alignment.VERT_CENTER
    amount_style.borders = thin_border

    date_style = xlwt.XFStyle()
    date_style.font = data_font
    date_style.num_format_str = 'YYYY-MM-DD'
    date_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
    date_style.alignment.vert = xlwt.Alignment.VERT_CENTER
    date_style.borders = thin_border

    # 合计行样式 - 浅黄背景+加粗
    total_bg = xlwt.Pattern()
    total_bg.pattern = xlwt.Pattern.SOLID_PATTERN
    total_bg.pattern_fore_colour = xlwt.Style.colour_map['gold']

    total_font = xlwt.Font()
    total_font.bold = True
    total_font.height = 220

    total_style = xlwt.XFStyle()
    total_style.font = total_font
    total_style.pattern = total_bg
    total_style.alignment.horz = xlwt.Alignment.HORZ_CENTER
    total_style.alignment.vert = xlwt.Alignment.VERT_CENTER
    total_style.borders = thin_border

    total_amount_style = xlwt.XFStyle()
    total_amount_style.font = total_font
    total_amount_style.pattern = total_bg
    total_amount_style.num_format_str = '#,##0.00'
    total_amount_style.alignment.horz = xlwt.Alignment.HORZ_RIGHT
    total_amount_style.alignment.vert = xlwt.Alignment.VERT_CENTER
    total_amount_style.borders = thin_border

    # 表头
    headers = ['日期', '摘要', '备注', '金额', '费用归属', '支出账户']
    for col, h in enumerate(headers):
        sheet.write(0, col, h, header_style)

    sheet.row(0).height_mismatch = True
    sheet.row(0).height = 400

    # 写入数据
    row_idx = 1
    for rec in records:
        desc_clean, remark = split_remark(rec['desc'])

        sheet.row(row_idx).height_mismatch = True
        sheet.row(row_idx).height = 360

        sheet.write(row_idx, 0, rec['date'], date_style)
        sheet.write(row_idx, 1, desc_clean, left_style)
        sheet.write(row_idx, 2, remark, left_style)
        sheet.write(row_idx, 3, rec['amount'], amount_style)
        sheet.write(row_idx, 4, '公司', center_style)
        sheet.write(row_idx, 5, rec['account'], center_style)

        row_idx += 1

    # 合计行
    total = sum(r['amount'] for r in records)
    sheet.row(row_idx).height_mismatch = True
    sheet.row(row_idx).height = 400
    sheet.write(row_idx, 1, '合计', total_style)
    sheet.write(row_idx, 3, total, total_amount_style)
    sheet.write(row_idx, 0, '', total_style)
    sheet.write(row_idx, 2, '', total_style)
    sheet.write(row_idx, 4, '', total_style)
    sheet.write(row_idx, 5, '', total_style)

    book.save(save_path)
    return save_path


class App:
    def __init__(self, root):
        self.root = root
        self.root.title('账单汇总工具')
        self.root.geometry('900x600')

        tk.Label(root, text='把数据粘贴到下方（格式：日期  摘要  金额  账户）', font=('微软雅黑', 12)).pack(pady=5)
        tk.Label(root, text='例：6月5日    刷单    48000    清和汇支付宝', font=('微软雅黑', 10), fg='gray').pack()

        self.text = tk.Text(root, font=('Consolas', 11), wrap=tk.NONE)
        self.text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text='生成Excel', font=('微软雅黑', 12), bg='#4CAF50', fg='white',
                  width=15, command=self.on_generate).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text='清空', font=('微软雅黑', 12), width=10,
                  command=self.on_clear).pack(side=tk.LEFT, padx=5)

        self.status = tk.Label(root, text='就绪', font=('微软雅黑', 10), fg='gray', anchor='w')
        self.status.pack(fill=tk.X, padx=10, pady=5)

    def on_generate(self):
        raw = self.text.get('1.0', tk.END).strip()
        if not raw:
            messagebox.showwarning('提示', '请先粘贴数据')
            return

        lines = [l for l in raw.split('\n') if l.strip()]
        records = []
        errors = []
        for i, line in enumerate(lines, 1):
            rec = parse_line(line)
            if rec:
                records.append(rec)
            else:
                errors.append(f'第{i}行格式错误: {line}')

        if not records:
            messagebox.showerror('错误', '没有解析到有效数据，请检查格式')
            return

        if errors:
            err_msg = '\n'.join(errors[:5])
            if len(errors) > 5:
                err_msg += f'\n...还有{len(errors)-5}行错误'
            if not messagebox.askyesno('部分错误', f'以下行解析失败，是否继续生成？\n{err_msg}'):
                return

        save_path = filedialog.asksaveasfilename(
            defaultextension='.xls',
            filetypes=[('Excel 97-2003', '*.xls')],
            initialfile=f'垫付款明细_{datetime.now().strftime("%m%d")}.xls'
        )
        if not save_path:
            return

        try:
            generate_excel(records, save_path)
            self.status.config(text=f'已生成: {save_path}  共{len(records)}条记录')
            messagebox.showinfo('完成', f'汇总表已生成！\n共 {len(records)} 条记录')
        except Exception as e:
            messagebox.showerror('错误', f'生成失败: {e}')

    def on_clear(self):
        self.text.delete('1.0', tk.END)
        self.status.config(text='已清空')


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
