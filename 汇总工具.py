import pandas as pd
import os
from datetime import datetime, timedelta
import glob

def find_column(columns, keywords):
    """根据关键词列表匹配列名"""
    for col in columns:
        col_str = str(col).strip()
        for kw in keywords:
            if kw in col_str:
                return col_str
    return None

def process_folder(folder_path, output_path=None, start_time_str="08:00"):
    """
    汇总文件夹内所有Excel表格，按店铺分组计算每单时间，按时间顺序排列
    """
    # 查找所有Excel文件
    patterns = [os.path.join(folder_path, "*.xlsx"), os.path.join(folder_path, "*.xls")]
    files = []
    for p in patterns:
        files.extend(glob.glob(p))

    if not files:
        print(f"未在 {folder_path} 找到Excel文件")
        return

    print(f"找到 {len(files)} 个文件: {files}")

    # 读取并汇总所有文件
    all_data = []
    for f in files:
        try:
            df = pd.read_excel(f)
            # 跳过完全空的行
            df = df.dropna(how='all')
            if len(df) == 0:
                continue
            # 添加来源文件列
            df['来源文件'] = os.path.basename(f)
            all_data.append(df)
            print(f"  读取 {os.path.basename(f)}: {len(df)} 行")
        except Exception as e:
            print(f"  读取 {os.path.basename(f)} 失败: {e}")

    if not all_data:
        print("没有成功读取任何数据")
        return

    # 合并所有数据
    combined = pd.concat(all_data, ignore_index=True)
    print(f"\n合并后总行数: {len(combined)}")

    # 识别店铺名列
    shop_col = None
    for col in combined.columns:
        col_str = str(col)
        if '店铺名称' in col_str or '客户' in col_str and '店铺' in col_str:
            # 优先找有实际数据的列
            non_null = combined[col].notna().sum()
            if non_null > 0:
                shop_col = col
                break

    if shop_col is None:
        #  fallback: 找包含最多非空文本的列
        best_col = None
        best_count = 0
        for col in combined.columns:
            if str(col).startswith('Unnamed'):
                continue
            count = combined[col].apply(lambda x: isinstance(x, str) and len(x) > 3).sum()
            if count > best_count:
                best_count = count
                best_col = col
        shop_col = best_col

    print(f"店铺列识别为: {shop_col}")

    # 识别单量列
    qty_col = None
    for col in combined.columns:
        col_str = str(col)
        if '补单数量' in col_str or '单量' in col_str or '数量' in col_str:
            qty_col = col
            break

    if qty_col is None:
        print("未找到单量列（补单数量/单量/数量），无法继续")
        return

    print(f"单量列识别为: {qty_col}")

    # 过滤有效数据行：店铺名不为空，单量大于0
    valid_mask = combined[shop_col].notna() & combined[qty_col].notna()
    # 进一步过滤：单量必须是数字且>0
    try:
        qty_numeric = pd.to_numeric(combined[qty_col], errors='coerce')
        valid_mask = valid_mask & (qty_numeric > 0)
    except:
        pass

    data = combined[valid_mask].copy()
    print(f"有效数据行数: {len(data)}")

    if len(data) == 0:
        print("没有有效数据")
        return

    # 确保单量是数字
    data[qty_col] = pd.to_numeric(data[qty_col], errors='coerce')
    data = data[data[qty_col] > 0].copy()

    # 清理店铺名称中的换行符和多余空格
    data[shop_col] = data[shop_col].astype(str).str.replace(r'[\n\r]+', '', regex=True).str.strip()

    # 按店铺分组，计算每个店铺的总单量和每单间隔
    shop_stats = data.groupby(shop_col)[qty_col].sum().reset_index()
    shop_stats.columns = [shop_col, '店铺总单量']
    shop_stats['每单间隔(分钟)'] = 600 / shop_stats['店铺总单量']

    print("\n店铺统计:")
    for _, row in shop_stats.iterrows():
        print(f"  {row[shop_col]}: 总单量={row['店铺总单量']}, 间隔={row['每单间隔(分钟)']:.1f}分钟")

    # 解析开始时间
    start_h, start_m = map(int, start_time_str.split(':'))
    base_time = datetime(2026, 1, 1, start_h, start_m)

    # 为每一单分配时间点
    result_rows = []
    for shop_name, group in data.groupby(shop_col, sort=False):
        stats = shop_stats[shop_stats[shop_col] == shop_name].iloc[0]
        interval = stats['每单间隔(分钟)']
        total_qty = stats['店铺总单量']

        # 按原始顺序排列组内数据
        group = group.reset_index(drop=True)

        # 为组内每一行分配时间
        for idx, (_, row) in enumerate(group.iterrows()):
            minutes_offset = idx * interval
            exec_time = base_time + timedelta(minutes=minutes_offset)

            new_row = row.to_dict()
            new_row['执行时间'] = exec_time.strftime('%H:%M')
            new_row['距离首单(分钟)'] = round(minutes_offset, 1)
            new_row['店铺总单量'] = total_qty
            new_row['每单间隔(分钟)'] = round(interval, 1)
            result_rows.append(new_row)

    result_df = pd.DataFrame(result_rows)

    # 按执行时间排序
    result_df = result_df.sort_values(by='执行时间').reset_index(drop=True)

    # 添加序号（避免与已有列冲突）
    if '序号' in result_df.columns:
        result_df.drop(columns=['序号'], inplace=True)
    result_df.insert(0, '序号', range(1, len(result_df) + 1))

    # 输出路径
    if output_path is None:
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        output_path = os.path.join(desktop, f"汇总表_{datetime.now().strftime('%m%d_%H%M')}.xlsx")

    # 写入Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        result_df.to_excel(writer, sheet_name='汇总表', index=False)

        # 添加店铺统计sheet
        shop_stats['预计完成时间'] = shop_stats['店铺总单量'].apply(
            lambda q: (base_time + timedelta(minutes=600 - 600/q)).strftime('%H:%M') if q > 0 else ''
        )
        shop_stats.to_excel(writer, sheet_name='店铺统计', index=False)

    print(f"\n汇总完成！已保存到: {output_path}")
    print(f"总单数: {len(result_df)}")
    return output_path

if __name__ == '__main__':
    import sys

    # 默认文件夹
    default_folder = r'C:\Users\Administrator\Desktop\1'

    if len(sys.argv) > 1:
        folder = sys.argv[1]
    else:
        folder = default_folder

    if not os.path.exists(folder):
        print(f"文件夹不存在: {folder}")
        input("按回车键退出...")
        sys.exit(1)

    process_folder(folder)
    input("\n按回车键退出...")
