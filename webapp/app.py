"""补单汇总工具 - Flask 网页应用"""
import os
import uuid
import shutil
import json
import glob
from datetime import datetime, timedelta, time, date
from flask import Flask, request, jsonify, send_file, render_template
import numpy as np

from utils import (
    read_excel_with_fallback,
    detect_abnormal_files,
    assign_shops_to_brushers,
    auto_schedule,
    interleave_shop_links,
    extract_customer_from_filename,
    extract_wps_images,
    extract_standard_images,
    extract_img_id,
    get_header_column_map,
    compute_stats,
    generate_output_excel,
)

app = Flask(__name__)
app.secret_key = os.urandom(24).hex()


def to_json_safe(df):
    """将DataFrame转换为JSON安全的dict列表"""
    records = df.fillna('').to_dict(orient='records')
    safe = []
    for row in records:
        safe_row = {}
        for k, v in row.items():
            if isinstance(v, (datetime, date, time)):
                safe_row[k] = str(v)
            elif isinstance(v, (np.integer,)):
                safe_row[k] = int(v)
            elif isinstance(v, (np.floating,)):
                safe_row[k] = float(v)
            elif isinstance(v, np.bool_):
                safe_row[k] = bool(v)
            elif v is None or (isinstance(v, float) and np.isnan(v)):
                safe_row[k] = ''
            else:
                safe_row[k] = v
        safe.append(safe_row)
    return safe

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'outputs')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ========== 路由 ==========

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload():
    """批量上传Excel文件"""
    if 'files' not in request.files:
        return jsonify({'error': '未选择文件'}), 400

    files = request.files.getlist('files')
    session_id = request.form.get('session_id', uuid.uuid4().hex)
    session_dir = os.path.join(UPLOAD_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)

    uploaded = []
    for f in files:
        if f.filename and f.filename.lower().endswith(('.xlsx', '.xls')):
            safe_name = f.filename.replace('/', '_').replace('\\', '_')
            path = os.path.join(session_dir, safe_name)
            f.save(path)
            uploaded.append(safe_name)

    return jsonify({
        'session_id': session_id,
        'files': uploaded,
        'count': len(uploaded)
    })


@app.route('/api/process', methods=['POST'])
def process():
    """执行完整处理流程"""
    data = request.get_json()
    session_id = data.get('session_id', '')
    start_time = data.get('start_time', '08:00')
    end_time = data.get('end_time', '23:00')
    num_workers = int(data.get('num_workers', 1))
    min_interval = float(data.get('min_interval', 20))

    session_dir = os.path.join(UPLOAD_DIR, session_id)
    if not os.path.isdir(session_dir):
        return jsonify({'error': '会话已过期，请重新上传文件'}), 400

    filepaths = []
    for ext in ('*.xlsx', '*.xls'):
        filepaths.extend(glob.glob(os.path.join(session_dir, ext)))
    filepaths = [f for f in filepaths if not os.path.basename(f).startswith('~$')]

    if not filepaths:
        return jsonify({'error': '未找到Excel文件'}), 400

    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    import io

    try:
        # 1. 异常检测
        abnormal = detect_abnormal_files(filepaths)
        invalid_basenames = set()
        for inv in abnormal['invalid']:
            invalid_basenames.add(inv['file'])
        for dup in abnormal['duplicates']:
            invalid_basenames.update(dup)

        valid_files = [f for f in filepaths if os.path.basename(f) not in invalid_basenames]

        if not valid_files:
            return jsonify({'error': '没有可处理的正常文件'}), 400

        # 2. 读取数据
        all_data = []
        customer_from_file = {}
        file_images = {}
        file_std_images = {}

        for f in valid_files:
            basename = os.path.basename(f)
            customer = extract_customer_from_filename(basename)
            customer_from_file[basename] = customer

            images = extract_wps_images(f)
            file_images[f] = images

            std_images_raw = extract_standard_images(f)
            std_images_mapped = {}
            if std_images_raw:
                header_map = get_header_column_map(f)
                for (row, col), img_data in std_images_raw.items():
                    std_col = header_map.get(col)
                    if std_col:
                        std_images_mapped.setdefault(std_col, {})[row] = img_data
                file_std_images[f] = std_images_mapped
            else:
                file_std_images[f] = {}

            df = read_excel_with_fallback(f)
            if len(df) == 0:
                continue

            if '客户(不填)' in df.columns:
                df['客户(不填)'] = df['客户(不填)'].astype(object)
                empty_mask = df['客户(不填)'].isna() | (df['客户(不填)'] == '')
                if empty_mask.any() and customer:
                    df.loc[empty_mask, '客户(不填)'] = customer
            elif customer:
                df['客户(不填)'] = customer

            df['来源文件'] = basename

            # 过滤无图片行
            if '__drawing_row__' in df.columns:
                pic_col = None
                for c in df.columns:
                    if str(c) == '主图':
                        pic_col = c
                        break
                if pic_col is not None:
                    valid_mask = []
                    for _, row in df.iterrows():
                        has_img = False
                        cell_val = row.get(pic_col)
                        if isinstance(cell_val, str):
                            img_id = extract_img_id(cell_val)
                            if img_id and img_id in images:
                                has_img = True
                        if not has_img:
                            dr = row.get('__drawing_row__')
                            if dr is not None:
                                dr_int = int(dr)
                                if pic_col in std_images_mapped and dr_int in std_images_mapped[pic_col]:
                                    has_img = True
                        valid_mask.append(has_img)
                    if not all(valid_mask):
                        df = df[valid_mask].reset_index(drop=True)

            all_data.append(df)

        combined = pd.concat(all_data, ignore_index=True)

        if '商品ID' in combined.columns:
            combined['商品ID'] = combined['商品ID'].apply(
                lambda x: str(int(x)) if pd.notna(x) and str(x).replace('.', '', 1).isdigit() else str(x) if pd.notna(x) else ''
            )

        if '备注' not in combined.columns:
            combined['备注'] = ''

        # 3. 识别关键列
        shop_col = None
        for col in combined.columns:
            if '店铺名称' in str(col) and combined[col].notna().sum() > 0:
                shop_col = col
                break
        if shop_col is None:
            for col in combined.columns:
                if ('客户' in str(col) and '店铺' in str(col)) or str(col) == '客户(不填)':
                    if combined[col].notna().sum() > 0:
                        shop_col = col
                        break
        if shop_col is None:
            for col in combined.columns:
                if str(col).startswith('Unnamed'):
                    continue
                if combined[col].apply(lambda x: isinstance(x, str) and len(x) > 3).sum() > 0:
                    shop_col = col
                    break

        qty_col = None
        for col in combined.columns:
            if '补单数量' in str(col) or '单量' in str(col) or '数量' in str(col):
                qty_col = col
                break

        if shop_col is None or qty_col is None:
            return jsonify({'error': '未识别到店铺列或单量列'}), 400

        # 4. 过滤 & 清理
        valid_mask = combined[shop_col].notna() & combined[qty_col].notna()
        qty_numeric = pd.to_numeric(combined[qty_col], errors='coerce')
        valid_mask = valid_mask & (qty_numeric > 0)
        data = combined[valid_mask].copy()
        data[qty_col] = pd.to_numeric(data[qty_col], errors='coerce')
        data = data[data[qty_col] > 0].copy()
        data[shop_col] = data[shop_col].astype(str).str.replace(r'[\n\r]+', '', regex=True).str.strip()

        # 5. 拆分 + 链接交叉
        expanded_rows = []
        for _, row in data.iterrows():
            qty = int(row[qty_col])
            if qty > 1:
                for _ in range(qty):
                    new_row = row.copy()
                    new_row[qty_col] = 1
                    expanded_rows.append(new_row)
            else:
                expanded_rows.append(row)
        data = pd.DataFrame(expanded_rows).reset_index(drop=True)
        data[qty_col] = pd.to_numeric(data[qty_col], errors='coerce')
        data = interleave_shop_links(data, shop_col)

        # 6. 调度
        shop_orders = data.groupby(shop_col)[qty_col].sum().astype(int).to_dict()
        if num_workers < 1:
            num_workers = 1

        schedule_list = auto_schedule(start_time, end_time, shop_orders, num_workers, min_interval)

        # 7. 映射回数据行
        shop_schedule_map = {}
        for item in schedule_list:
            shop_schedule_map.setdefault(item['店铺'], []).append(item)

        start_h, start_m = map(int, start_time.split(':'))
        base_time = datetime(2026, 1, 1, start_h, start_m)

        result_rows = []
        for shop_name, group in data.groupby(shop_col, sort=False):
            group = group.reset_index(drop=True)
            assignments_local = shop_schedule_map.get(shop_name, [])
            total_qty = len(group)
            for idx, (_, row) in enumerate(group.iterrows()):
                if idx < len(assignments_local):
                    item = assignments_local[idx]
                else:
                    item = {'时间分钟': 0, '刷手': '刷手1'}
                exec_time = base_time + timedelta(minutes=item['时间分钟'])
                new_row = row.to_dict()
                new_row['执行时间'] = exec_time.strftime('%H:%M')
                new_row['刷手'] = item['刷手']
                new_row['距离首单(分钟)'] = round(item['时间分钟'], 1)
                new_row['店铺总单量'] = total_qty
                result_rows.append(new_row)

        result_df = pd.DataFrame(result_rows)
        result_df = result_df.sort_values(by='执行时间').reset_index(drop=True)

        if '序号' in result_df.columns:
            result_df.drop(columns=['序号'], inplace=True)
        if '备注' not in result_df.columns:
            result_df['备注'] = ''

        # 列顺序
        cols = list(result_df.columns)
        ordered = ['执行时间', '刷手', '备注']
        unnamed_cols = [c for c in cols if str(c).startswith('Unnamed')]
        remaining = [c for c in cols if c not in ordered and c not in unnamed_cols]
        result_df = result_df[ordered + unnamed_cols + remaining]
        result_df.insert(0, '序号', range(1, len(result_df) + 1))

        # 8. 统计
        df_file_stats, df_customer_stats, df_shop_stats = compute_stats(data, shop_col, qty_col, customer_from_file)

        # 店铺调度统计
        total_work_minutes = (int(end_time.split(':')[0]) * 60 + int(end_time.split(':')[1])) - (start_h * 60 + start_m)
        shop_stats = data.groupby(shop_col)[qty_col].sum().reset_index()
        shop_stats.columns = [shop_col, '店铺总单量']
        shop_last_time = {}
        shop_avg_interval = {}
        for shop, items in shop_schedule_map.items():
            times = sorted([it['时间分钟'] for it in items])
            shop_last_time[shop] = times[-1] if times else 0
            shop_avg_interval[shop] = (times[-1] - times[0]) / (len(times) - 1) if len(times) > 1 else 0
        shop_stats['平均间隔(分钟)'] = shop_stats[shop_col].map(shop_avg_interval)
        shop_stats['预计完成时间'] = shop_stats[shop_col].map(
            lambda s: (base_time + timedelta(minutes=shop_last_time.get(s, 0))).strftime('%H:%M')
        )

        # 9. 生成输出文件
        output_id = uuid.uuid4().hex
        output_dir = os.path.join(OUTPUT_DIR, session_id)
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, f'汇总表_{output_id}.xlsx')
        generate_output_excel(output_path, result_df, shop_stats, df_file_stats, df_customer_stats, df_shop_stats,
                               file_images=file_images, file_std_images=file_std_images)

        # 10. 构建预览数据
        anomaly_rows = []
        for dup in abnormal['duplicates']:
            anomaly_rows.append({'文件名': f"{dup[0]} <-> {dup[1]}", '异常原因': '内容完全相同的文件'})
        for inv in abnormal['invalid']:
            anomaly_rows.append({'文件名': inv['file'], '异常原因': inv['reason']})

        summary_data = to_json_safe(result_df.head(200))

        return jsonify({
            'output_id': output_id,
            'session_id': session_id,
            'summary': summary_data,
            'summary_columns': [str(c) for c in result_df.columns],
            'total_rows': len(result_df),
            'file_stats': to_json_safe(df_file_stats),
            'customer_stats': to_json_safe(df_customer_stats),
            'shop_stats': to_json_safe(df_shop_stats),
            'shop_schedule': to_json_safe(shop_stats),
            'anomalies': anomaly_rows,
            'total_files': len(valid_files),
            'total_customers': len(df_customer_stats),
            'total_shops': len(df_shop_stats),
            'total_orders': int(data[qty_col].sum()),
        })

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500


@app.route('/api/download/<session_id>/<output_id>')
def download(session_id, output_id):
    """下载生成的Excel文件"""
    output_dir = os.path.join(OUTPUT_DIR, session_id)
    if not os.path.isdir(output_dir):
        return jsonify({'error': '文件不存在或已过期'}), 404
    # Find the file matching this output_id
    for fname in os.listdir(output_dir):
        if output_id in fname and fname.endswith('.xlsx'):
            return send_file(
                os.path.join(output_dir, fname),
                as_attachment=True,
                download_name='汇总表.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    return jsonify({'error': '文件已过期'}), 404


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
