import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re
import os
import sys
from tqdm import tqdm
from collections import defaultdict

def get_sheet_files(zip_ref):
    """获取所有工作表XML文件路径"""
    return [name for name in zip_ref.namelist() 
            if 'worksheets/sheet' in name.lower() and name.endswith('.xml')]

def parse_sheet_number(filename):
    """从文件名提取工作表序号"""
    match = re.search(r'sheet(\d+)\.xml', filename, re.IGNORECASE)
    return int(match.group(1)) if match else -1

def get_sheet_name(zip_ref, sheet_num):
    """获取实际工作表名称"""
    try:
        with zip_ref.open('xl/workbook.xml') as f:
            workbook = ET.parse(f).getroot()
            ns = {'ns': workbook.tag.split('}')[0].strip('{')} if '}' in workbook.tag else {'ns': ''}
            for sheet in workbook.findall('.//ns:sheets/ns:sheet', ns):
                if sheet.get('sheetId') == str(sheet_num + 1):
                    return sheet.get('name')
    except Exception as e:
        print(f"获取工作表名称错误: {e}")
    return f"sheet{sheet_num}"

def get_cell_value(cell, shared_strings, ns):
    """获取单元格值，优化处理各种类型"""
    if cell is None:
        return ''
    
    if cell.get('t') == 'inlineStr':
        is_node = cell.find('.//ns:is', ns)
        if is_node is not None:
            t_node = is_node.find('.//ns:t', ns)
            return t_node.text if t_node is not None else ''
    
    elif cell.get('t') == 's':
        v_node = cell.find('.//ns:v', ns)
        if v_node is not None:
            try:
                return shared_strings[int(v_node.text)]
            except (ValueError, IndexError):
                return ''
    
    v_node = cell.find('.//ns:v', ns)
    return v_node.text if v_node is not None else ''

def deep_scan_excel(file_path):
    print(">>> 启动Excel扫描引擎 (G和最小L值合并版) <<<")
    
    try:
        with zipfile.ZipFile(file_path) as z:
            # 获取共享字符串
            shared_strings = []
            try:
                with z.open('xl/sharedStrings.xml') as f:
                    sst = ET.parse(f).getroot()
                    ns_sst = {'ns': sst.tag.split('}')[0].strip('{')} if '}' in sst.tag else {'ns': ''}
                    shared_strings = [t.text if t.text else '' for t in sst.findall('.//ns:t', ns_sst)]
            except Exception as e:
                print(f"共享字符串读取警告: {e}")
            
            # 使用更高效的数据结构
            gh_data = defaultdict(lambda: {
                'min_l': float('inf'),
                'count': 0,
                'source_sheets': set()
            })
            
            # 添加非零值检测标志
            found_non_zero = False
            
            # 单次扫描处理
            for sheet_file in tqdm(get_sheet_files(z), desc="处理工作表中"):
                try:
                    sheet_num = parse_sheet_number(sheet_file)
                    sheet_name = get_sheet_name(z, sheet_num)
                    
                    with z.open(sheet_file) as f:
                        root = ET.parse(f).getroot()
                        ns = {'ns': root.tag.split('}')[0].strip('{')} if '}' in root.tag else {'ns': ''}
                        
                        cells = {}
                        row_numbers = set()
                        for cell in root.findall('.//ns:c', ns):
                            if (ref := cell.get('r')):
                                col = re.sub(r'\d+', '', ref).upper()
                                row = re.sub(r'[A-Za-z]+', '', ref)
                                cells[(col, row)] = cell
                                row_numbers.add(int(row))
                        
                        if not row_numbers:
                            continue
                            
                        min_row = min(row_numbers)
                        process_rows = [str(r) for r in row_numbers if r > min_row]
                        
                        # 记录本工作表已处理的G-H组合（用于去重）
                        sheet_processed_gh = set()
                        
                        for row in process_rows:
                            q_cell = cells.get(('Q', row))
                            if q_cell is None:
                                continue
                                
                            q_val = get_cell_value(q_cell, shared_strings, ns)
                            if q_val and q_val not in ('0', '0.0'):
                                found_non_zero = True  # 标记找到非零值
                                g_val = get_cell_value(cells.get(('G', row)), shared_strings, ns)
                                h_val = get_cell_value(cells.get(('H', row)), shared_strings, ns)
                                l_val = get_cell_value(cells.get(('L', row)), shared_strings, ns)
                                
                                try:
                                    l_num = float(l_val) if l_val else float('inf')
                                    gh_key = (g_val, h_val)
                                    
                                    # 更新最小L值和来源
                                    if l_num < gh_data[gh_key]['min_l']:
                                        gh_data[gh_key]['min_l'] = l_num
                                        gh_data[gh_key]['source_sheets'] = {sheet_name}
                                    elif l_num == gh_data[gh_key]['min_l']:
                                        gh_data[gh_key]['source_sheets'].add(sheet_name)
                                    
                                    # 同一工作表内相同G-H组合只计一次
                                    if gh_key not in sheet_processed_gh:
                                        gh_data[gh_key]['count'] += 1
                                        sheet_processed_gh.add(gh_key)
                                except ValueError:
                                    continue
                
                except Exception as e:
                    print(f"处理工作表 {sheet_file} 时出错: {e}")
            
            # 检查是否找到非零值
            if not found_non_zero:
                print("\n▶ 未在任何工作表中找到Q列非零值")
                return
            
            # 第二阶段：合并相同G和最小L值的记录
            gl_merged = defaultdict(lambda: {
                'h_values': set(),
                'total_count': 0,
                'source_sheets': set()
            })
            
            for (g_val, h_val), data in gh_data.items():
                if data['min_l'] == float('inf'):
                    continue
                    
                gl_key = (g_val, data['min_l'])
                gl_merged[gl_key]['h_values'].add(h_val)
                gl_merged[gl_key]['total_count'] += data['count']
                gl_merged[gl_key]['source_sheets'].update(data['source_sheets'])
            
            # 生成最终结果
            output_data = []
            for (g_val, min_l), data in gl_merged.items():
                output_data.append({
                    'G': g_val,
                    'H': ' | '.join(sorted(data['h_values'])),
                    '最小L值': min_l,
                    '出现次数': data['total_count'],
                    '来源工作表': ', '.join(sorted(data['source_sheets']))
                })

            # 生成输出文件名（基于输入文件名）
            base_name = os.path.splitext(file)[0]  # 去掉扩展名
            output_file = f"{base_name}-结果.xlsx"
        
            # 输出结果
            if output_data:
                df = pd.DataFrame(output_data)
                df = df[['G', 'H', '最小L值', '出现次数', '来源工作表']]
            
                print("\n=== 数据验证 ===")
                print(f"总合并记录数: {len(df)}")
                print("前5行示例:")
                print(df.head())
            
                empty_gh = df[(df['G'] == '') | (df['H'] == '')]
                if not empty_gh.empty:
                    print(f"\n警告: 发现 {len(empty_gh)} 条记录G/H列为空")
            
                df.to_excel(output_file, index=False, engine='openpyxl')
                print(f"\n▶ 结果已保存到 {output_file} 文件！")
        
    except Exception as e:
        print(f"!!! 扫描失败: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("使用方法: python report.py 文件名.xlsx")
        sys.exit(1)
    
    file = sys.argv[1]
    if not os.path.exists(file):
        print(f"文件不存在: {file}")
        sys.exit(1)
    deep_scan_excel(file)