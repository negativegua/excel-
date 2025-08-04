import pandas as pd
import os
import sys
from tqdm import tqdm
import re

# 配置列索引（从0开始）
COLUMN_INDICES = [6, 7, 11, 12, 17]  # G,H,L,M,Q列
COLUMN_NAMES = ['G', 'H', 'L', 'M', 'Q']

def truncate_h_column(h_value):
    """保留H列最后9个字符"""
    return str(h_value)[-9:] if pd.notna(h_value) else ""

def sanitize_sheet_name(name, max_len=28):
    """生成合法工作表名称"""
    clean_name = re.sub(r'[\\/*?:[\]]', '', str(name))
    return clean_name[-max_len:] if len(clean_name) > max_len else clean_name

def process_all_sheets(input_file):
    print(f"\n处理文件: {input_file} (大小: {round(os.path.getsize(input_file)/(1024*1024),2)}MB)")
    
    try:
        xl = pd.ExcelFile(input_file)
        all_sheets_data = []
        
        # 首先收集所有工作表的数据
        for sheet in xl.sheet_names:
            df = pd.read_excel(
                input_file,
                sheet_name=sheet,
                usecols=COLUMN_INDICES,
                dtype={'L': 'int32', 'Q': 'float32'}
            )
            df.columns = COLUMN_NAMES
            df['H_short'] = df['H'].apply(truncate_h_column)
            df['Sheet'] = sheet  # 标记来源工作表
            all_sheets_data.append(df)
        
        if not all_sheets_data:
            print("没有有效数据，跳过处理")
            return
            
        # 合并所有工作表数据
        combined_df = pd.concat(all_sheets_data, ignore_index=True)
        
        # 处理合并后的数据
        output_file = f"{os.path.splitext(input_file)[0]}_combined_output.xlsx"
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            for m in [0, 1]:
                if m in combined_df['M'].values:
                    process_m_data(combined_df[combined_df['M'] == m].copy(), m, writer)
        
        print(f"\n结果保存到: {output_file}")
    except Exception as e:
        print(f"\n处理失败: {str(e)}")

def process_m_data(df, m, writer):
    """处理单个M值的数据"""
    summary_data = []
    
    # 按G和H_short分组
    grouped = df.groupby(['G', 'H_short'])
    
    for (g, h_short), group in tqdm(grouped, desc=f"处理M={m}"):
        try:
            # 获取原始H值
            original_h = group['H'].iloc[0] if not group.empty else ""
            
            # 关键点数据
            key_data = {}
            for l in [85, 100, 115]:
                match = group[group['L'] == l]
                if not match.empty:
                    key_data[f"Q@L{l}"] = match['Q'].iloc[0]
                else:
                    key_data[f"Q@L{l}"] = None
            
            # 查找第一个非零Q值（完整50-150范围）
            first_non_zero = None
            for l in range(50, 151):
                match = group[group['L'] == l]
                if not match.empty and match['Q'].iloc[0] != 0:
                    first_non_zero = match.iloc[0]
                    break
            
            # 添加到汇总
            summary_data.append({
                'G': g,
                'H_short': h_short,
                'H_full': original_h,
                'M': m,
                **key_data,
                'First_L': first_non_zero['L'] if first_non_zero is not None else None,
                'First_Q': first_non_zero['Q'] if first_non_zero is not None else None,
                '数据来源': group['Sheet'].nunique()  # 涉及的工作表数量
            })
            
            # 处理范围数据（确保去重）
            process_range_data(group, g, h_short, m, writer)
            
        except Exception as e:
            print(f"\n处理组 {g}_{h_short} 出错: {str(e)}")
            continue
    
    # 写入汇总表
    if summary_data:
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(
            writer,
            sheet_name=f"Summary_M{m}",
            index=False
        )

def process_range_data(group, g, h_short, m, writer):
    """处理并写入范围数据（确保去重）"""
    # 先按L和Q去重
    unique_data = group.drop_duplicates(['L', 'Q'])
    
    # 85-100范围
    range_85_100 = unique_data[
        (unique_data['L'] >= 85) & (unique_data['L'] <= 100)
    ][['L', 'Q']]
    
    if not range_85_100.empty:
        full_range = pd.DataFrame({'L': range(85, 101)})
        merged = pd.merge(full_range, range_85_100, on='L', how='left')
        
        # 转置
        transposed = merged.set_index('L').T
        transposed.columns = [f"L{col}" for col in transposed.columns]
        transposed.insert(0, 'G', g)
        transposed.insert(1, 'H', h_short)
        transposed.insert(2, 'M', m)
        
        # 写入
        sheet_name = sanitize_sheet_name(f"{g}_{h_short}_85-100_M{m}")
        transposed.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )
    
    # 100-115范围
    range_100_115 = unique_data[
        (unique_data['L'] >= 100) & (unique_data['L'] <= 115)
    ][['L', 'Q']]
    
    if not range_100_115.empty:
        full_range = pd.DataFrame({'L': range(100, 116)})
        merged = pd.merge(full_range, range_100_115, on='L', how='left')
        
        transposed = merged.set_index('L').T
        transposed.columns = [f"L{col}" for col in transposed.columns]
        transposed.insert(0, 'G', g)
        transposed.insert(1, 'H', h_short)
        transposed.insert(2, 'M', m)
        
        sheet_name = sanitize_sheet_name(f"{g}_{h_short}_100-115_M{m}")
        transposed.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("使用方法: python report.py 文件名.xlsx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print(f"文件不存在: {input_file}")
        sys.exit(1)
    
    process_all_sheets(input_file)