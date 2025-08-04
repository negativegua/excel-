import pandas as pd
import sys
import os
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")

def safe_compare(df1, df2):
    """安全比较函数，处理索引和列名差异"""
    # 重置索引并确保列名一致
    df1 = df1.reset_index(drop=True)
    df2 = df2.reset_index(drop=True)
    
    # 标准化列名
    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()
    
    # 确保列顺序一致
    df2 = df2[df1.columns]
    
    try:
        return df1.compare(df2)
    except ValueError as e:
        print(f"⚠️ 比较失败: {str(e)}")
        print("尝试逐行比较...")
        
        diff_locations = df1 != df2
        diff_rows = diff_locations.any(axis=1)
        
        if not diff_rows.any():
            return pd.DataFrame()
        else:
            return df1[diff_rows].compare(df2[diff_rows])

def compare_excel(file1, file2, ignore_columns=None, sheet1_name=None, sheet2_name=None):
    """支持不同Sheet名比较"""
    if ignore_columns is None:
        ignore_columns = []

    print(f"\n⏳ 预处理：忽略列 {ignore_columns}")
    print(f"🔍 文件验证:\n- 文件1: {Path(file1).resolve()}\n- 文件2: {Path(file2).resolve()}")

    try:
        with pd.ExcelFile(file1) as xls1, pd.ExcelFile(file2) as xls2:
            sheets1 = xls1.sheet_names
            sheets2 = xls2.sheet_names

        sheet1 = sheet1_name if sheet1_name else sheets1[0]
        sheet2 = sheet2_name if sheet2_name else sheets2[0]
        
        print(f"\n📌 比较配置:")
        print(f"- 文件1 Sheet: '{sheet1}' (可选: {sheets1})")
        print(f"- 文件2 Sheet: '{sheet2}' (可选: {sheets2})")

        df1 = pd.read_excel(file1, sheet_name=sheet1).fillna("")
        df2 = pd.read_excel(file2, sheet_name=sheet2).fillna("")

        print(f"\n🔎 列名对比:")
        print(f"- 文件1: {df1.columns.tolist()}")
        print(f"- 文件2: {df2.columns.tolist()}")

        cols_to_drop = [col for col in ignore_columns if col in df1.columns]
        if cols_to_drop:
            print(f"\n🗑️ 忽略列: {cols_to_drop}")
            df1 = df1.drop(columns=cols_to_drop)
            df2 = df2.drop(columns=cols_to_drop)
            print(f"🔧 比较列: {df1.columns.tolist()}")

        print([repr(c) for c in df1.columns])  # 显示列名的原始形式
        print(f"文件1 形状: {df1.shape}, 文件2 形状: {df2.shape}")

        # 使用安全比较
        diff = safe_compare(df1, df2)
        if diff.empty:
            print("\n✅ 内容完全相同")
            return True
        else:
            print("\n❌ 发现差异:")
            print(diff if not diff.empty else "(无有效差异)")
            return False

    except Exception as e:
        print(f"\n❌ 错误: {str(e)}")
        return False

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Excel文件比较工具")
    parser.add_argument("file1", help="第一个Excel文件")
    parser.add_argument("file2", help="第二个Excel文件")
    parser.add_argument("--ignore", nargs="+", default=[], help="要忽略的列名")
    parser.add_argument("--sheet1", help="指定文件1的sheet名")
    parser.add_argument("--sheet2", help="指定文件2的sheet名")
    args = parser.parse_args()

    success = compare_excel(
        args.file1, args.file2,
        ignore_columns=args.ignore,
        sheet1_name=args.sheet1,
        sheet2_name=args.sheet2
    )
    input("\n按回车键退出...")
    sys.exit(0 if success else 1)