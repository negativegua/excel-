import os
import numpy as np
from scipy import signal

def batch_resample_txt_files(input_folder, output_folder, original_rate, target_rate):
    """
    批量重采样文件夹中的所有TXT文件
    :param input_folder: 输入文件夹路径
    :param output_folder: 输出文件夹路径
    :param original_rate: 原始采样率(Hz)
    :param target_rate: 目标采样率(Hz)
    """
    # 确保输出文件夹存在
    os.makedirs(output_folder, exist_ok=True)
    
    # 获取输入文件夹中所有TXT文件
    txt_files = [f for f in os.listdir(input_folder) if f.endswith('.txt')]
    
    for filename in txt_files:
        # 构造完整文件路径
        input_path = os.path.join(input_folder, filename)
        output_path = os.path.join(output_folder, f"resampled_{filename}")
        
        # 读取数据
        try:
            data = np.loadtxt(input_path)
            
            # 计算重采样后的点数
            original_length = len(data)
            target_length = int(original_length * target_rate / original_rate)
            
            # 执行重采样
            resampled_data = signal.resample(data, target_length)
            
            # 保存结果
            np.savetxt(output_path, resampled_data)
            print(f"成功处理: {filename} -> resampled_{filename}")
            
        except Exception as e:
            print(f"处理文件 {filename} 时出错: {str(e)}")

# 使用示例
if __name__ == "__main__":
    # 设置参数
    input_dir = r"C:\Users\admin\Desktop\滑台注入数据\data"  # 输入文件夹路径
    output_dir = r"C:\Users\admin\Desktop\滑台注入数据\processed"  # 输出文件夹路径
    original_fs = 10000  # 原始采样率(Hz)
    target_fs = 2000  # 目标采样率(Hz)
    
    # 执行批量处理
    batch_resample_txt_files(input_dir, output_dir, original_fs, target_fs)
    print("批量重采样完成！")