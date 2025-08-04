import os
import numpy as np

def batch_downsample_txt_files(input_folder, output_folder, downsample_factor):
    """
    批量降采样文件夹中的所有TXT文件
    :param input_folder: 输入文件夹路径(使用原始字符串或双反斜杠)
    :param output_folder: 输出文件夹路径
    :param downsample_factor: 降采样因子(如4表示每4个点取1个)
    """
    # 确保输出文件夹存在
    os.makedirs(output_folder, exist_ok=True)
    
    # 获取输入文件夹中所有TXT文件
    txt_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.txt')]
    
    for filename in txt_files:
        # 构造完整文件路径
        input_path = os.path.join(input_folder, filename)
        output_path = os.path.join(output_folder, f"downsampled_{filename}")
        
        try:
            # 读取数据
            data = np.loadtxt(input_path)
            
            # 执行降采样 - 每隔downsample_factor个点取一个点
            downsampled_data = data[::downsample_factor]
            
            # 保存结果
            np.savetxt(output_path, downsampled_data)
            print(f"成功降采样: {filename} (原始点数: {len(data)}, 降采样后: {len(downsampled_data)})")
            
        except Exception as e:
            print(f"处理文件 {filename} 时出错: {str(e)}")

# 使用示例
if __name__ == "__main__":
    # 设置参数 - 注意路径写法
    input_dir = r"C:\Users\admin\Desktop\滑台注入数据\data"  # 原始字符串
    output_dir = r"C:\Users\admin\Desktop\滑台注入数据\processed"  # 原始字符串
    factor = 5  # 降采样因子(10000Hz→2000Hz需要factor=5)
    
    # 执行批量处理
    batch_downsample_txt_files(input_dir, output_dir, factor)
    print("批量降采样完成！")