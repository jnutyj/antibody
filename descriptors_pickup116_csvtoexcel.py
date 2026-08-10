# -*- coding: utf-8 -*-
"""
Created on Tue Sep  2 16:26:34 2025

@author: Xin.Li
"""

import pandas as pd
import os

# 输入文件路径
input_file = "protein_descriptors.csv"  # 替换为您的CSV文件路径

# 输出文件路径
output_file = "output.xlsx"  # 替换为您想要的输出路径

# 要提取的列名或列索引列表
# 可以使用列名: ["列1", "列5", "列10"]
# 或列索引(从1开始): [1, 5, 10]
# 或混合使用: ["列1", 5, "列10"]
columns_to_extract = ["Name", "All_AggScore", "All_Aggrescan_a4v", "All_Formal_Charge", "CDR_AggScore",
                      "CDR_Aggrescan_a4v", "CDR_Hydrophobic_Patch_Energy", "CDR_Hydrophobic_Patch_Energy_gt15", "CDR_Hydrophobic_Patch_Energy_gt30",
                      "CDR_Negative_Patch_Energy", "CDR_Negative_Patch_Energy_gt30", "CDR_Negative_Patch_Energy_gt50", "CDR_Positive_Patch_Energy", 
                      "CDR_Positive_Patch_Energy_gt30", "CDR_Positive_Patch_Energy_gt50", "CDR_Zeta_Potential", "CDR_Zyggregator_profile_smoothed", 
                      "CDR_Zyggregator_profile_smoothed_pos", "FR_AggScore", "FR_Aggrescan_a4v", "FR_Greasy_SASA", 
                      "FR_Hydrophobic_Patch_Energy", "FR_Hydrophobic_Patch_Energy_gt15", "FR_Hydrophobic_Patch_Energy_gt30", "FR_Negative_Patch_Energy",
                      "FR_Negative_Patch_Energy_gt30", "FR_Negative_Patch_Energy_gt50", "FR_Positive_Patch_Energy", "FR_Positive_Patch_Energy_gt30", 
                      "FR_Positive_Patch_Energy_gt50", "FR_Zeta_Potential", "FR_Zyggregator_profile_smoothed", "FR_Zyggregator_profile_smoothed_pos", 
                      "H1_AggScore", "H1_Aggrescan_a4v", "H1_Hydrophobic_Patch_Energy", "H1_Hydrophobic_Patch_Energy_gt15", 
                      "H1_Hydrophobic_Patch_Energy_gt30", "H1_Negative_Patch_Energy", "H1_Negative_Patch_Energy_gt30", "H1_Negative_Patch_Energy_gt50", 
                      "H1_Positive_Patch_Energy", "H1_Positive_Patch_Energy_gt30", "H1_Positive_Patch_Energy_gt50", "H2_AggScore", 
                      "H2_Aggrescan_a4v", "H2_Hydrophobic_Patch_Energy", "H2_Hydrophobic_Patch_Energy_gt15", "H2_Hydrophobic_Patch_Energy_gt30", 
                      "H2_Negative_Patch_Energy", "H2_Negative_Patch_Energy_gt30", "H2_Negative_Patch_Energy_gt50", "H2_Positive_Patch_Energy", 
                      "H2_Positive_Patch_Energy_gt30", "H2_Positive_Patch_Energy_gt50", "H3_AggScore", "H3_Aggrescan_a4v", 
                      "H3_Hydrophobic_Patch_Energy", "H3_Hydrophobic_Patch_Energy_gt15", "H3_Hydrophobic_Patch_Energy_gt30", "H3_Negative_Patch_Energy", 
                      "H3_Negative_Patch_Energy_gt30", "H3_Negative_Patch_Energy_gt50", "H3_Positive_Patch_Energy", "H3_Positive_Patch_Energy_gt30", 
                      "H3_Positive_Patch_Energy_gt50", "HFR1_AggScore", "HFR1_Aggrescan_a4v", "HFR1_Hydrophobic_Patch_Energy", 
                      "HFR1_Hydrophobic_Patch_Energy_gt15", "HFR1_Hydrophobic_Patch_Energy_gt30", "HFR1_Negative_Patch_Energy", "HFR1_Negative_Patch_Energy_gt30", 
                      "HFR1_Negative_Patch_Energy_gt50", "HFR1_Positive_Patch_Energy", "HFR1_Positive_Patch_Energy_gt30", "HFR1_Positive_Patch_Energy_gt50", 
                      "HFR2_AggScore", "HFR2_Aggrescan_a4v", "HFR2_Hydrophobic_Patch_Energy", "HFR2_Hydrophobic_Patch_Energy_gt15", 
                      "HFR2_Hydrophobic_Patch_Energy_gt30", "HFR2_Negative_Patch_Energy", "HFR2_Negative_Patch_Energy_gt30", "HFR2_Negative_Patch_Energy_gt50", 
                      "HFR2_Positive_Patch_Energy", "HFR2_Positive_Patch_Energy_gt30", "HFR2_Positive_Patch_Energy_gt50", "HFR3_AggScore", 
                      "HFR3_Aggrescan_a4v", "HFR3_Hydrophobic_Patch_Energy", "HFR3_Hydrophobic_Patch_Energy_gt15", "HFR3_Hydrophobic_Patch_Energy_gt30", 
                      "HFR3_Negative_Patch_Energy", "HFR3_Negative_Patch_Energy_gt30", "HFR3_Negative_Patch_Energy_gt50", "HFR3_Positive_Patch_Energy", 
                      "HFR3_Positive_Patch_Energy_gt30", "HFR3_Positive_Patch_Energy_gt50", "HFR4_AggScore", "HFR4_Aggrescan_a4v", 
                      "HFR4_Hydrophobic_Patch_Energy", "HFR4_Hydrophobic_Patch_Energy_gt15", "HFR4_Hydrophobic_Patch_Energy_gt30", "HFR4_Negative_Patch_Energy", 
                      "HFR4_Negative_Patch_Energy_gt30", "HFR4_Negative_Patch_Energy_gt50", "HFR4_Positive_Patch_Energy", "HFR4_Positive_Patch_Energy_gt30", 
                      "HFR4_Positive_Patch_Energy_gt50", "Max_Size_Hyd_Patches", "Max_Size_Neg_Patches", "Max_Size_Pos_Patches", 
                      "Sum_Size_Hyd_Patches", "Sum_Size_Neg_Patches", "Sum_Size_Pos_Patches", "pI_PROPKA_based"]  # 替换为您需要的列

# 读取CSV文件
df = pd.read_csv(input_file)

# 处理列索引（如果提供的是数字而不是列名）
available_columns = df.columns.tolist()
processed_columns = []

for col in columns_to_extract:
    if isinstance(col, int) or (isinstance(col, str) and col.isdigit()):
        col_idx = int(col) - 1  # 转换为0-based索引
        if 0 <= col_idx < len(available_columns):
            processed_columns.append(available_columns[col_idx])
        else:
            print(f"警告: 列索引 {col} 超出范围 (1-{len(available_columns)})")
    else:
        if col in available_columns:
            processed_columns.append(col)
        else:
            print(f"警告: 列 '{col}' 在CSV文件中不存在")

# 提取指定列
extracted_df = df[processed_columns]

# 确保输出目录存在
os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)

# 保存为Excel文件
extracted_df.to_excel(output_file, index=False)

print(f"成功提取并保存到 {output_file}")
