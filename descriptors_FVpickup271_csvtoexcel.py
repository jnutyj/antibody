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
output_file = "output_Fv.xlsx"  # 替换为您想要的输出路径

# 要提取的列名或列索引列表
# 可以使用列名: ["列1", "列5", "列10"]
# 或列索引(从1开始): [1, 5, 10]
# 或混合使用: ["列1", 5, "列10"]
columns_to_extract = ["Name", "All_AggScore", "All_Formal_Charge", "All_Hydrophobic_Patch_Energy", "All_Hydrophobic_Patch_Energy_gt15", "All_Hydrophobic_Patch_Energy_gt30", 
                      "All_Negative_Patch_Energy", "All_Negative_Patch_Energy_gt30", "All_Negative_Patch_Energy_gt50", "All_Positive_Patch_Energy", "All_Positive_Patch_Energy_gt30", 
                      "All_Positive_Patch_Energy_gt50", "CDRH_AggScore", "CDRH_Formal_Charge", "CDRH_Hydrophobic_Patch_Energy", "CDRH_Hydrophobic_Patch_Energy_gt15", 
                      "CDRH_Hydrophobic_Patch_Energy_gt30", "CDRH_Negative_Patch_Energy", "CDRH_Negative_Patch_Energy_gt30", "CDRH_Negative_Patch_Energy_gt50", 
                      "CDRH_Positive_Patch_Energy", "CDRH_Positive_Patch_Energy_gt30", "CDRH_Positive_Patch_Energy_gt50", "CDRL_AggScore", "CDRL_Formal_Charge", 
                      "CDRL_Hydrophobic_Patch_Energy", "CDRL_Hydrophobic_Patch_Energy_gt15", "CDRL_Hydrophobic_Patch_Energy_gt30", "CDRL_Negative_Patch_Energy", 
                      "CDRL_Negative_Patch_Energy_gt30", "CDRL_Negative_Patch_Energy_gt50", "CDRL_Positive_Patch_Energy", "CDRL_Positive_Patch_Energy_gt30", 
                      "CDRL_Positive_Patch_Energy_gt50", "CDR_AggScore", "CDR_Aggrescan_a4v", "CDR_Aggrescan_a4v_pos", "CDR_Formal_Charge", "CDR_Hydrophobic_Patch_Energy", 
                      "CDR_Hydrophobic_Patch_Energy_gt15", "CDR_Hydrophobic_Patch_Energy_gt30", "CDR_Negative_Patch_Energy", "CDR_Negative_Patch_Energy_gt30", 
                      "CDR_Negative_Patch_Energy_gt50", "CDR_Positive_Patch_Energy", "CDR_Positive_Patch_Energy_gt30", "CDR_Positive_Patch_Energy_gt50", "FRH_AggScore", 
                      "FRH_Aggrescan_a4v", "FRH_Aggrescan_a4v_pos", "FRH_Formal_Charge", "FRH_Hydrophobic_Patch_Energy", "FRH_Hydrophobic_Patch_Energy_gt15", 
                      "FRH_Hydrophobic_Patch_Energy_gt30", "FRH_Negative_Patch_Energy", "FRH_Negative_Patch_Energy_gt30", "FRH_Negative_Patch_Energy_gt50", 
                      "FRH_Positive_Patch_Energy", "FRH_Positive_Patch_Energy_gt30", "FRH_Positive_Patch_Energy_gt50", "FRL_AggScore", "FRL_Formal_Charge", 
                      "FRL_Hydrophobic_Patch_Energy", "FRL_Hydrophobic_Patch_Energy_gt15", "FRL_Hydrophobic_Patch_Energy_gt30", "FRL_Negative_Patch_Energy", 
                      "FRL_Negative_Patch_Energy_gt30", "FRL_Negative_Patch_Energy_gt50", "FRL_Positive_Patch_Energy", "FRL_Positive_Patch_Energy_gt30", 
                      "FRL_Positive_Patch_Energy_gt50", "FR_AggScore", "FR_Aggrescan_a4v", "FR_Aggrescan_a4v_pos", "FR_Formal_Charge", "FR_Hydrophobic_Patch_Energy", 
                      "FR_Hydrophobic_Patch_Energy_gt15", "FR_Hydrophobic_Patch_Energy_gt30", "FR_Negative_Patch_Energy", "FR_Negative_Patch_Energy_gt30", 
                      "FR_Negative_Patch_Energy_gt50", "FR_Positive_Patch_Energy", "FR_Positive_Patch_Energy_gt30", "FR_Positive_Patch_Energy_gt50", "Formal_Charge_eV", 
                      "H1_AggScore", "H1_Formal_Charge", "H1_Hydrophobic_Patch_Energy", "H1_Hydrophobic_Patch_Energy_gt15", "H1_Hydrophobic_Patch_Energy_gt30", 
                      "H1_Negative_Patch_Energy", "H1_Negative_Patch_Energy_gt30", "H1_Negative_Patch_Energy_gt50", "H1_Positive_Patch_Energy", "H1_Positive_Patch_Energy_gt30", 
                      "H1_Positive_Patch_Energy_gt50", "H2_AggScore", "H2_Formal_Charge", "H2_Hydrophobic_Patch_Energy", "H2_Hydrophobic_Patch_Energy_gt15", 
                      "H2_Hydrophobic_Patch_Energy_gt30", "H2_Negative_Patch_Energy", "H2_Negative_Patch_Energy_gt30", "H2_Negative_Patch_Energy_gt50", 
                      "H2_Positive_Patch_Energy", "H2_Positive_Patch_Energy_gt30", "H2_Positive_Patch_Energy_gt50", "H3_AggScore", "H3_Formal_Charge", 
                      "H3_Hydrophobic_Patch_Energy", "H3_Hydrophobic_Patch_Energy_gt15", "H3_Hydrophobic_Patch_Energy_gt30", "H3_Negative_Patch_Energy", 
                      "H3_Negative_Patch_Energy_gt30", "H3_Negative_Patch_Energy_gt50", "H3_Positive_Patch_Energy", "H3_Positive_Patch_Energy_gt30", 
                      "H3_Positive_Patch_Energy_gt50", "HFR1_AggScore", "HFR1_Formal_Charge", "HFR1_Hydrophobic_Patch_Energy", "HFR1_Hydrophobic_Patch_Energy_gt15", 
                      "HFR1_Hydrophobic_Patch_Energy_gt30", "HFR1_Negative_Patch_Energy", "HFR1_Negative_Patch_Energy_gt30", "HFR1_Negative_Patch_Energy_gt50", 
                      "HFR1_Positive_Patch_Energy", "HFR1_Positive_Patch_Energy_gt30", "HFR1_Positive_Patch_Energy_gt50", "HFR2_AggScore", "HFR2_Formal_Charge", 
                      "HFR2_Hydrophobic_Patch_Energy", "HFR2_Hydrophobic_Patch_Energy_gt15", "HFR2_Hydrophobic_Patch_Energy_gt30", "HFR2_Negative_Patch_Energy", 
                      "HFR2_Negative_Patch_Energy_gt30", "HFR2_Negative_Patch_Energy_gt50", "HFR2_Positive_Patch_Energy", "HFR2_Positive_Patch_Energy_gt30", 
                      "HFR2_Positive_Patch_Energy_gt50", "HFR3_AggScore", "HFR3_Formal_Charge", "HFR3_Hydrophobic_Patch_Energy", "HFR3_Hydrophobic_Patch_Energy_gt15", 
                      "HFR3_Hydrophobic_Patch_Energy_gt30", "HFR3_Negative_Patch_Energy", "HFR3_Negative_Patch_Energy_gt30", "HFR3_Negative_Patch_Energy_gt50", 
                      "HFR3_Positive_Patch_Energy", "HFR3_Positive_Patch_Energy_gt30", "HFR3_Positive_Patch_Energy_gt50", "HFR4_AggScore", "HFR4_Formal_Charge", 
                      "HFR4_Greasy_SASA", "HFR4_Hydrophobic_Patch_Energy", "HFR4_Hydrophobic_Patch_Energy_gt15", "HFR4_Hydrophobic_Patch_Energy_gt30", 
                      "HFR4_Negative_Patch_Energy", "HFR4_Negative_Patch_Energy_gt30", "HFR4_Negative_Patch_Energy_gt50", "HFR4_Positive_Patch_Energy", 
                      "HFR4_Positive_Patch_Energy_gt30", "HFR4_Positive_Patch_Energy_gt50", "L1_AggScore", "L1_Formal_Charge", "L1_Hydrophobic_Patch_Energy", 
                      "L1_Hydrophobic_Patch_Energy_gt15", "L1_Hydrophobic_Patch_Energy_gt30", "L1_Negative_Patch_Energy", "L1_Negative_Patch_Energy_gt30", 
                      "L1_Negative_Patch_Energy_gt50", "L1_Positive_Patch_Energy", "L1_Positive_Patch_Energy_gt30", "L1_Positive_Patch_Energy_gt50", "L2_AggScore", 
                      "L2_Formal_Charge", "L2_Hydrophobic_Patch_Energy", "L2_Hydrophobic_Patch_Energy_gt15", "L2_Hydrophobic_Patch_Energy_gt30", "L2_Negative_Patch_Energy", 
                      "L2_Negative_Patch_Energy_gt30", "L2_Negative_Patch_Energy_gt50", "L2_Positive_Patch_Energy", "L2_Positive_Patch_Energy_gt30", "L2_Positive_Patch_Energy_gt50", 
                      "L3_AggScore", "L3_Formal_Charge", "L3_Hydrophobic_Patch_Energy", "L3_Hydrophobic_Patch_Energy_gt15", "L3_Hydrophobic_Patch_Energy_gt30", 
                      "L3_Negative_Patch_Energy", "L3_Negative_Patch_Energy_gt30", "L3_Negative_Patch_Energy_gt50", "L3_Positive_Patch_Energy", "L3_Positive_Patch_Energy_gt30", 
                      "L3_Positive_Patch_Energy_gt50", "LFR1_AggScore", "LFR1_Formal_Charge", "LFR1_Hydrophobic_Patch_Energy", "LFR1_Hydrophobic_Patch_Energy_gt15", 
                      "LFR1_Hydrophobic_Patch_Energy_gt30", "LFR1_Negative_Patch_Energy", "LFR1_Negative_Patch_Energy_gt30", "LFR1_Negative_Patch_Energy_gt50", 
                      "LFR1_Positive_Patch_Energy", "LFR1_Positive_Patch_Energy_gt30", "LFR1_Positive_Patch_Energy_gt50", "LFR2_AggScore", "LFR2_Formal_Charge", 
                      "LFR2_Hydrophobic_Patch_Energy", "LFR2_Hydrophobic_Patch_Energy_gt15", "LFR2_Hydrophobic_Patch_Energy_gt30", "LFR2_Negative_Patch_Energy", 
                      "LFR2_Negative_Patch_Energy_gt30", "LFR2_Negative_Patch_Energy_gt50", "LFR2_Positive_Patch_Energy", "LFR2_Positive_Patch_Energy_gt30", 
                      "LFR2_Positive_Patch_Energy_gt50", "LFR3_AggScore", "LFR3_Formal_Charge", "LFR3_Hydrophobic_Patch_Energy", "LFR3_Hydrophobic_Patch_Energy_gt15", 
                      "LFR3_Hydrophobic_Patch_Energy_gt30", "LFR3_Negative_Patch_Energy", "LFR3_Negative_Patch_Energy_gt30", "LFR3_Negative_Patch_Energy_gt50", 
                      "LFR3_Positive_Patch_Energy", "LFR3_Positive_Patch_Energy_gt30", "LFR3_Positive_Patch_Energy_gt50", "LFR4_AggScore", "LFR4_Formal_Charge", 
                      "LFR4_Hydrophobic_Patch_Energy", "LFR4_Hydrophobic_Patch_Energy_gt15", "LFR4_Hydrophobic_Patch_Energy_gt30", "LFR4_Negative_Patch_Energy", 
                      "LFR4_Negative_Patch_Energy_gt30", "LFR4_Negative_Patch_Energy_gt50", "LFR4_Positive_Patch_Energy", "LFR4_Positive_Patch_Energy_gt30", 
                      "LFR4_Positive_Patch_Energy_gt50", "Max_Size_Hyd_Patches", "Max_Size_Neg_Patches", "Max_Size_Pos_Patches", "Sum_Size_Hyd_Patches", "Sum_Size_Neg_Patches", 
                      "Sum_Size_Pos_Patches", "VH_AggScore", "VH_Fv_Formal_Charge", "VH_Fv_Hydrophobic_Patch_Energy", "VH_Fv_Hydrophobic_Patch_Energy_gt15", 
                      "VH_Fv_Hydrophobic_Patch_Energy_gt30", "VH_Fv_Negative_Patch_Energy", "VH_Fv_Negative_Patch_Energy_gt30", "VH_Fv_Negative_Patch_Energy_gt50", 
                      "VH_Fv_Positive_Patch_Energy", "VH_Fv_Positive_Patch_Energy_gt30", "VH_Fv_Positive_Patch_Energy_gt50", "VL_AggScore", "VL_Formal_Charge", 
                      "VL_Fv_AggScore", "VL_Fv_Formal_Charge", "VL_Fv_Hydrophobic_Patch_Energy", "VL_Fv_Hydrophobic_Patch_Energy_gt15", "VL_Fv_Hydrophobic_Patch_Energy_gt30", 
                      "VL_Fv_Negative_Patch_Energy", "VL_Fv_Negative_Patch_Energy_gt30", "VL_Fv_Negative_Patch_Energy_gt50", "VL_Fv_Positive_Patch_Energy", 
                      "VL_Fv_Positive_Patch_Energy_gt30", "VL_Fv_Positive_Patch_Energy_gt50", "pI_PROPKA_based", "pI_model_pKa_based"]  # 替换为您需要的列

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
