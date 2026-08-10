import itertools
import os
from collections import defaultdict

def generate_mutated_sequences(original_seq, mutations_dict, max_sites):
    """
    生成组合突变的氨基酸序列，剔除原始序列，并限制同时突变的位点数量
    
    参数:
    original_seq (str): 原始氨基酸序列
    mutations_dict (dict): 突变位点字典 {位置: [突变氨基酸列表]}
    max_sites (int): 最大同时突变位点数
    
    返回:
    list: 元组列表，格式为(突变ID, 突变序列)
    """
    # 预处理：移除原始氨基酸，确保只生成实际突变的序列
    filtered_mutations = {}
    for pos, options in mutations_dict.items():
        orig_aa = original_seq[pos-1]
        # 过滤掉原始氨基酸和无效位置
        if pos > len(original_seq):
            continue
        filtered_options = [aa for aa in options if aa != orig_aa and aa in "ACDEFGHIKLMNPQRSTVWY"]
        if filtered_options:
            filtered_mutations[pos] = filtered_options
    
    positions = list(filtered_mutations.keys())
    results = []
    
    # 生成1到max_sites个位点的所有组合
    #for num_sites in range(1, max_sites + 1):
    for num_sites in [max_sites]:
        for site_combo in itertools.combinations(positions, num_sites):
            # 获取当前位点组合的所有突变组合
            mutations_at_sites = [filtered_mutations[pos] for pos in site_combo]
            
            for aa_combo in itertools.product(*mutations_at_sites):
                mutated_seq = list(original_seq)
                mut_id_parts = []
                
                # 应用突变
                for pos, mut_aa in zip(site_combo, aa_combo):
                    orig_aa = original_seq[pos-1]
                    mutated_seq[pos-1] = mut_aa
                    mut_id_parts.append(f"{orig_aa}{pos}{mut_aa}")
                
                # 构建结果
                mutated_seq_str = ''.join(mutated_seq)
                mut_id = "_".join(mut_id_parts)
                results.append((f">mut_{mut_id}", mutated_seq_str))
    
    print(f"共生成 {len(results)} 条突变序列（已剔除原始序列）")
    return results

def save_sequences_to_txt(mutated_sequences, filename="mutated_sequences.txt"):
    """将突变序列保存到TXT文件"""
    with open(filename, "w") as f:
        for header, seq in mutated_sequences:
            f.write(f"{header}\n{seq}\n")
            print(seq)
    print(f"结果已保存至: {os.path.abspath(filename)}")

if __name__ == "__main__":
    # ============== 配置参数 ==============
    original_sequence = "QVELVESGGGLVQPGGSLRLSCAASGFTFSSYAMSWVRQAPGKGLEWVSAINREGTRTYYADSVKGRFTISRDNSKNTLYLQMNSLRAEDTAVYYCARGKGNTHKPYGYVRYFDVWGQGTLVTVSS"  # 替换为您的氨基酸序列
    
    # 格式: {位置: [突变选项列表]}
    mutations_dict = {
        100: ['A','Q','E','I','L','M','F','S','T','W','Y','V'],    # 位置1可突变为A或V（原始M已过滤）
        105: ['A', 'N', 'D', 'Q','E','G','I','L','M','F','S','T','W','Y','V'],    # 位置5可突变为S或D（原始T已过滤）
        111: ['Q','M']  # 位置7可突变为L/F/W（原始E已过滤）
        
    }
    
    max_sites = 2  # 最多同时突变2个位点
    output_filename = "mutated_sequences.txt"
    # =====================================
    
    # 生成突变序列（自动剔除原始序列）
    mutated_sequences = generate_mutated_sequences(
        original_sequence,
        mutations_dict,
        max_sites
    )
    
    # 保存结果到文件
    save_sequences_to_txt(mutated_sequences, output_filename)
