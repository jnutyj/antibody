#!/usr/bin/env python3
import pandas as pd
### this file is based on schrodinger residue scanning result to print out mutation sequence with affinity and stability less than 0.

def mutate_sequence(original_seq, mutations):
    """
    Mutates a given sequence based on a list of mutation strings (e.g., '52(ASN->ALA)').
    """
    # Convert sequence to a list so we can mutate individual positions
    seq_list = list(original_seq)
    aa_3_to_1 = {
        'ALA': 'A', 'ARG': 'R', 'ASN': 'N', 'ASP': 'D', 'CYS': 'C',
        'GLU': 'E', 'GLN': 'Q', 'GLY': 'G', 'HIS': 'H', 'ILE': 'I',
        'LEU': 'L', 'LYS': 'K', 'MET': 'M', 'PHE': 'F', 'PRO': 'P',
        'SER': 'S', 'THR': 'T', 'TRP': 'W', 'TYR': 'Y', 'VAL': 'V'
    }
    
    for mutation in mutations:
        # Parse the mutation string: extract position, original, and new amino acid
        parts = mutation.split('(')
        position = int(parts[0])
        original_3aa, new_3aa = parts[1].replace(')', '').split('->')
        original_aa=aa_3_to_1[original_3aa]
        new_aa = aa_3_to_1[new_3aa]
        # Python uses 0-based indexing, so subtract 1 from the 1-based position
        idx = position - 1 
        
        # Verify the original amino acid matches before mutating to prevent errors
        if seq_list[idx] == original_aa:
            seq_list[idx] = new_aa
            #print(f"Success: Mutated {original_aa} to {new_aa} at position {position}")
        else:
            print(f"Warning: Expected {original_aa} at position {position}, but found {seq_list[idx]}. Skipping.")

    return "".join(seq_list)

#mutation=["52(ASN->ALA)","54(SER->ALA)"]
#seq="QVELVESGGGLVQPGGSLRLSCAASGFTFSSYAMSWVRQAPGKGLEWVSAINASGTRTYYADSVKGRFTISRDNSKNTLYLQMNSLRAEDTAVYYCARGKGNTHKPYGYVRYFDVWGQGTLVTVSS"

#print(mutate_sequence(seq,mutation))

#df=pd.read_csv("residue_scanning_6-results.csv")

#filter_df=df[(df['delta Affinity'] < 0) & (df['delta Stability'] < 0)]
filter_df=pd.read_csv("residue_mutation-results.csv")
#print(filter_df['delta Stability'].tolist())
org_seq="QVQLVESGGGVVQPGRSLRLDCKASGITFSNSGMHWVRQAPGKGLEWVAVIWYDGSKRYYADSVKGRFTISRDNSKNTLFLQMNSLRAEDTAVYYCATNDDYWGQGTLVTVSS"
i=1
for words in filter_df['Mutations']:
    #print(words.replace("A:","").split(","))
    mutation=words.replace("B:","").split(",")
    seq=mutate_sequence(org_seq,mutation)
    affinity= filter_df.loc[filter_df['Mutations'] == words, "delta Affinity"].to_numpy()[0]
    stability=filter_df.loc[filter_df['Mutations'] == words, "delta Stability"].to_numpy()[0]
    #print(seq,affinity,stability)
    #print(stability)
    #print(">m%i.H"%(i))
    print(seq)
    #print(">m%i.L"%(i))
    #print("snivolumab.H.m%i"%(i))
    #print("DIVLTQSPATLSLSPGERATLSCRASQSVSSSYLAWYQQKPGQAPRLLIYGASSRATGVPARFSGSGSGTDFTLTISSLEPEDFATYYCLQIYNMPITFGQGTKVEIK")
    #print("LDSPDRPWNPPTFSPALLVVTEGDNATFTCSFSNTSESFVLNWYRMSPSNQTDKLAAFPEDRSQPGQDCRFRVTQLPNGRDFHMSVVRARRNDSGTYLCGAISLAPKAQIKESLRAELRVTERRAEVPTAH:%s:EIVLTQSPATLSLSPGERATLSCRASQSVSSYLAWYQQKPGQAPRLLIYDASNRATGIPARFSGSGSGTDFTLTISSLEPEDFAVYYCQQSSNWPRTFGQGTKVEIK"%(seq))
    #print("LDSPDRPWNPPTFSPALLVVTEGDNATFTCSFSNTSESFVLNWYRMSPSNQTDKLAAFPEDRSQPGQDCRFRVTQLPNGRDFHMSVVRARRNDSGTYLCGAISLAPKAQIKESLRAEL:%s:EIVLTQSPATLSLSPGERATLSCRASQSVSSYLAWYQQKPGQAPRLLIYDASNRATGIPARFSGSGSGTDFTLTISSLEPEDFAVYYCQQSSNWPRTFGQGTKVEIK"%(seq)) 
    #print("%s:EIVLTQSPATLSLSPGERATLSCRASQSVSSYLAWYQQKPGQAPRLLIYDASNRATGIPARFSGSGSGTDFTLTISSLEPEDFAVYYCQQSSNWPRTFGQGTKVEIK"%(seq))
    i=i+1
