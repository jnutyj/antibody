#!/usr/bin/env python3
import numpy as np
def format_sequence_mutations(wild_type: str, mutant: str, chain: str = "B") -> str:
    """Compares two pre-aligned sequences and prints mutations in the format:
    
    Chain:Index->THREE_LETTER_CODE
    """
    if len(wild_type) != len(mutant):
        raise ValueError("Sequences must be pre-aligned and of equal length.")

    # Dictionary to convert 1-letter codes to 3-letter uppercase codes
    aa_mapping = {
        'A': 'ALA', 'R': 'ARG', 'N': 'ASN', 'D': 'ASP', 'C': 'CYS',
        'E': 'GLU', 'Q': 'GLN', 'G': 'GLY', 'H': 'HIS', 'I': 'ILE',
        'L': 'LEU', 'K': 'LYS', 'M': 'MET', 'F': 'PHE', 'P': 'PRO',
        'S': 'SER', 'T': 'THR', 'W': 'TRP', 'Y': 'TYR', 'V': 'VAL',
        '-': 'GAP' # Handles deletions if present
    }

    mutation_list = []

    for index, (wt_res, mut_res) in enumerate(zip(wild_type, mutant), start=1):
        if wt_res != mut_res:
            # Get the 3-letter code (defaults to original letter if not in standard dict)
            three_letter_mut = aa_mapping.get(mut_res.upper(), mut_res.upper())
            
            # Format: B:102->GLN
            formatted_string = f"{chain}:{index}->{three_letter_mut}"
            mutation_list.append(formatted_string)
            
    # Join all mutations with a single space
    return " ".join(mutation_list)


wt='QVQLVESGGGVVQPGRSLRLDCKASGITFSNSGMHWVRQAPGKGLEWVAVIWYDGSKRYYADSVKGRFTISRDNSKNTLFLQMNSLRAEDTAVYYCATNDDYWGQGTLVTVSS'
seqs=np.loadtxt("file.txt",dtype="str")
#print(seqs)
for i in seqs:
   print(format_sequence_mutations(wt, i, "B"))
