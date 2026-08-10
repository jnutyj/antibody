#!/bin/bash

#prepwizard fold_nivolumab_pd1_model_0.pdb proteinprep_out.mae -fillsidechains -disulfides -assign_all_residues -rehtreat -max_states 1 -antibody_cdr_scheme Kabat -samplewater -f S-OPLS
#$SCHRODINGER/run residue_scanning_backend.py proteinprep_out.mae -muts_file mutation.txt -ligand_asl "chain A"
#exit

#for i in $(cat file3.txt);do
#    echo "${i}:DIVLTQSPATLSLSPGERATLSCRASQSVSSSYLAWYQQKPGQAPRLLIYGASSRATGVPARFSGSGSGTDFTLTISSLEPEDFATYYCLQIYNMPITFGQGTKVEIK:DAEFRHDSGYE"
#done
#exit

folder="AF3_07082026"
#name="VBB102_Diffab_"
#for i in $(seq 1 1 99);do
for i in $(cat file2.dat);do
    if [ ! -d ${folder}/s${i}/ ];then
       mkdir -p ${folder}/s${i}/
       #cp AF2_orig/${i}_relaxed_rank_001_*.pdb AF2
    fi
    cd ${folder}/s${i}/
    #ls 
    #prepwizard ../s${i}_model.cif proteinprep_out.mae -fillsidechains -disulfides -assign_all_residues -rehtreat -max_states 1 -antibody_cdr_scheme Kabat -samplewater -f S-OPLS
    #structalign proteinprep_out.mae ../../crystal_out.mae -asl 'chain B' -asl_mobile 'chain B'> align_crystal.dat    
    #rmsd=$(tail -n 1 align_crystal.dat| awk '{print $2}')
    #echo "${i} ${rmsd}"
    cd ..
    #$SCHRODINGER/run -FROM psp calc_protein_descriptors.py -i s${i} -o s${i} -p 7.0 -s Kabat
    cd ../
done

#python3 extract.py
python3 ../descriptors_FVpickup271_csvtoexcel.py
python3 ../FV_DA3_analysis_271.py
