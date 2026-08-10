#!/usr/bin/env python3
import numpy as np
import pandas as pd
import os


mols=np.loadtxt("file2.dat",dtype=str)
#print(mols)
#exit(0)
dfs=[]
for i in mols:
  print(i)
  file="AF3_07082026/s%s/protein_descriptors.csv"%(i)
  if os.path.exists(file):
    print("yes")
    df=pd.read_csv(file)
    df['Name'] = "%s"%(i)
    dfs.append(df)
combined_df=pd.concat(dfs,ignore_index=True)
combined_df.to_csv("protein_descriptors.csv",index=False)
