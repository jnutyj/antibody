#!/usr/bin/env python3
import pymol2
import glob

def cif2pdb(fname):
    with pymol2.PyMOL() as pymol:
    	pymol.cmd.load("%s.cif"%(fname),"mol_obj")
    	pymol.cmd.save("%s.pdb"%(fname), selection="mol_obj")



fls=glob.glob("*.cif")

for i in fls:
   #print(i)
   filename=i.split(".")[0]
   #print(filename)
   cif2pdb(filename)
