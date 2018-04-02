# coord2gene
Coord2gene is a simple Python script that converts genomic coordinates to gene names.  Useful for mapping SNPs to genes, and optimized for speed.

# Purpose of Software
This script will read genomic coordinates (chromosome and locus) from an Excel file (xlsx) and look up the gene name using PyEnsembl.  Using a local copy of the human genome, it can run very fast (300k+ coordinates in about 10 minutes).

# Dependencies
1. Python 3

2. openpyxl

3. PyEnsembl
   - A locally cached copy of a genome, installed using PyEnsembl
    
    

# Installation
Windows:  Currently, this script will not run on Windows due to limitations of 3rd party packages on which it depends.
