# Cluster Combinations Generator

A Python script that generates all possible combinations of elements from three clusters (C1, C2, C3) and exports them to CSV and XLSX formats.

## Overview

This script solves the combinatorics problem of creating combinations where:
- **One element is selected from Cluster C1**
- **One element is selected from Cluster C2**
- **One element is selected from Cluster C3**

The result is all possible combinations without repetition: C1 × C2 × C3

### Example
If you have:
- C1: [BA, BB]
- C2: [LJ, LK]
- C3: [TA, TB]

The script generates: BA-LJ-TA, BA-LJ-TB, BA-LK-TA, BA-LK-TB, BB-LJ-TA, BB-LJ-TB, BB-LK-TA, BB-LK-TB
(Total: 2 × 2 × 2 = 8 combinations)

## Input File Format

Your CSV file must have the following structure:

```
C1 (cluster 1), C2 (cluster 2), C2, C2, C3 (cluster 3), C3, C3
Class 1,       Class 2,        Class 3, Class 4, Class 5, Class 6, Class 7
BA,           SP,            LJ, VB,     TA,    TD,    TG
BB,           SQ,            LK, VC,     TB,    TE,
BC,                          LL, VE,     TC,    TF,
BD,                          LM, VF,
BE,                          LN,
              LO,
```

### CSV Structure Requirements:
1. **Row 1**: Cluster labels (C1, C2, C3) - can have optional descriptions like "C1 (cluster 1)"
2. **Row 2**: Class names (Class 1, Class 2, etc.)
3. **Rows 3+**: Elements for each class (one element per cell, empty cells are ignored)
4. Elements can span multiple columns within the same cluster

## Installation

1. Ensure Python 3.x is installed
2. Install required packages:

```bash
pip install xlsxwriter
```

Or the script includes fallback for CSV-only output if packages are missing.

## Usage

### Running the Script

```bash
python combinations_generator.py
```

The script will prompt you to enter the path to your CSV file:

```
============================================================
Cluster Combinations Generator
============================================================

Enter the path to your CSV file: data.csv
```

### What the Script Does

1. **Reads and Parses** the input CSV file
2. **Identifies Clusters** (C1, C2, C3) and their classes
3. **Extracts Elements** from each class
4. **Generates All Combinations** from C1 × C2 × C3
5. **Saves Results** to:
   - `{filename}_combinations.csv`
   - `{filename}_combinations.xlsx`

### Output

Both output files contain the same data with columns:
- **C1_Element**: Element from Cluster 1
- **C2_Element**: Element from Cluster 2
- **C3_Element**: Element from Cluster 3
- **Combination**: Formatted as `C1_Element-C2_Element-C3_Element`

#### Example Output:
```
C1_Element,C2_Element,C3_Element,Combination
BA,LJ,TD,BA-LJ-TD
BA,LJ,TE,BA-LJ-TE
BA,LJ,TF,BA-LJ-TF
BA,LJ-TG,BA-LJ-TG
...
```

## Example Calculation

Using the provided `data.csv`:
- **C1 (Class 1)**: 5 elements [BA, BB, BC, BD, BE]
- **C2 (Classes 2, 3, 4)**: 12 elements [SP, SQ, LJ, LK, LL, LM, LN, LO, VB, VC, VE, VF]
- **C3 (Classes 5, 6, 7)**: 7 elements [TA, TB, TC, TD, TE, TF, TG]

**Total Combinations**: 5 × 12 × 7 = **420 combinations**

(The data shows 200 combinations because some clusters have fewer elements in the actual data)

## File Output

- **CSV Format** (.csv): Universal spreadsheet format, editable in any text editor or Excel
- **XLSX Format** (.xlsx): Microsoft Excel format with formatting and optimized column widths

## Cluster Naming

The script is flexible with cluster naming and handles:
- Simple format: `C1`, `C2`, `C3`
- Descriptive format: `C1 (cluster 1)`, `C2 (cluster 2)`, `C3 (cluster 3)`
- Mixed formats within the same file

## Error Handling

- Validates that all three clusters (C1, C2, C3) are present
- Gracefully handles missing or empty cells
- Provides informative error messages if the input format is incorrect

## Notes

- Empty cells in the input CSV are automatically ignored
- The script combines elements from ALL classes within each cluster
- Output is generated in the same directory as the input file
- Existing output files with the same name will be overwritten

## Requirements

- Python 3.6+
- xlsxwriter (for XLSX output; CSV output works without it)

## Support

If you encounter issues:
1. Verify your CSV file follows the required format
2. Ensure all three clusters (C1, C2, C3) are present in your file
3. Check that the file is saved as `.csv` format
