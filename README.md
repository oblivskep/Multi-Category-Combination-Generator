# Multi-Category Combination Generator

A structured tool for generating all possible combinations across multiple input categories (e.g., product variants, configuration matrices, SKU expansions).

This Python engine generates a complete cross-product (C1 × C2 × C3) from categorized CSV inputs and exports results to CSV and Excel formats.

---

## Overview

This project generates all possible structured combinations across three defined categories.

For each output row:

- One element is selected from Category C1  
- One element is selected from Category C2  
- One element is selected from Category C3  

The result is a complete structured cross-product:

Total Combinations = |C1| × |C2| × |C3|

Manual configuration of multi-category data often leads to missing combinations, inconsistent structures, and downstream workflow errors. This generator ensures complete coverage, structured output, and scalable expansion.

---

## Visual Schema

Input Categories:

    C1:  A1   A2   A3
    C2:  B1   B2
    C3:  C1   C2   C3

Cross-Product Generation:

             ┌───────────────┐
             │   C1 Elements │
             └───────────────┘
                      │
                      ▼
             ┌───────────────┐
             │   C2 Elements │
             └───────────────┘
                      │
                      ▼
             ┌───────────────┐
             │   C3 Elements │
             └───────────────┘
                      │
                      ▼
        ┌─────────────────────────────┐
        │     Generated Output Rows   │
        │  A1-B1-C1                   │
        │  A1-B1-C2                   │
        │  A1-B2-C1                   │
        │  ...                        │
        └─────────────────────────────┘

---

## Example

If the input contains:

- C1: [A, B]  
- C2: [X, Y]  
- C3: [1, 2]  

The generator produces:

A-X-1  
A-X-2  
A-Y-1  
A-Y-2  
B-X-1  
B-X-2  
B-Y-1  
B-Y-2  

Total combinations: 2 × 2 × 2 = 8

---

## Input File Format

The input must be a CSV file structured as follows:

C1 (optional description), C2 (optional description), C2, C3, C3  
Class 1, Class 2, Class 3, Class 4, Class 5  
A1, B1, B2, C1, C2  
A2, B3, B4, C3,  

### Structure Rules

1. Row 1 → Category labels (C1, C2, C3 — descriptions allowed)
2. Row 2 → Optional class metadata
3. Row 3+ → Elements (one element per cell)
4. Empty cells are ignored automatically
5. A category may span multiple columns
6. All columns belonging to the same category are grouped before generating combinations

For a visual example of how this spreadsheet should look when opened in Excel or LibreOffice, see:

**example_input.png**

---

## What the Script Does

1. Reads and parses structured CSV input  
2. Identifies category groupings (C1, C2, C3)  
3. Extracts valid elements from each category  
4. Generates full cross-product combinations  
5. Exports results to:
   - `{filename}_combinations.csv`
   - `{filename}_combinations.xlsx`

---

## Output Structure

Each output file contains:

- C1_Element  
- C2_Element  
- C3_Element  
- Combination (formatted as C1-C2-C3)

Example:

C1_Element,C2_Element,C3_Element,Combination  
A,X,1,A-X-1  
A,X,2,A-X-2  

---

## Use Cases

This engine supports:

- Product variant generation (size × color × model)  
- SKU combination expansion  
- Configuration matrices  
- Inventory combinations  
- Test case permutation generation  
- Bundle generation systems  
- Structured dataset expansion  
- Manufacturing option modeling  

---

## Installation

Requires Python 3.6+

Install dependency:

    pip install xlsxwriter

Or:

    pip install -r requirements.txt

---

## Running the Script

    python multi_category_combination_generator.py

You will be prompted to enter the path to your CSV file.

---

## Features

- Flexible category naming  
- Multi-column category support  
- Automatic empty-cell filtering  
- Structured CSV and Excel export  
- Input validation and error handling  

---

## Requirements

- Python 3.6+
- xlsxwriter (for XLSX export)

---

## License

MIT License

---

## Author

Developed as part of an automation and structured data workflow portfolio.