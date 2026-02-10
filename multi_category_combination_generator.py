import csv
import os
from itertools import product
from pathlib import Path


def read_and_parse_csv(filepath):
    """
    Read CSV file and extract clusters, classes, and elements.
    """
    with open(filepath, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        rows = list(reader)
    
    # Extract cluster information from first row
    clusters = rows[0]
    
    # Extract class names from second row
    class_names = rows[1]
    
    # Extract elements (rows 2 onwards)
    elements_data = rows[2:]
    
    # Parse clusters and their classes
    cluster_structure = {}
    
    for col_idx, (cluster, class_name) in enumerate(zip(clusters, class_names)):
        if not cluster or cluster.strip() == '':
            continue
        
        cluster = cluster.strip()
        class_name = class_name.strip() if class_name else f"Class_{col_idx}"
        
        if cluster not in cluster_structure:
            cluster_structure[cluster] = {}
        
        # Get elements for this class (non-empty cells in this column)
        elements = []
        for row in elements_data:
            if col_idx < len(row):
                elem = row[col_idx].strip()
                if elem:
                    elements.append(elem)
        
        cluster_structure[cluster][class_name] = elements
    
    return cluster_structure


def generate_combinations(cluster_structure):
    """
    Generate all combinations: one element from C1 × C2 × C3
    """
    # Extract elements from each cluster - aggregate all clusters with the same key
    cluster_aggregates = {'C1': [], 'C2': [], 'C3': []}
    
    for cluster_name in cluster_structure.keys():
        classes_in_cluster = cluster_structure[cluster_name]
        elements = []
        
        for class_name, class_elements in sorted(classes_in_cluster.items()):
            elements.extend(class_elements)
        
        # Match cluster names - handle both "C1" and "C1 (cluster 1)" formats
        cluster_key = cluster_name.strip().split()[0].upper()
        
        # Aggregate elements from all clusters with the same key
        if cluster_key in cluster_aggregates:
            cluster_aggregates[cluster_key].extend(elements)
    
    # Validate that all clusters exist and have elements
    for key in ['C1', 'C2', 'C3']:
        if not cluster_aggregates[key]:
            raise ValueError(f"Cluster {key} has no elements. Found clusters: {list(cluster_structure.keys())}")
    
    # Generate all combinations
    combinations = []
    for c1_elem in cluster_aggregates['C1']:
        for c2_elem in cluster_aggregates['C2']:
            for c3_elem in cluster_aggregates['C3']:
                combinations.append([c1_elem, c2_elem, c3_elem, f"{c1_elem}-{c2_elem}-{c3_elem}"])
    
    return combinations


def save_combinations_csv(combinations, output_filename):
    """
    Save combinations to CSV format
    """
    base_name = Path(output_filename).stem
    output_dir = Path(output_filename).parent
    csv_path = output_dir / f"{base_name}_combinations.csv"
    
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['C1_Element', 'C2_Element', 'C3_Element', 'Combination'])
        for row in combinations:
            writer.writerow(row)
    
    print(f"✓ CSV file saved: {csv_path}")
    return csv_path


def save_combinations_xlsx(combinations, output_filename):
    """
    Save combinations to XLSX format using xlsxwriter as alternative
    """
    try:
        import xlsxwriter
        
        base_name = Path(output_filename).stem
        output_dir = Path(output_filename).parent
        xlsx_path = output_dir / f"{base_name}_combinations.xlsx"
        
        # Create a new XLSX file
        workbook = xlsxwriter.Workbook(str(xlsx_path))
        worksheet = workbook.add_worksheet('Combinations')
        
        # Create header format
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        # Write headers
        headers = ['C1_Element', 'C2_Element', 'C3_Element', 'Combination']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
        
        # Write data
        for row_idx, row in enumerate(combinations, start=1):
            for col_idx, value in enumerate(row):
                worksheet.write(row_idx, col_idx, value)
        
        # Set column widths
        worksheet.set_column(0, len(headers)-1, 15)
        
        workbook.close()
        print(f"✓ XLSX file saved: {xlsx_path}")
        return xlsx_path
    
    except ImportError:
        print("⚠ Could not create XLSX file (xlsxwriter not available)")
        print("  But CSV file has been created successfully")
        return None
    except Exception as e:
        print(f"⚠ Could not create XLSX file: {str(e)}")
        print("  But CSV file has been created successfully")
        return None


def main():
    """
    Main function to orchestrate the combination generation
    """
    print("=" * 60)
    print("Multi-Category Combination Generator")
    print("=" * 60)
    
    # Get input file from user
    while True:
        input_file = input("\nEnter the path to your CSV file: ").strip()
        
        if not input_file:
            print("Error: Please provide a valid file path")
            continue
        
        if not os.path.exists(input_file):
            print(f"Error: File '{input_file}' not found")
            continue
        
        if not input_file.lower().endswith('.csv'):
            print("Error: File must be a CSV file (.csv)")
            continue
        
        break
    
    try:
        print(f"\nReading file: {input_file}")
        cluster_structure = read_and_parse_csv(input_file)
        
        # Display parsed structure
        print("\nParsed Cluster Structure:")
        print("-" * 40)
        for cluster_name in sorted(cluster_structure.keys()):
            print(f"\n{cluster_name}:")
            for class_name, elements in sorted(cluster_structure[cluster_name].items()):
                print(f"  {class_name}: {len(elements)} elements - {elements}")
        
        # Generate combinations
        print("\n\nGenerating combinations...")
        combinations = generate_combinations(cluster_structure)
        
        print(f"Total combinations generated: {len(combinations)}")
        
        # Save results
        print("\nSaving results...")
        csv_path = save_combinations_csv(combinations, input_file)
        xlsx_path = save_combinations_xlsx(combinations, input_file)
        
        print("\n" + "=" * 60)
        print("Process completed successfully!")
        print("=" * 60)
        print(f"\nOutput files:")
        print(f"  CSV:  {csv_path}")
        if xlsx_path:
            print(f"  XLSX: {xlsx_path}")
        
        # Show sample of combinations
        print(f"\nSample of generated combinations (first 5):")
        for i, combo in enumerate(combinations[:5], 1):
            print(f"  {i}. {combo[3]}")
        if len(combinations) > 5:
            print(f"  ... and {len(combinations) - 5} more")
    
    except Exception as e:
        print(f"\nError: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
