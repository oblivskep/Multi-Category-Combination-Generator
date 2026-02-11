import csv
import os
import sys
import tempfile
from pathlib import Path


def _detect_category_mapping(headers):
    """
    Detect category mapping from row-1 headers.
    Returns (col_to_category, legacy_mode).
    """
    normalized_headers = [(h or "").strip() for h in headers]
    legacy = any(
        h and h.split()[0].upper() in {"C1", "C2", "C3"}
        for h in normalized_headers
    )

    col_to_category = [None] * len(headers)

    if legacy:
        for idx, header in enumerate(normalized_headers):
            if not header:
                continue
            token = header.split()[0].upper()
            if token in {"C1", "C2", "C3"}:
                col_to_category[idx] = token
        return col_to_category, True

    # Human-friendly headers: group adjacent columns with the same header
    groups = []
    current_group = None
    for idx, header in enumerate(normalized_headers):
        if not header:
            continue
        if current_group is None or header != current_group:
            groups.append(header)
            current_group = header
        group_idx = len(groups) - 1
        if group_idx >= 3:
            raise ValueError(
                f"Expected exactly 3 categories from row 1, found {len(groups)}. "
                f"Headers seen (left-to-right): {groups}"
            )
        col_to_category[idx] = f"C{group_idx + 1}"

    if len(groups) != 3:
        raise ValueError(
            f"Expected exactly 3 categories from row 1, found {len(groups)}. "
            f"Headers seen (left-to-right): {groups}"
        )

    return col_to_category, False


def read_and_parse_csv(filepath):
    """
    Read CSV file and extract clusters, classes, and elements.
    """
    with open(filepath, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        rows = list(reader)

    if len(rows) < 2:
        raise ValueError(
            "CSV must include at least two rows. "
            "Row 2 (class names) is required."
        )

    # Extract cluster information from first row
    clusters = rows[0]

    # Extract class names from second row
    class_names = rows[1]

    # Extract elements (rows 3 onwards)
    elements_data = rows[2:]

    # Parse clusters and their classes
    cluster_structure = {}

    col_to_category, _legacy_mode = _detect_category_mapping(clusters)

    for col_idx, cluster_key in enumerate(col_to_category):
        if not cluster_key:
            continue

        class_name = None
        if col_idx < len(class_names):
            class_name = class_names[col_idx]
        class_name = class_name.strip() if class_name else f"Class_{col_idx}"

        if cluster_key not in cluster_structure:
            cluster_structure[cluster_key] = {}

        # Get elements for this class (non-empty cells in this column)
        elements = []
        for row in elements_data:
            if col_idx < len(row):
                elem = row[col_idx].strip()
                if elem:
                    elements.append(elem)

        cluster_structure[cluster_key][class_name] = elements

    return cluster_structure


def generate_combinations(cluster_structure):
    """
    Generate all combinations: one element from C1 × C2 × C3
    """
    # Extract elements from each cluster - aggregate all clusters with the same key
    cluster_aggregates = {'C1': [], 'C2': [], 'C3': []}

    for cluster_key in cluster_structure.keys():
        classes_in_cluster = cluster_structure[cluster_key]
        elements = []

        for class_name, class_elements in sorted(classes_in_cluster.items()):
            elements.extend(class_elements)

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


def _write_temp_csv(content):
    temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False, encoding='utf-8', newline='')
    try:
        temp_file.write(content)
        return temp_file.name
    finally:
        temp_file.close()


def run_self_tests():
    print("\nRunning self-tests...")

    tests = []

    # Legacy C1/C2/C3 style
    legacy_csv = "\n".join([
        "C1,C2,C2,C3",
        "Class 1,Class 2,Class 3,Class 4",
        "A1,B1,B2,C1",
        "A2,B3,,C2",
    ])
    tests.append(("legacy_c1_c2_c3", legacy_csv, 12))

    # Human-friendly Category 1/2/3 style
    human_csv = "\n".join([
        "Category 1,Category 2,Category 2,Category 3",
        "Class 1,Class 2,Class 3,Class 4",
        "A1,B1,B2,C1",
        "A2,B3,,C2",
    ])
    tests.append(("human_friendly", human_csv, 12))

    # Distinct headers Product Type / Color / Size
    distinct_csv = "\n".join([
        "Product Type,Color,Size",
        "Type,Color,Size",
        "A,Red,Small",
        "B,Blue,Large",
    ])
    tests.append(("distinct_headers", distinct_csv, 8))

    passed = 0
    for name, content, expected_count in tests:
        path = _write_temp_csv(content)
        try:
            cluster_structure = read_and_parse_csv(path)
            combos = generate_combinations(cluster_structure)
            if len(combos) != expected_count:
                raise AssertionError(
                    f"Expected {expected_count} combinations, got {len(combos)}"
                )
            print(f"✓ {name}")
            passed += 1
        finally:
            try:
                os.remove(path)
            except OSError:
                pass

    print(f"Self-tests passed: {passed}/{len(tests)}")


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
    if "--self-test" in sys.argv:
        run_self_tests()
        return

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
