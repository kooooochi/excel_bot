import os
import sys
from pathlib import Path
import pandas as pd
from tqdm import tqdm


def process_excel_file(input_path, output_path):
    """
    Process a single Excel file and save to output directory.

    Args:
        input_path: Path to input Excel file
        output_path: Path to output Excel file
    """
    print(f"Processing: {input_path}")

    # Read Excel file
    df = pd.read_excel(input_path)

    # Add your processing logic here
    # Example: Add a new column with row numbers
    df['processed'] = True
    df['row_number'] = range(1, len(df) + 1)

    # Save to output directory
    df.to_excel(output_path, index=False)
    print(f"Saved to: {output_path}")


def main():
    input_dir = Path("input")
    output_dir = Path("output")

    # Ensure output directory exists
    output_dir.mkdir(exist_ok=True)

    # Find all Excel files in input directory
    excel_files = list(input_dir.glob("*.xlsx")) + list(input_dir.glob("*.xls"))

    # Filter out temporary files (starting with ~$)
    excel_files = [f for f in excel_files if not f.name.startswith("~$")]

    if not excel_files:
        print("No Excel files found in input directory.")
        sys.exit(0)

    print(f"Found {len(excel_files)} Excel file(s) to process.")

    # Process each file with progress bar
    for input_file in tqdm(excel_files, desc="Processing files"):
        try:
            output_file = output_dir / f"processed_{input_file.name}"
            process_excel_file(input_file, output_file)
        except Exception as e:
            print(f"Error processing {input_file}: {e}")
            sys.exit(1)

    print("\nAll files processed successfully!")


if __name__ == "__main__":
    main()
