import argparse
import warnings

import pandas as pd

warnings.filterwarnings(
    "ignore", message="Workbook contains no default style, apply openpyxl's default"
)


def merge_xlsx(input_files, output_file):

    # Read and concatenate all input files
    dataframes = [pd.read_excel(file) for file in input_files]
    merged_df = pd.concat(dataframes, ignore_index=True)

    # Write the merged DataFrame to an Excel file
    merged_df.to_excel(output_file, index=False)


def main():
    parser = argparse.ArgumentParser(description="Merge multiple Excel files into one.")
    parser.add_argument(
        "input_files",
        nargs="+",
        help="List of input Excel files to merge.",
    )
    parser.add_argument(
        "-o",
        "--output",
        required=True,
        help="Output file name for the merged Excel file.",
    )

    args = parser.parse_args()

    merge_xlsx(args.input_files, args.output)
    print(f"Merged {len(args.input_files)} files into {args.output}")
