import argparse
import warnings

import pandas as pd
from colorama import Fore, Style, init
from tqdm import tqdm

warnings.filterwarnings(
    "ignore", message="Workbook contains no default style, apply openpyxl's default"
)

init(autoreset=True)  # Initialize colorama


def merge_xlsx(input_files, output_file):
    dataframes = []
    print(Fore.CYAN + "Merging Excel files...\n")
    for idx, file in enumerate(tqdm(input_files, desc="Processing files", unit="file")):
        print(
            f"{Fore.YELLOW}[{idx+1}/{len(input_files)}]{Style.RESET_ALL} "
            f"Reading {Fore.GREEN}{file}{Style.RESET_ALL}"
        )
        df = pd.read_excel(file)
        dataframes.append(df)
    merged_df = pd.concat(dataframes, ignore_index=True)
    print(
        Fore.CYAN
        + f"\nWriting merged file to {Fore.GREEN}{output_file}{Style.RESET_ALL}"
    )
    merged_df.to_excel(output_file, index=False)
    print(Fore.GREEN + "Merge complete!" + Style.RESET_ALL)


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
