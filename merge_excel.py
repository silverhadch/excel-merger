import argparse
import glob
from tqdm import tqdm
from openpyxl import Workbook, load_workbook

# ----------------------
# Arguments
# ----------------------
parser = argparse.ArgumentParser()
parser.add_argument(
    "--source-info",
    action="store_true",
    help="Include source information"
)
args = parser.parse_args()
print("Source info enabled:", args.source_info)

# ----------------------
# Collect files
# ----------------------
files = glob.glob("files/*.xlsx")

# ----------------------
# Output workbook (streaming)
# ----------------------
wb_out = Workbook(write_only=True)
ws_out = wb_out.create_sheet("Merged")

header_written = False

# ----------------------
# Process files
# ----------------------
for f in tqdm(files, desc="Processing files"):
    wb_in = load_workbook(f, read_only=True, data_only=True)
    for sheet_name in tqdm(wb_in.sheetnames, desc=f"  {f}", leave=False):
        ws_in = wb_in[sheet_name]

        rows_iter = ws_in.iter_rows(values_only=True)

        try:
            # Read first row
            first_row = next(rows_iter)
        except StopIteration:
            # Empty sheet, skip
            continue

        # Write header once
        if not header_written:
            header = list(first_row)
            if args.source_info:
                header += ["source_file", "source_sheet"]
            ws_out.append(header)
            header_written = True
        else:
            # Skip header of subsequent sheets
            pass

        # Write remaining rows
        for row in rows_iter:
            row = list(row)
            if args.source_info:
                row += [f, sheet_name]
            ws_out.append(row)

# ----------------------
# Save result
# ----------------------
wb_out.save("result/merged.xlsx")
print("Done → result/merged.xlsx")

