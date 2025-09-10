import argparse
import pandas as pd
import glob as glb

# Setup Cmdline Parser
parser = argparse.ArgumentParser()

parser.add_argument(
    "--source-info",
    dest="source_info",
    action="store_true",
    help="Include source information"
)

args = parser.parse_args()

print("Source info enabled:", args.source_info)

# collect all Excel files in files folder
files = glb.glob("files/*.xlsx")

merged = pd.DataFrame()

for f in files:
    # read all sheets from the file
    xls = pd.ExcelFile(f)
    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)
        if source_info:
            df["source_file"] = f
            df["source_sheet"] = sheet_name
        merged = pd.concat([merged, df], ignore_index=True)

# write to a single Excel file
merged.to_excel("result/merged.xlsx", index=False)
