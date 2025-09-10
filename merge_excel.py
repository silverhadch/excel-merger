import argparse
import glob
from tqdm import tqdm
from openpyxl import Workbook, load_workbook
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue
import threading
from collections import OrderedDict

# ----------------------
# Arguments
# ----------------------
parser = argparse.ArgumentParser()
parser.add_argument(
    "--source-info",
    action="store_true",
    help="Include source information"
)
parser.add_argument(
    "--keep-duplicates",
    action="store_true",
    help="Keep duplicate rows instead of removing them"
)
parser.add_argument(
    "--threads",
    type=int,
    default=8,
    help="Number of reading threads"
)
args = parser.parse_args()

print("Source info enabled:", args.source_info)
print("Keep duplicates:", args.keep_duplicates)
print("Threads:", args.threads)

# ----------------------
# Collect files
# ----------------------
files = glob.glob("files/*.xlsx")

# ----------------------
# Output workbook (streaming)
# ----------------------
wb_out = Workbook(write_only=True)
ws_out = wb_out.create_sheet("Merged")

# ----------------------
# Global header management
# ----------------------
header_lock = threading.Lock()
all_headers = OrderedDict()  # Maps header name to column index
header_ready = threading.Event()
seen_rows = set() if not args.keep_duplicates else None

# ----------------------
# Thread-safe queue for rows
# ----------------------
row_queue = queue.Queue(maxsize=10000)  # Limit memory usage

def file_reader(f):
    wb_in = load_workbook(f, read_only=True, data_only=True)
    for sheet_name in wb_in.sheetnames:
        ws_in = wb_in[sheet_name]
        rows_iter = ws_in.iter_rows(values_only=True)
        try:
            first_row = next(rows_iter)
        except StopIteration:
            continue

        # Process headers
        file_headers = [str(h).strip() if h is not None else "" for h in first_row]

        # Add source info headers if needed
        source_headers = []
        if args.source_info:
            source_headers = ["source_file", "source_sheet"]

        # Update global headers with lock
        with header_lock:
            # Add any new headers to our global header dict
            for header in file_headers + source_headers:
                if header not in all_headers:
                    all_headers[header] = len(all_headers)

            # If this is the first file, send the header row
            if not header_ready.is_set():
                full_header_row = [""] * len(all_headers)
                for header, idx in all_headers.items():
                    full_header_row[idx] = header
                row_queue.put(("header", full_header_row))
                header_ready.set()

        # Process data rows
        for row in rows_iter:
            # Convert row to list and handle None values
            row_data = [str(cell).strip() if cell is not None else "" for cell in row]

            # Add source info if needed
            if args.source_info:
                row_data += [f, sheet_name]

            # Create a properly ordered row based on global headers
            ordered_row = [""] * len(all_headers)

            # Map each value to its correct column based on header position
            for i, value in enumerate(row_data):
                if i < len(file_headers):
                    header_name = file_headers[i]
                else:
                    # This handles source info columns
                    header_name = source_headers[i - len(file_headers)]

                if header_name in all_headers:
                    col_idx = all_headers[header_name]
                    ordered_row[col_idx] = value

            # Check for duplicates if needed
            row_tuple = tuple(ordered_row)
            if args.keep_duplicates or row_tuple not in seen_rows:
                if not args.keep_duplicates:
                    seen_rows.add(row_tuple)
                row_queue.put(("row", ordered_row))

    return f

# ----------------------
# Writer thread
# ----------------------
def writer():
    while True:
        item = row_queue.get()
        if item == "DONE":
            break
        typ, data = item
        ws_out.append(data)
        row_queue.task_done()

# ----------------------
# Start writer thread
# ----------------------
writer_thread = threading.Thread(target=writer)
writer_thread.start()

# ----------------------
# Start reading threads
# ----------------------
with ThreadPoolExecutor(max_workers=args.threads) as executor:
    futures = [executor.submit(file_reader, f) for f in files]
    for _ in tqdm(as_completed(futures), total=len(futures), desc="Processing files"):
        pass

# ----------------------
# Finish
# ----------------------
row_queue.put("DONE")
writer_thread.join()
wb_out.save("result/merged.xlsx")
print("Done → result/merged.xlsx")
