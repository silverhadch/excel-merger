import argparse
import glob
from tqdm import tqdm
from openpyxl import Workbook, load_workbook
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue
import threading

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
    "--threads",
    type=int,
    default=8,
    help="Number of reading threads"
)
args = parser.parse_args()

print("Source info enabled:", args.source_info)
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

header_written = threading.Event()

# ----------------------
# Thread-safe queue for rows
# ----------------------
row_queue = queue.Queue(maxsize=10000)  # Limit memory usage

def file_reader(f):
    rows_to_write = []
    wb_in = load_workbook(f, read_only=True, data_only=True)
    for sheet_name in wb_in.sheetnames:
        ws_in = wb_in[sheet_name]
        rows_iter = ws_in.iter_rows(values_only=True)
        try:
            first_row = next(rows_iter)
        except StopIteration:
            continue

        # Include source info in header/rows
        if not header_written.is_set():
            header = list(first_row)
            if args.source_info:
                header += ["source_file", "source_sheet"]
            row_queue.put(("header", header))
            header_written.set()
        else:
            pass  # skip first row

        for row in rows_iter:
            row_list = list(row)
            if args.source_info:
                row_list += [f, sheet_name]
            row_queue.put(("row", row_list))
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

