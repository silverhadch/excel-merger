import argparse
import glob
from tqdm import tqdm
from openpyxl import Workbook, load_workbook
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue
import threading
from collections import OrderedDict
import time
import gc

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
    default=4,  # Reduced default to save memory
    help="Number of reading threads"
)
parser.add_argument(
    "--batch-size",
    type=int,
    default=500,  # Reduced batch size
    help="Number of rows to process in batch"
)
parser.add_argument(
    "--max-memory-rows",
    type=int,
    default=100000,  # Limit total rows in memory
    help="Maximum rows to keep in memory before flushing"
)
args = parser.parse_args()

print("Source info enabled:", args.source_info)
print("Keep duplicates:", args.keep_duplicates)
print("Threads:", args.threads)
print("Batch size:", args.batch_size)
print("Max memory rows:", args.max_memory_rows)

# ----------------------
# Collect files
# ----------------------
files = glob.glob("files/*.xlsx")
print(f"Found {len(files)} files to process")

# ----------------------
# Output workbook (streaming)
# ----------------------
wb_out = Workbook(write_only=True)
ws_out = wb_out.create_sheet("Merged")

# ----------------------
# Global header management
# ----------------------
header_lock = threading.Lock()
all_headers = OrderedDict()
header_ready = threading.Event()
seen_rows = set() if not args.keep_duplicates else None

# ----------------------
# Memory management
# ----------------------
memory_counter = 0
memory_lock = threading.Lock()

def check_memory_usage():
    """Check if we're approaching memory limits and trigger GC if needed"""
    global memory_counter
    with memory_lock:
        if memory_counter >= args.max_memory_rows:
            gc.collect()  # Force garbage collection
            memory_counter = 0
            return True
    return False

# ----------------------
# Thread-safe queue for rows with memory limits
# ----------------------
row_queue = queue.Queue(maxsize=1000)  # Smaller queue to limit memory

def process_sheet(ws, sheet_name, f):
    """Process sheet with streaming to avoid loading all rows at once"""
    rows_processed = 0

    # Regular worksheet - process row by row
    if hasattr(ws, 'iter_rows'):
        try:
            rows_iter = ws.iter_rows(values_only=True)
            first_row = next(rows_iter, None)
            if not first_row:
                return 0

            file_headers = [str(h).strip() if h is not None else "" for h in first_row]
            source_headers = ["source_file", "source_sheet"] if args.source_info else []

            # Cache header mapping
            header_mapping = []
            for i, header in enumerate(file_headers):
                with header_lock:
                    if header not in all_headers:
                        all_headers[header] = len(all_headers)
                    header_mapping.append((i, all_headers[header]))

            # Add source headers to mapping
            source_mapping = []
            for header in source_headers:
                with header_lock:
                    if header not in all_headers:
                        all_headers[header] = len(all_headers)
                    source_mapping.append(all_headers[header])

            # Send header if first file
            if not header_ready.is_set():
                with header_lock:
                    full_header_row = [""] * len(all_headers)
                    for header, idx in all_headers.items():
                        full_header_row[idx] = header
                    row_queue.put(("header", full_header_row))
                    header_ready.set()

            # Process rows in streaming fashion
            batch = []
            for row in rows_iter:
                if row is None:
                    continue

                # Create ordered row
                ordered_row = [""] * len(all_headers)

                # Map file data
                for src_idx, dst_idx in header_mapping:
                    if src_idx < len(row) and row[src_idx] is not None:
                        ordered_row[dst_idx] = str(row[src_idx]).strip()

                # Add source info
                if args.source_info:
                    for i, dst_idx in enumerate(source_mapping):
                        if i == 0:
                            ordered_row[dst_idx] = f
                        elif i == 1:
                            ordered_row[dst_idx] = sheet_name

                # Check duplicates
                row_tuple = tuple(ordered_row)
                if args.keep_duplicates or row_tuple not in seen_rows:
                    if not args.keep_duplicates:
                        seen_rows.add(row_tuple)
                    batch.append(ordered_row)
                    rows_processed += 1

                    # Check memory and send batch
                    if len(batch) >= args.batch_size or check_memory_usage():
                        row_queue.put(("batch", batch))
                        batch = []
                        global memory_counter
                        with memory_lock:
                            memory_counter += len(batch)

            # Send remaining rows
            if batch:
                row_queue.put(("batch", batch))
                with memory_lock:
                    memory_counter += len(batch)

        except Exception as e:
            print(f"Warning: Could not read worksheet '{sheet_name}' in file '{f}': {e}")

    return rows_processed

def file_reader(f):
    try:
        total_rows = 0
        wb_in = load_workbook(f, read_only=True, data_only=True)

        for sheet_name in wb_in.sheetnames:
            ws_in = wb_in[sheet_name]

            # Skip chartsheets and hidden sheets to save memory
            if (hasattr(ws_in, 'sheet_state') and ws_in.sheet_state == 'hidden') or \
               (hasattr(ws_in, 'charts') and not hasattr(ws_in, 'iter_rows')):
                continue

            rows_processed = process_sheet(ws_in, sheet_name, f)
            total_rows += rows_processed

            # Force GC after each sheet to free memory
            gc.collect()

    except Exception as e:
        print(f"Error processing file {f}: {e}")
    finally:
        if 'wb_in' in locals():
            wb_in.close()
        gc.collect()  # Clean up after each file

    return f, total_rows

# ----------------------
# Writer thread with memory awareness
# ----------------------
def writer():
    processed = 0
    start_time = time.time()

    while True:
        item = row_queue.get()
        if item == "DONE":
            break

        typ, data = item
        if typ == "batch":
            for row in data:
                ws_out.append(row)
                processed += 1

                # Progress and memory reporting
                if processed % 5000 == 0:
                    elapsed = time.time() - start_time
                    print(f"Processed {processed} rows in {elapsed:.2f}s ({processed/elapsed:.0f} rows/s) - Memory: {memory_counter}/{args.max_memory_rows}")
        else:
            ws_out.append(data)

        row_queue.task_done()

        # Force GC periodically
        if processed % 10000 == 0:
            gc.collect()

    print(f"Total rows written: {processed}")

# ----------------------
# Start writer thread
# ----------------------
writer_thread = threading.Thread(target=writer)
writer_thread.start()

# ----------------------
# Start reading threads with memory limits
# ----------------------
start_time = time.time()
total_files_processed = 0

with ThreadPoolExecutor(max_workers=args.threads) as executor:
    futures = {executor.submit(file_reader, f): f for f in files}

    for future in tqdm(as_completed(futures), total=len(futures), desc="Processing files"):
        try:
            f, rows_processed = future.result()
            total_files_processed += 1
            if total_files_processed % 10 == 0:
                gc.collect()  # GC after every 10 files
        except Exception as e:
            print(f"Error: {e}")

# ----------------------
# Finish with cleanup
# ----------------------
row_queue.put("DONE")
writer_thread.join()

# Clean up large data structures
if seen_rows:
    seen_rows.clear()
all_headers.clear()

gc.collect()

total_time = time.time() - start_time
print(f"Processing completed in {total_time:.2f} seconds")

wb_out.save("result/merged.xlsx")
print("Done → result/merged.xlsx")
