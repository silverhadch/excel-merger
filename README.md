# Excel Merger - Memory-Optimized High Performance

This project merges multiple Excel files into one with aggressive memory optimization and performance tuning. It handles different column orders, automatically manages headers, and processes files efficiently while keeping memory usage under control.

## Usage

1. Put your input Excel files into the `files/` directory.
2. Run the script with memory-appropriate settings:

```bash
python merge_excel.py [--source-info] [--keep-duplicates] [--threads THREADS] [--batch-size BATCH_SIZE] [--max-memory-rows MAX_MEMORY_ROWS]
```

### Memory-Safe Options

* `--threads THREADS` (default: 4)
  Number of parallel threads for processing. Lower values reduce memory usage.

* `--batch-size BATCH_SIZE` (default: 500)
  Number of rows to process in batches. Smaller batches use less memory.

* `--max-memory-rows MAX_MEMORY_ROWS` (default: 100000)
  Maximum number of rows to keep in memory before flushing. Critical for large files.

### Data Processing Options

* `--source-info`
  Include source information columns:

  * `source_file`: original Excel file name
  * `source_sheet`: sheet name

* `--keep-duplicates`
  Keep duplicate rows (default: duplicates are removed)

3. The merged Excel file will appear in the `result/` directory.

## Memory-Optimized Performance Recommendations

### For Small Files (<10MB each)

```bash
python merge_excel.py --threads 4 --batch-size 500 --max-memory-rows 100000
```

### For Medium Files (10-100MB each)

```bash
python merge_excel.py --threads 2 --batch-size 300 --max-memory-rows 50000
```

### For Large Files (100MB-1GB each)

```bash
python merge_excel.py --threads 1 --batch-size 200 --max-memory-rows 30000
```

### For Very Large Files (>1GB each)

```bash
python merge_excel.py --threads 1 --batch-size 100 --max-memory-rows 20000
```

### Memory-Safe with Source Information

```bash
python merge_excel.py --source-info --threads 2 --batch-size 200 --max-memory-rows 40000
```

## Memory Management Features

* **Controlled Memory Usage**: Hard limits on maximum rows in memory
* **Streaming Processing**: Row-by-row processing avoids loading entire files
* **Automatic Garbage Collection**: Periodic cleanup to free memory
* **Batch Size Control**: Configurable batch processing to balance speed vs memory
* **Queue Size Limits**: Bounded queues prevent memory overflow
* **Selective Processing**: Skips memory-intensive chartsheets and hidden sheets

## Example

### Input Files with Different Structures:

**file1.xlsx:**

| Name  | Age | Phone    |
| ----- | --- | -------- |
| Alice | 30  | 123-4567 |
| Bob   | 25  | 987-6543 |

**file2.xlsx:**

| Phone    | Name  | Age | Email                                     |
| -------- | ----- | --- | ----------------------------------------- |
| 555-1234 | Carol | 28  | [carol@email.com](mailto:carol@email.com) |
| 444-5678 | Dave  | 35  | [dave@email.com](mailto:dave@email.com)   |

### Output (merged.xlsx):

| Name  | Age | Phone    | Email                                     |
| ----- | --- | -------- | ----------------------------------------- |
| Alice | 30  | 123-4567 |                                           |
| Bob   | 25  | 987-6543 |                                           |
| Carol | 28  | 555-1234 | [carol@email.com](mailto:carol@email.com) |
| Dave  | 35  | 444-5678 | [dave@email.com](mailto:dave@email.com)   |

## Technical Architecture

### Memory Optimization

* Streaming read/write mode with row-by-row processing
* Configurable memory limits with automatic flushing
* Periodic garbage collection
* Bounded queue sizes
* Efficient data structures with pre-allocation

### Performance Features

* Multi-threaded processing within memory constraints
* Batched operations for reduced overhead
* Real-time memory monitoring and reporting
* Automatic header detection and column mapping
* Duplicate removal with efficient hashing

### Error Handling

* Graceful memory error recovery
* Continue processing after individual file errors
* Detailed memory usage reporting
* Resource cleanup guarantees

## Setup

```bash
# Create virtual environment
python -m venv venv

# Activate (Linux/macOS)
source venv/bin/activate

# Activate (Windows)
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### Dependencies

```text
openpyxl>=3.1.0
tqdm>=4.66.0
```

## Memory Usage Benchmarks

Example memory usage on 16GB RAM system:

| File Size   | Configuration                        | Max Memory | Time |
| ----------- | ------------------------------------ | ---------- | ---- |
| 100MB total | Default (4 threads)                  | \~500MB    | 45s  |
| 1GB total   | Medium (2 threads)                   | \~1.2GB    | 3m   |
| 10GB total  | Conservative (1 thread)              | \~2GB      | 15m  |
| 50GB+ total | Ultra-safe (1 thread, small batches) | \~2.5GB    | 60m+ |

## Troubleshooting

### Memory Issues

* **Out of Memory**: Reduce `--max-memory-rows` and `--batch-size`
* **Slow Processing**: Increase `--batch-size` slightly if memory allows
* **Swap Usage**: Reduce `--threads` and `--max-memory-rows`

### Performance Issues

* **CPU Underutilized**: Carefully increase `--threads` if memory available
* **Disk I/O Bound**: Ensure files are on fast storage (SSD recommended)

### File Processing

* **Chartsheets**: Automatically skipped to save memory
* **Hidden Sheets**: Automatically skipped
* **Corrupted Files**: Error handling continues with other files

## Monitoring

The script provides real-time memory monitoring:

```
Processed 5000 rows in 12.45s (402 rows/s) - Memory: 45000/50000
```

## License

This project is licensed under the [GNU General Public License v3.0](LICENSE).

## Warning

For extremely large datasets (>50GB total), consider:

1. Using a database instead of Excel
2. Splitting the merge into multiple batches
3. Ensuring sufficient disk space for output file
4. Monitoring system resources during operation

The script includes safety limits, but extremely large merges may still require significant resources.
