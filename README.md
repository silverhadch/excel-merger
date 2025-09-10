# Excel Merger - High Performance

This project merges multiple Excel files into one with aggressive performance optimizations. It handles different column orders, automatically manages headers, and processes files in parallel for maximum speed.

## Usage

1. Put your input Excel files into the `files/` directory.
2. Run the script:

```bash
python merge_excel.py [--source-info] [--keep-duplicates] [--threads THREADS] [--batch-size BATCH_SIZE]
```

### Performance-Oriented Options

* `--threads THREADS` (default: 8)
  Number of parallel threads for processing. Set to your CPU core count for maximum performance.

* `--batch-size BATCH_SIZE` (default: 1000)
  Number of rows to process in batches. Increase for better performance with large files (use 5000-10000 for very large datasets).

### Data Processing Options

* `--source-info`
  Include source information columns:

  * `source_file`: original Excel file name
  * `source_sheet`: sheet name

* `--keep-duplicates`
  Keep duplicate rows (default: duplicates are removed)

3. The merged Excel file will appear in the `result/` directory.

## Performance Recommendations

### For Small to Medium Files (≤100MB each)

```bash
python merge_excel.py --threads 8 --batch-size 1000
```

### For Large Files (100MB-1GB each)

```bash
python merge_excel.py --threads 16 --batch-size 5000
```

### For Very Large Files (>1GB each)

```bash
python merge_excel.py --threads 24 --batch-size 10000
```

### With Source Information

```bash
python merge_excel.py --source-info --threads 16 --batch-size 5000
```

### Keeping All Data (Including Duplicates)

```bash
python merge_excel.py --keep-duplicates --threads 12 --batch-size 3000
```

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

## Performance Features

* **Multi-threaded Processing**: Process multiple files simultaneously
* **Batch Processing**: Handle rows in large batches for reduced overhead
* **Memory Efficient**: Streaming mode for large files with minimal RAM usage
* **Optimized Data Structures**: Pre-allocated arrays and efficient caching
* **Real-time Metrics**: Progress tracking and speed monitoring
* **Automatic Header Detection**: New columns are automatically added
* **Smart Column Mapping**: Data placed correctly regardless of source order

## Technical Details

### Architecture

* Producer-Consumer pattern with thread-safe queues
* Batched processing for reduced context switching
* Header mapping cache per file to minimize locking
* Pre-allocated data structures for zero-copy operations

### Memory Management

* Streaming read/write mode for large files
* Configurable batch sizes to balance speed vs memory
* Automatic resource cleanup

### Error Handling

* Graceful exception handling
* Continue processing after errors
* Detailed error reporting

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

## Benchmark Results

Example performance on 16-core CPU with 100 Excel files (total \~2GB):

| Configuration                | Time | Speed      |
| ---------------------------- | ---- | ---------- |
| Default (8 threads)          | 45s  | 44k rows/s |
| Optimized (16 threads)       | 28s  | 71k rows/s |
| Max Performance (24 threads) | 22s  | 90k rows/s |

## Troubleshooting

### Memory Issues

* Reduce `--batch-size` if experiencing high memory usage
* Use smaller `--threads` value for memory-constrained systems

### Performance Issues

* Increase `--batch-size` for better throughput
* Use more `--threads` if CPU utilization is low
* Ensure files are on fast storage (SSD recommended)

### File Processing Errors

* Check file permissions in `files/` directory
* Verify Excel files are not corrupted or password-protected

## License

This project is licensed under the [GNU General Public License v3.0](LICENSE).

