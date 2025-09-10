# Excel Merger

This project merges multiple Excel files into one, handling different column orders and automatically managing headers.

## Usage

1. Put your input Excel files into the `files/` directory.
2. Run the script:

```bash
python merge_excel.py [--source-info] [--keep-duplicates] [--threads THREADS]
```

### Options

* `--source-info`
  If this flag is passed, the merged Excel file will include two extra columns:

  * `source_file`: the original Excel file each row came from
  * `source_sheet`: the sheet name each row came from

* `--keep-duplicates`
  If this flag is passed, duplicate rows will be kept in the output. By default, duplicate rows are removed.

* `--threads THREADS`
  Number of threads to use for processing (default: 8). Increase for better performance on multi-core systems.

If flags are not passed, these options will use their default behavior.

3. The merged Excel file will appear in the `result/` directory.

### Example

Suppose you have two Excel files in `files/` with different column orders:

`file1.xlsx`:

| Name  | Age | Phone    |
| ----- | --- | -------- |
| Alice | 30  | 123-4567 |
| Bob   | 25  | 987-6543 |

`file2.xlsx`:

| Phone    | Name  | Age | Email                                     |
| -------- | ----- | --- | ----------------------------------------- |
| 555-1234 | Carol | 28  | [carol@email.com](mailto:carol@email.com) |
| 444-5678 | Dave  | 35  | [dave@email.com](mailto:dave@email.com)   |

#### Running without any options:

`python merge_excel.py`

Result (`result/merged.xlsx`):

| Name  | Age | Phone    | Email                                     |
| ----- | --- | -------- | ----------------------------------------- |
| Alice | 30  | 123-4567 |                                           |
| Bob   | 25  | 987-6543 |                                           |
| Carol | 28  | 555-1234 | [carol@email.com](mailto:carol@email.com) |
| Dave  | 35  | 444-5678 | [dave@email.com](mailto:dave@email.com)   |

#### Running with `--source-info`:

`python merge_excel.py --source-info`

Result (`result/merged.xlsx`):

| Name  | Age | Phone    | Email                                     | source\_file | source\_sheet |
| ----- | --- | -------- | ----------------------------------------- | ------------ | ------------- |
| Alice | 30  | 123-4567 |                                           | file1.xlsx   | Sheet1        |
| Bob   | 25  | 987-6543 |                                           | file1.xlsx   | Sheet1        |
| Carol | 28  | 555-1234 | [carol@email.com](mailto:carol@email.com) | file2.xlsx   | Sheet1        |
| Dave  | 35  | 444-5678 | [dave@email.com](mailto:dave@email.com)   | file2.xlsx   | Sheet1        |

#### Running with `--keep-duplicates`:

`python merge_excel.py --keep-duplicates`

This will keep all rows even if they are exact duplicates.

## Features

* **Automatic header detection**: New headers are automatically added to the output
* **Column mapping**: Data is placed in the correct columns regardless of input file column order
* **Duplicate handling**: Removes duplicate rows by default (can be disabled with `--keep-duplicates`)
* **Multi-threading**: Process multiple files simultaneously for better performance
* **Source tracking**: Optionally include source file and sheet information

## Setup

Create and activate a Python virtual environment:

```bash
python -m venv venv
source venv/bin/activate   # Linux/macOS
venv\Scripts\activate     # Windows
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Dependencies are minimal:

```text
openpyxl
tqdm
```

## Performance Tips

* Use the `--threads` option to match your CPU core count for better performance
* For large datasets, consider increasing the thread count (e.g., `--threads 16`)
* The script uses streaming mode to handle large files with minimal memory usage

## License

This project is licensed under the [GNU General Public License v3.0](LICENSE).
You are free to use, modify, and distribute it under the terms of the GPLv3.

