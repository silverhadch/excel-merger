# Excel Merger

This project merges multiple Excel files into one.

## Usage

1. Put your input Excel files into the `files/` directory.
2. Run the script:

```bash
python merge_excel.py [--source-info]
```

### Options

* `--source-info`
  If this flag is passed, the merged Excel file will include two extra columns:

  * `source_file`: the original Excel file each row came from
  * `source_sheet`: the sheet name each row came from

If the flag is not passed, these columns will be omitted.

3. The merged Excel file will appear in the `result/` directory.

### Example

Suppose you have two Excel files in `files/`:

`file1.xlsx`:

| Name  | Age |
| ----- | --- |
| Alice | 30  |
| Bob   | 25  |

`file2.xlsx`:

| Name  | Age |
| ----- | --- |
| Carol | 28  |
| Dave  | 35  |

#### Running without `--source-info`:

`python merge_excel.py`

Result (`result/merged.xlsx`):

| Name  | Age |
| ----- | --- |
| Alice | 30  |
| Bob   | 25  |
| Carol | 28  |
| Dave  | 35  |

#### Running with `--source-info`:

`python merge_excel.py --source-info`

Result (`result/merged.xlsx`):

| Name  | Age | source\_file | source\_sheet |
| ----- | --- | ------------ | ------------- |
| Alice | 30  | file1.xlsx   | Sheet1        |
| Bob   | 25  | file1.xlsx   | Sheet1        |
| Carol | 28  | file2.xlsx   | Sheet1        |
| Dave  | 35  | file2.xlsx   | Sheet1        |

---

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
pandas
openpyxl
```

---

## License

This project is licensed under the [GNU General Public License v3.0](LICENSE).
You are free to use, modify, and distribute it under the terms of the GPLv3.
