# Excel Merger

This project merges multiple Excel files into one.

## Usage

1. Put your input Excel files into the `files/` directory.  
2. Run the script.  
3. The merged Excel file will appear in the `result/` directory.  

## Setup

Create and activate a Python virtual environment:

```bash
python -m venv venv
source venv/bin/activate   # Linux/macOS
venv\Scripts\activate    # Windows
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
