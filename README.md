## Excel Name Matching

### Project Overview
This project standardizes names in an Excel file using fuzzy matching techniques.

### How It Works
- Reads `input.xlsx`, containing two sheets:
  - **Sheet1**: Names that need correction.
  - **Sheet2**: The correct reference names.
- Matches names in **Sheet1** to the closest correct name in **Sheet2**.
- Replaces names in **Sheet1** with the matched names from **Sheet2**.
- Saves the updated names to `output.xlsx`.

### Installation
1. Clone the repository or download the script.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Ensure `input.xlsx` is placed in the working directory.

### Usage
Run the script with:
```bash
python main.py
```
After execution, check `output.xlsx` for updated names.

### File Structure
```
/excel_name_matching/
│── input.xlsx       # Raw input file (Sheet1 & Sheet2)
│── output.xlsx      # Processed file (after name correction)
│── main.py         # Python script for name matching
│── requirements.txt # Dependencies list
│── README.md        # Project documentation
```

### Configuration
- The script uses `rapidfuzz` for fuzzy name matching.
- Names are replaced **only if similarity is above 80%**.
- Modify this threshold in `main.py` if necessary.
- The script retains the original name if no close match is found.

### Example
#### Input
**Sheet1 (Names to Fix):**
```
2pi System Private Limited
Apple
Microsoft
```
**Sheet2 (Correct Names):**
```
2pi Systems
Apple Inc
Microsoft Corp
```

#### Output (`output.xlsx`):
```
2pi Systems
Apple Inc
Microsoft Corp
```

### Dependencies
This project requires:
```
pandas
openpyxl
rapidfuzz
```
Install them with:
```bash
pip install -r requirements.txt
```

### Contributions
Contributions are welcome. Modify and improve as needed.

### License
This project is licensed under the MIT License. See the `LICENSE` file for details.