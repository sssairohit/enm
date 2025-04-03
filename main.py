import pandas as pd
from rapidfuzz import process, fuzz

input_file = "input.xlsx"
output_file = "output.xlsx"

df1 = pd.read_excel(input_file, sheet_name="Sheet1", header=None)
df2 = pd.read_excel(input_file, sheet_name="Sheet2", header=None)

names_to_fix = df1[0].tolist()
correct_names = df2[0].tolist()

def find_best_match(name, choices):
    match, score, _ = process.extractOne(name, choices, scorer=fuzz.partial_ratio)
    return match if score > 80 else name

df1[0] = df1[0].apply(lambda x: find_best_match(x, correct_names))

df1.to_excel(output_file, index=False, header=False)

print(f"Updated names saved to {output_file}")
