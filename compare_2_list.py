import pandas as pd
from fuzzywuzzy import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Paths
file_path = r'D:\OneDrive - Lowcode Minds Technology Pvt Ltd\Desktop\Compare two list\Compare_2_list.xlsx'
output_path = r'D:\OneDrive - Lowcode Minds Technology Pvt Ltd\Desktop\Compare two list\matches_output.xlsx'
sheet_name = 'Sheet1'

# Read data
df = pd.read_excel(file_path, sheet_name=sheet_name)
list1 = df['List 1'].dropna().tolist()
list2 = df['List 2'].dropna().tolist()

matches = []

# # Match List 1 -> List 2
# for item1 in list1:
#     best_match = None
#     best_score = 0
#     for item2 in list2:
#         score = fuzz.token_sort_ratio(str(item1), str(item2))
#         if score > best_score:
#             best_score = score
#             best_match = item2
#     matches.append((item1, best_match, best_score, "List 1"))

# Match List 2 -> List 1
for item2 in list2:
    best_match = None
    best_score = 0
    for item1 in list1:
        score = fuzz.token_sort_ratio(str(item2), str(item1))
        if score > best_score:
            best_score = score
            best_match = item1
    # Prevent duplicates
    if not any(m[0] == item2 and m[1] == best_match for m in matches):
        matches.append((item2, best_match, best_score, "List 2"))

# Save matches
result_df = pd.DataFrame(matches, columns=['Word', 'Best Match', 'Match Percentage', 'List Source'])
result_df.to_excel(output_path, index=False)

# Load workbook for coloring
wb = load_workbook(output_path)
ws = wb.active

# Define fills
green_fill = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")   # Light green
blue_fill = PatternFill(start_color="FFD9E1F2", end_color="FFD9E1F2", fill_type="solid")    # Light blue
yellow_fill = PatternFill(start_color="FFFFFF99", end_color="FFFFFF99", fill_type="solid")  # Yellow

# Apply color formatting
for row in ws.iter_rows(min_row=2, min_col=1, max_col=4):
    score = row[2].value
    if score is None:
        continue

    if score >= 85:
        fill = green_fill
    elif score >= 60:
        fill = blue_fill
    else:
        fill = yellow_fill

    for cell in row:
        cell.fill = fill

wb.save(output_path)
print(f"✅ Two-way matching complete! File saved to:\n{output_path}")
