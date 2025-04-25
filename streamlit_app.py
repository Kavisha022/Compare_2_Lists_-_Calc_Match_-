import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from io import BytesIO

st.title("Fuzzy Matching Tool")
uploaded_file = st.file_uploader("Upload Excel File with 'List 1' and 'List 2'", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    if 'List 1' not in df.columns or 'List 2' not in df.columns:
        st.error("Excel must contain columns 'List 1' and 'List 2'")
    else:
        list1 = df['List 1'].dropna().tolist()
        list2 = df['List 2'].dropna().tolist()

        matches = []

        # Match List 2 -> List 1
        for item2 in list2:
            best_match = None
            best_score = 0
            for item1 in list1:
                score = fuzz.token_sort_ratio(str(item2), str(item1))
                if score > best_score:
                    best_score = score
                    best_match = item1
            if not any(m[0] == item2 and m[1] == best_match for m in matches):
                matches.append((item2, best_match, best_score, "List 2"))

        result_df = pd.DataFrame(matches, columns=['Word', 'Best Match', 'Match Percentage', 'List Source'])

        st.success("Matching Completed!")
        st.dataframe(result_df)

        output = BytesIO()
        result_df.to_excel(output, index=False)
        output.seek(0)

        #Coloring
        wb = load_workbook(output)
        ws = wb.active

        green_fill = PatternFill(start_color="FF00FF00", end_color="FF00FF00", fill_type="solid")   # Light green
        blue_fill = PatternFill(start_color="FF66CCFF", end_color="FF66CCFF", fill_type="solid")    # Light blue
        yellow_fill = PatternFill(start_color="FFFFCC00", end_color="FFFFCC00", fill_type="solid")  # Yellow

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

        # Save to new output
        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        # Download link
        st.download_button("📥 Download Colored Excel", data=final_output, file_name="matches_output.xlsx")


