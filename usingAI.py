import streamlit as st
import pandas as pd
from sentence_transformers import SentenceTransformer, util
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from io import BytesIO

# Load the AI model
@st.cache_resource
def load_model():
    return SentenceTransformer('all-MiniLM-L6-v2')

model = load_model()

st.title("🤖 AI-Powered List Matcher")
st.markdown("Upload an Excel file with **'List 1'** and **'List 2'** columns. The AI will match them based on meaning.")

uploaded_file = st.file_uploader("📤 Upload Excel File", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    if 'List 1' not in df.columns or 'List 2' not in df.columns:
        st.error("Excel must contain columns 'List 1' and 'List 2'")
    else:
        list1 = df['List 1'].dropna().astype(str).tolist()
        list2 = df['List 2'].dropna().astype(str).tolist()

        st.info("🔍 Matching with AI model...")

        emb1 = model.encode(list1, convert_to_tensor=True)
        emb2 = model.encode(list2, convert_to_tensor=True)

        matches = []

        for i, sentence1 in enumerate(list1):
            cosine_scores = util.pytorch_cos_sim(emb1[i], emb2)[0]
            best_score_idx = cosine_scores.argmax().item()
            best_score = cosine_scores[best_score_idx].item()
            best_match = list2[best_score_idx]
            matches.append((sentence1, best_match, round(best_score * 100, 2)))

        result_df = pd.DataFrame(matches, columns=['List 1', 'Best Match from List 2', 'Match Percentage'])
        st.success("✅ Matching complete!")
        st.dataframe(result_df)

        # --- Create Excel with colors ---
        output = BytesIO()
        result_df.to_excel(output, index=False)
        output.seek(0)

        wb = load_workbook(output)
        ws = wb.active

        # Colors: vibrant style
        green_fill = PatternFill(start_color="FF00FF00", end_color="FF00FF00", fill_type="solid")   # Bright green
        blue_fill = PatternFill(start_color="FF66CCFF", end_color="FF66CCFF", fill_type="solid")    # Sky blue
        yellow_fill = PatternFill(start_color="FFFFCC00", end_color="FFFFCC00", fill_type="solid")  # Golden yellow

        for row in ws.iter_rows(min_row=2, min_col=1, max_col=3):
            score = row[2].value
            if score >= 85:
                fill = green_fill
            elif score >= 60:
                fill = blue_fill
            else:
                fill = yellow_fill
            for cell in row:
                cell.fill = fill

        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        st.download_button("📥 Download Colored Excel", data=final_output, file_name="ai_matches_output.xlsx")
