# 📊 Excel Word Matching Tools

This repository contains a set of tools to compare and match words between two Excel sheets (`Sheet1` and `Sheet2`) using different matching techniques:

1. **Strict Levenshtein Distance Matching**
2. **Streamlit App for Interactive Levenshtein Matching**
3. **Streamlit App for Semantic Matching using Sentence Transformers**

---

## 🚀 Features

- Compare words from two Excel sheets column-wise
- Use Levenshtein distance for strict word similarity
- Use Sentence Transformers for deep semantic similarity
- Automatic Excel file generation with color-coded results
- Interactive user interface using **Streamlit**
- Excel output download with highlighted matches

---

## 📁 Files & Structure

- `strict_match.py` - Script for command-line based Levenshtein matching
- `strict_match_streamlit.py` - Streamlit app for Levenshtein-based word comparison
- `semantic_match_streamlit.py` - Streamlit app using semantic similarity with Hugging Face model

---

## 🧠 Matching Logic

### Levenshtein Distance Matching
- Words are matched if:
  - **Edit Distance ≤ 2**
  - **Similarity ≥ 50%**
- Matches are highlighted:
  - ✅ Pink: High similarity (≥85%)
  - ✅ Blue: Moderate similarity (60–84%)
  - ✅ Yellow: Low similarity (<60%)

### Semantic Matching (NLP-based)
- Uses Hugging Face's **MiniLM-L6-v2** model
- Cosine similarity between sentence embeddings
- Matches are shown only if similarity is **≥ 50%**
- Colored Excel output similar to the above

---

## 🖥️ How to Run

### 1. 📌 Install Dependencies

```bash
pip install pandas openpyxl streamlit rapidfuzz sentence-transformers
