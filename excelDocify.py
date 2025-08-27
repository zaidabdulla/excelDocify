import streamlit as st
import pandas as pd
import requests
from dotenv import load_dotenv
import os
import json
import tiktoken  # counts tokens safely

from openpyxl import load_workbook

from openpyxl.utils import range_boundaries

def extract_pivot_info_with_data(file_path, sheet_name):
    """
    Extract pivot tables, their data source and underlying data for AI analysis.
    Returns list of dicts with pivot details and sample data.
    """
    wb = load_workbook(file_path, data_only=False)
    if sheet_name not in wb.sheetnames:
        return []

    ws = wb[sheet_name]
    pivot_info = []

    for pivot in ws._pivots:
        pivot_details = {
            "pivot_name": getattr(pivot, "name", "Unnamed Pivot"),
            "cache_id": pivot.cacheId,
            "source_range": None,
            "source_sheet": None,
            "sample_data": None
        }

        if pivot.cache and pivot.cache.cacheSource:
            ws_src = pivot.cache.cacheSource.worksheetSource
            pivot_details["source_range"] = ws_src.ref
            pivot_details["source_sheet"] = ws_src.sheet

            # ✅ Read source range as DataFrame
            if ws_src.sheet in wb.sheetnames and ws_src.ref:
                src_ws = wb[ws_src.sheet]
                min_col, min_row, max_col, max_row = range_boundaries(ws_src.ref)
                data = []
                for row in src_ws.iter_rows(min_row=min_row, max_row=max_row,
                                            min_col=min_col, max_col=max_col,
                                            values_only=True):
                    data.append(row)
                if data:
                    headers, rows = data[0], data[1:6]  # take 5 rows max
                    df = pd.DataFrame(rows, columns=headers)
                    pivot_details["sample_data"] = df.to_dict(orient="records")

        pivot_info.append(pivot_details)

    return pivot_info



# Load API key
load_dotenv()
API_KEY = os.getenv("OPENROUTER_API_KEY")

BASE_URL = "https://openrouter.ai/api/v1/chat/completions"

st.title("📊 Excel to AI Insights (OpenRouter)")

def get_excel_column_letter(col_idx):
    """Convert column index to Excel column letter (e.g., 0='A', 1='B', 26='AA')"""
    result = ""
    while True:
        col_idx, remainder = divmod(col_idx, 26)
        result = chr(65 + remainder) + result
        if not col_idx:
            break
        col_idx -= 1
    return result

def extract_sheet_headers(file_path, sheet_name):
    """Extract headers from a specific sheet by detecting the first fully-filled row."""
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # ✅ Find rows where ALL cells are non-empty
    full_rows = df_raw[df_raw.notna().all(axis=1)]
    
    if not full_rows.empty:
        # Pick the first fully filled row as header
        header_row_idx = full_rows.index[0]
        headers = df_raw.iloc[header_row_idx].tolist()
        
        # Create list with column information
        header_info = []
        for idx, header in enumerate(headers):
            if pd.notna(header):
                col_letter = get_excel_column_letter(idx)
                header_info.append({
                    'column': col_letter,
                    'name': str(header),
                    'index': idx + 1
                })
        return header_info
    return []

def extract_unique_headers(file_path, selected_sheets=None):
    """Extract unique headers from selected sheets or all sheets."""
    xls = pd.ExcelFile(file_path)
    
    # If no sheets specified, use all sheets
    sheets_to_process = selected_sheets if selected_sheets else xls.sheet_names
    
    # Dictionary to store headers by sheet
    headers_by_sheet = {}
    unique_header_names = set()
    
    for sheet_name in sheets_to_process:
        if sheet_name in xls.sheet_names:  # Validate sheet exists
            headers = extract_sheet_headers(file_path, sheet_name)
            headers_by_sheet[sheet_name] = headers
            unique_header_names.update(h['name'] for h in headers)
    
    return list(unique_header_names), headers_by_sheet

def count_tokens(text, model="gpt-3.5-turbo"):
    """Count tokens safely using tiktoken."""
    try:
        enc = tiktoken.encoding_for_model(model)
        return len(enc.encode(text))
    except:
        return len(text.split())

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    # Add sheet selection
    st.subheader("📑 Available Sheets")
    selected_sheets = st.multiselect(
        "Select sheets to analyze (leave empty to analyze all sheets):",
        options=sheet_names,
        default=None
    )
    
        # ---- Pivot Analysis ----
    all_pivots = {}
    for sheet in (selected_sheets if selected_sheets else sheet_names):
        try:
            pivots = extract_pivot_info_with_data(uploaded_file, sheet)
            if pivots:
                all_pivots[sheet] = pivots
        except Exception as e:
            st.error(f"Error reading pivots in {sheet}: {e}")
    
    

    # Get headers for selected sheets (or all if none selected)
    unique_headers, headers_by_sheet = extract_unique_headers(uploaded_file, selected_sheets)

    # Display headers
    st.subheader("🛠 Headers Analysis")
    
    # Show headers by sheet
    for sheet, headers in headers_by_sheet.items():
        with st.expander(f"Headers in {sheet}", expanded=False):
            st.write(f"Found {len(headers)} columns:")
            
            # Create a DataFrame for better display
            if headers:
                headers_df = pd.DataFrame(headers)
                headers_df = headers_df[['column', 'name', 'index']]  # Reorder columns
                headers_df.columns = ['Excel Column', 'Header Name', 'Column Index']
                st.dataframe(headers_df, use_container_width=True)
            else:
                st.warning("No headers found in this sheet")
    
    # Show unique headers across selected sheets
    st.subheader("🔄 Unique Headers Across Selected Sheets")
    # Create a set of unique header names
    all_header_info = []
    
    for header_name in unique_headers:
        all_header_info.append({
            'Header Name': header_name,
            'Found In Sheets': [
                sheet for sheet, headers in headers_by_sheet.items() 
                if any(h['name'] == header_name for h in headers)
            ]
        })
    
    if all_header_info:
        headers_summary_df = pd.DataFrame(all_header_info)
        st.write(f"Total unique columns: {len(headers_summary_df)}")
        st.dataframe(headers_summary_df, use_container_width=True)

    workbook_summary = {}

    for sheet in sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, dtype=str)

            sheet_summary = {
                "row_count": len(df),
                "col_count": len(df.columns),
                "first_columns": list(df.columns[:3]),  # only 3 columns
                "sample_row": df.head(1).to_dict(orient="records"),  # only 1 row
            }
            workbook_summary[sheet] = sheet_summary

            with st.expander(f"📄 Preview: {sheet}", expanded=False):
                st.dataframe(df.head(5), use_container_width=True)

        except Exception as e:
            st.error(f"Error reading sheet {sheet}: {e}")

    data_str = json.dumps(workbook_summary, default=str)

    # Show token usage
    token_count = count_tokens(data_str)
    st.write(f"📏 Approx. input size: {token_count} tokens")

    if "messages" not in st.session_state:
        st.session_state.messages = []

    data_payload = {
        "workbook_summary": workbook_summary,
        "pivots": all_pivots
    }
    data_str = json.dumps(data_payload, default=str)
    
    # ---------- Default Initial Prompt ----------
    if st.button("Get Initial AI Analysis"):
        user_prompt = {
    "role": "user",
    "content": (
        "Here is a preview of an Excel workbook with metadata and pivot tables.\n\n"
        f"{data_str}\n\n"
        "Please:\n"
        "- Summarize workbook structure (sheets, rows, columns)\n"
        "- Describe detected pivot tables, their source ranges, and what they indicate based on sample data\n"
        "- Provide 2–3 possible business insights from pivots\n"
        "⚠️ Keep answer under 250 words."
    ),
    }

        st.session_state.messages = [
            {"role": "system", "content": "You are an expert data analyst."},
            user_prompt,
        ]

        headers = {
            "Authorization": f"Bearer {API_KEY}",
            "HTTP-Referer": "http://localhost",
            "X-Title": "Excel AI Insights",
        }
        payload = {
            "model": "mistralai/mixtral-8x7b-instruct",
            "messages": st.session_state.messages,
        }

        with st.spinner("Analyzing workbook..."):
            response = requests.post(BASE_URL, headers=headers, json=payload)

            try:
                result = response.json()
            except Exception as e:
                st.error(f"❌ Failed to parse JSON response: {e}")
                st.write(response.text)
                st.stop()

            with st.expander("📡 Raw API Response", expanded=False):
                st.json(result)

            answer = None
            if "choices" in result:
                answer = result["choices"][0]["message"]["content"]
            elif "error" in result:
                st.error(f"❌ API returned error: {result['error']}")
            else:
                st.error("❌ Unexpected API response format")

            if answer:
                st.subheader("📈 Initial AI Analysis")
                st.write(answer)

    # ---------- Custom Prompt ----------
    st.subheader("💬 Custom Prompt to AI")
    custom_prompt = st.text_area("Enter your custom question about the workbook:")

    if st.button("Send Custom Prompt"):
        if not custom_prompt.strip():
            st.warning("⚠️ Please enter a prompt first.")
        else:
            custom_message = {"role": "user", "content": f"{custom_prompt}\n\nWorkbook context:\n{data_str}"}
            st.session_state.messages.append(custom_message)

            headers = {
                "Authorization": f"Bearer {API_KEY}",
                "HTTP-Referer": "http://localhost",
                "X-Title": "Excel AI Insights",
            }
            payload = {
                "model": "mistralai/mixtral-8x7b-instruct",
                "messages": st.session_state.messages,
            }

            with st.spinner("Sending custom query..."):
                response = requests.post(BASE_URL, headers=headers, json=payload)

                try:
                    result = response.json()
                except Exception as e:
                    st.error(f"❌ Failed to parse JSON response: {e}")
                    st.write(response.text)
                    st.stop()

                with st.expander("📡 Raw API Response (Custom)", expanded=False):
                    st.json(result)

                answer = None
                if "choices" in result:
                    answer = result["choices"][0]["message"]["content"]
                elif "error" in result:
                    st.error(f"❌ API returned error: {result['error']}")
                else:
                    st.error("❌ Unexpected API response format")

                if answer:
                    st.subheader("📝 Custom AI Response")
                    st.write(answer)
