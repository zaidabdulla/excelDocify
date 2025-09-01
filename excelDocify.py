import streamlit as st
import pandas as pd
import requests#from dotenv import load_dotenv
import os
import json
import tiktoken  # counts tokens safely
from dotenv import load_dotenv

from openpyxl import load_workbook

from openpyxl.worksheet.table import Table
from openpyxl.utils import range_boundaries

import zipfile
import xml.etree.ElementTree as ET

# Load API key
load_dotenv()
API_KEY = os.getenv("OPENROUTER_API_KEY")

BASE_URL = "https://openrouter.ai/api/v1/chat/completions"

st.title("📊 Excel to AI Insights (OpenRouter)")

def extract_pivot_metadata_fast(file_path):
    """
    Extract pivot table metadata (pivot name, source range, source sheet)
    by reading XML directly from the .xlsx (faster than openpyxl).
    Returns (list of dicts, combined_text).
    """
    pivots = []

    with zipfile.ZipFile(file_path, "r") as z:
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        # 🔹 Step 1: Map pivotCacheId -> pivotTable name
        pivot_name_map = {}
        pivot_table_files = [f for f in z.namelist() if f.startswith("xl/pivotTables/pivotTable")]
        for pt_file in pivot_table_files:
            with z.open(pt_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                pivot_name = root.attrib.get("name", pt_file.split("/")[-1])
                cache_id = root.attrib.get("cacheId")  # links to pivotCacheDefinition
                if cache_id:
                    pivot_name_map[cache_id] = pivot_name

        # 🔹 Step 2: Read pivotCacheDefinitions (source info)
        pivot_cache_files = [
            f for f in z.namelist() if f.startswith("xl/pivotCache/pivotCacheDefinition")
        ]

        for idx, cache_file in enumerate(pivot_cache_files, start=1):
            with z.open(cache_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()

                ws_source = root.find(".//main:cacheSource/main:worksheetSource", ns)

                cache_id = str(idx)  # pivot caches are usually 1-based indexed
                pivot_name = pivot_name_map.get(cache_id, f"Pivot_{idx}")

                pivots.append({
                    "pivot_name": pivot_name,
                    "source_sheet": ws_source.attrib.get("sheet", "Unknown") if ws_source is not None else "Unknown",
                    "source_range": ws_source.attrib.get("ref", "Unknown") if ws_source is not None else "Unknown",
                })

    # 🔹 Step 3: Combine into single string for OpenAI
    combined_text = "\n".join(
        [f"- {p['pivot_name']} (Sheet: {p['source_sheet']}, Range: {p['source_range']})" for p in pivots]
    )

    return pivots, combined_text

def extract_table_metadata(file_path):
    """
    Extract Excel table metadata:
    - Table name
    - Table range
    - Table sheet name
    Returns (list of dicts, combined_text).
    """
    tables = []

    with zipfile.ZipFile(file_path, "r") as z:
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        # Step 1: Map tableId -> sheet name (from worksheet.xml.rels)
        table_sheet_map = {}
        sheet_files = [f for f in z.namelist() if f.startswith("xl/worksheets/sheet") and f.endswith(".xml")]
        rel_files = [f for f in z.namelist() if f.startswith("xl/worksheets/_rels/sheet") and f.endswith(".xml.rels")]

        # Map rel -> sheet
        for sheet_file, rel_file in zip(sheet_files, rel_files):
            sheet_name = sheet_file.split("/")[-1].replace(".xml", "")
            with z.open(rel_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                for rel in root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                    target = rel.attrib.get("Target", "")
                    if "tables/table" in target:
                        table_id = target.split("table")[-1].split(".xml")[0]
                        table_sheet_map[table_id] = sheet_name

        # Step 2: Extract table definitions
        table_files = [f for f in z.namelist() if f.startswith("xl/tables/table")]

        for tfile in table_files:
            table_id = tfile.split("table")[-1].split(".xml")[0]
            with z.open(tfile) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                table_name = root.attrib.get("name", f"Table_{table_id}")
                table_range = root.attrib.get("ref", "Unknown")
                table_sheet = table_sheet_map.get(table_id, "Unknown")

                tables.append({
                    "table_name": table_name,
                    "table_range": table_range,
                    "table_sheet": table_sheet,
                })

    # Step 3: Prepare combined text for AI summary
    combined_text = "\n".join(
        [f"- {t['table_name']} (Sheet: {t['table_sheet']}, Range: {t['table_range']})"
         for t in tables]
    )

    return tables, combined_text

def extract_data_validation_metadata(file_path):
    """
    Extract Data Validation rules metadata:
    - Rule name
    - Rule type (list, whole, decimal, date, etc.)
    - Operator (between, equal, lessThan, etc.)
    - Criteria (interpreted from formula1 & formula2)
    - Range (sqref)
    - Sheet name (actual Excel name, not sheet1.xml)
    Returns: (list of dicts, combined_text).
    """
    validations = []

    with zipfile.ZipFile(file_path, "r") as z:
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        # 1️⃣ Parse workbook.xml to get real sheet names
        sheet_id_map = {}
        with z.open("xl/workbook.xml") as f:
            tree = ET.parse(f)
            root = tree.getroot()

            for sheet in root.findall(".//main:sheets/main:sheet", ns):
                sheet_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                sheet_name = sheet.attrib.get("name")
                sheet_rid = sheet.attrib.get("sheetId")
                sheet_id_map[sheet_rid] = sheet_name

        # 2️⃣ Now read worksheets and map them back
        sheet_files = [f for f in z.namelist() if f.startswith("xl/worksheets/sheet") and f.endswith(".xml")]

        for sheet_file in sheet_files:
            # Extract numeric sheet index (sheet1.xml → "1")
            sheet_index = sheet_file.split("/")[-1].replace("sheet", "").replace(".xml", "")
            sheet_name = sheet_id_map.get(sheet_index, f"Sheet{sheet_index}")

            with z.open(sheet_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()

                for dv in root.findall(".//main:dataValidation", ns):
                    dv_type = dv.attrib.get("type", "unknown")
                    dv_range = dv.attrib.get("sqref", "Unknown")
                    operator = dv.attrib.get("operator", None)

                    # Extract formulas
                    f1 = dv.find("main:formula1", ns)
                    f2 = dv.find("main:formula2", ns)
                    formula1 = f1.text if f1 is not None else None
                    formula2 = f2.text if f2 is not None else None

                    # Build readable criteria
                    criteria = None
                    if dv_type in ["whole", "decimal", "date", "time"]:
                        if operator == "between":
                            criteria = f"Value between {formula1} and {formula2}"
                        elif operator == "notBetween":
                            criteria = f"Value not between {formula1} and {formula2}"
                        elif operator == "equal":
                            criteria = f"Value = {formula1}"
                        elif operator == "notEqual":
                            criteria = f"Value ≠ {formula1}"
                        elif operator == "lessThan":
                            criteria = f"Value < {formula1}"
                        elif operator == "lessThanOrEqual":
                            criteria = f"Value ≤ {formula1}"
                        elif operator == "greaterThan":
                            criteria = f"Value > {formula1}"
                        elif operator == "greaterThanOrEqual":
                            criteria = f"Value ≥ {formula1}"
                        else:
                            criteria = f"Formula1={formula1}, Formula2={formula2}"
                    elif dv_type == "list":
                        criteria = f"Allowed values from list: {formula1}"
                    elif dv_type == "custom":
                        criteria = f"Custom formula: {formula1}"
                    else:
                        criteria = f"Formula1={formula1}, Formula2={formula2}"

                    # Generate "name" for the rule
                    rule_name = f"{dv_type}_{operator or ''}_{formula1 or ''}".strip("_")

                    validations.append({
                        "rule_name": rule_name,
                        "rule_type": dv_type,
                        "operator": operator,
                        "range": dv_range,
                        "sheet": sheet_name,  # ✅ Real Excel sheet name
                        "formula1": formula1,
                        "formula2": formula2,
                        "criteria": criteria
                    })

    # Combine into text for AI summary
    combined_text = "\n".join(
        [f"- {v['rule_name']} (Type: {v['rule_type']}, Operator: {v['operator']}, "
         f"Sheet: {v['sheet']}, Range: {v['range']}, Criteria: {v['criteria']})"
         for v in validations]
    )

    return validations, combined_text

def extract_chart_metadata(file_path):
    """
    Extract chart metadata (chart type, title, sheet name).
    Returns (list of dicts, combined_text).
    """
    charts = []

    with zipfile.ZipFile(file_path, "r") as z:
        ns = {
            "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        }

        # map drawings to sheets
        sheet_files = [f for f in z.namelist() if f.startswith("xl/worksheets/sheet") and f.endswith(".xml")]
        rels = {f: f.replace("worksheets", "worksheets/_rels") + ".rels" for f in sheet_files}

        for sheet_file in sheet_files:
            sheet_name = sheet_file.split("/")[-1].replace(".xml", "")

            # open worksheet XML
            with z.open(sheet_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()

                # Look for <drawing r:id="rIdX">
                drawing_elems = root.findall(".//main:drawing", ns)
                if not drawing_elems:
                    continue

                # open sheet relationships
                rel_file = rels[sheet_file]
                if rel_file not in z.namelist():
                    continue
                with z.open(rel_file) as rf:
                    rel_tree = ET.parse(rf)
                    rel_root = rel_tree.getroot()

                    # map rId -> target
                    rel_map = {
                        r.attrib["Id"]: r.attrib["Target"]
                        for r in rel_root.findall(".//", {"": "http://schemas.openxmlformats.org/package/2006/relationships"})
                        if "Id" in r.attrib
                    }

                # now resolve drawing -> chart
                for d in drawing_elems:
                    rId = d.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    if not rId or rId not in rel_map:
                        continue
                    drawing_target = rel_map[rId].replace("..", "xl")

                    if drawing_target not in z.namelist():
                        continue

                    # read drawing xml
                    with z.open(drawing_target) as df:
                        d_tree = ET.parse(df)
                        d_root = d_tree.getroot()

                        # look for chart references
                        for c_elem in d_root.findall(".//c:chart", ns):
                            c_rId = c_elem.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                            # map chart rels
                            d_rel_file = drawing_target.replace("drawings", "drawings/_rels") + ".rels"
                            if d_rel_file not in z.namelist():
                                continue
                            with z.open(d_rel_file) as drf:
                                dr_tree = ET.parse(drf)
                                dr_root = dr_tree.getroot()
                                chart_map = {
                                    r.attrib["Id"]: r.attrib["Target"]
                                    for r in dr_root.findall(".//", {"": "http://schemas.openxmlformats.org/package/2006/relationships"})
                                    if "Id" in r.attrib
                                }

                            if c_rId in chart_map:
                                chart_target = "xl/" + chart_map[c_rId].lstrip("/")
                                if chart_target not in z.namelist():
                                    continue

                                # open chart XML
                                with z.open(chart_target) as cf:
                                    c_tree = ET.parse(cf)
                                    c_root = c_tree.getroot()

                                    # detect chart type
                                    chart_type = "Unknown"
                                    for ct in ["barChart", "lineChart", "pieChart", "scatterChart", "areaChart"]:
                                        if c_root.find(f".//c:{ct}", ns) is not None:
                                            chart_type = ct.replace("Chart", "")
                                            break

                                    # detect title
                                    title_elem = c_root.find(".//c:title//a:t", ns)
                                    chart_title = title_elem.text if title_elem is not None else "Untitled"

                                    charts.append({
                                        "chart_type": chart_type,
                                        "chart_title": chart_title,
                                        "sheet": sheet_name
                                    })

    # Combine text for OpenAI summary
    combined_text = "\n".join(
        [f"- {c['chart_title']} (Type: {c['chart_type']}, Sheet: {c['sheet']})" for c in charts]
    )

    return charts, combined_text

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
    
    

    pivots, text_for_ai = extract_pivot_metadata_fast(uploaded_file)

    print("Extracted Pivots:", pivots)
    print("\nCombined Summary for AI:\n", text_for_ai)
 
    # ---- Pivot Table Detection ----
    tables, text_for_ai = extract_table_metadata(uploaded_file)

    print("Extracted Tables:", tables)
    print("\nCombined Summary for AI:\n", text_for_ai)
    
    validations, text_for_ai = extract_data_validation_metadata(uploaded_file)

    print("Validation Rules Found:", validations)
    print("\nCombined Summary for AI:\n", text_for_ai)

    charts, text_for_ai = extract_chart_metadata(uploaded_file)

    print("Charts Found:", charts)
    print("\nCombined Summary for AI:\n", text_for_ai)

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
        "pivots": pivots,
        "tables": tables
    }
    data_str = json.dumps(data_payload, default=str)

    # ---------- Default Initial Prompt ----------
    if st.button("Get Initial AI Analysis"):
        user_prompt = {
            "role": "user",
            "content": (
                "Here is a minimal preview of an Excel workbook with multiple sheets. "
                "Only metadata (rows/cols, first 3 columns) and 1 sample row per sheet are included.\n\n"
                f"{data_str}\n\n"
                "Please:\n"
                "- Summarize workbook structure (sheets, rows, columns)\n"
                "- Suggest possible next analysis steps\n"
                "- Give 3 high-level business insights\n"
                "⚠️ Keep answer under 200 words."
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
