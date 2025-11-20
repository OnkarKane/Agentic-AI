# import os
# import pandas as pd
# from google import genai
# from google.genai import types
# from io import BytesIO
# import re # Needed for parsing the Markdown table
# import xlsxwriter # Dependency needed for pd.ExcelWriter engine='xlsxwriter'

# # --- 1. CORE FUNCTIONS ---

# # Function to automatically adjust column widths in the final Excel file
# def auto_fit_excel_columns(df: pd.DataFrame, file_path: str):
#     """Saves the DataFrame to Excel and adjusts column widths using xlsxwriter."""
    
#     # Use the ExcelWriter engine (xlsxwriter)
#     writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    
#     # Write the dataframe to a specific sheet
#     df.to_excel(writer, sheet_name='DataExtraction', index=False)
    
#     # Get the xlsxwriter workbook and worksheet objects.
#     workbook  = writer.book
#     worksheet = writer.sheets['DataExtraction']

#     # Iterate through each column and set the width
#     for i, col in enumerate(df.columns):
#         # Calculate the maximum length of the column data
#         # Check all elements (including the header name)
#         header_len = len(col)
#         # Use str.len().max() to find the longest string length in the column
#         max_len = df[col].astype(str).str.len().max()
        
#         # Use the maximum of the data length or the header length, plus padding
#         column_width = max(max_len if pd.notna(max_len) else 0, header_len) + 5 # Added a bit more padding (5)
        
#         # Limit the width to a reasonable maximum (e.g., 80) for the Comments column readability
#         column_width = min(column_width, 80) 
        
#         # Set the column width (col_num, col_num, width)
#         worksheet.set_column(i, i, column_width)

#     # Close the Pandas Excel writer and output the Excel file.
#     writer.close()


# def create_system_prompt() -> str:
#     """Generates the final, master prompt with maximum consolidation in the Education section."""
    
#     return """
#     You are the **Data Extraction and Structuring Agent (DESA)**. Your task is to meticulously analyze the raw, unstructured PDF content and convert **all** explicit data into a highly organized, comprehensive Markdown Table.

#     ### 🎯 Goal
#     Your output MUST strictly follow the provided **SEMANTIC NAMING CONVENTIONS** and achieve **Non-Redundant 100% Capture**.

#     ### 📜 Output & Creation Rules (STRICT COMPLIANCE REQUIRED)

#     1.  **Output Format (CRITICAL):** Your ENTIRE output MUST be a single Markdown Table with exactly three header columns: **| Key | Value | Comments |**. NO other text, headings, explanations, or code blocks are permitted.

#     2.  **Key Naming (SEMANTIC ENFORCEMENT):**
#         *   **CRITICAL NAME DECOMPOSITION RULE:** The full name (e.g., "Vijay Kumar") MUST be split into two distinct rows: **Key: "First Name" and Key: "Last Name"**. This is a non-negotiable instruction.
#         *   **Employment:** Keys MUST use semantic names like **"Current Job," "Previous Job,"** etc. (and their related details).
#         *   **Education:** Keys MUST use semantic names like **"12th Standard," "Undergraduate Degree," "Postgraduate Degree."**
#         *   **Consolidation (Education/Skills):** Subjects, Ranking/Class Position, and Honors Status **MUST** be placed in the **Comments** of the relevant row, **NOT** as separate Key:Value pairs.
#         *   **Skills:** Use **"Skill X"** and **"Skill X Experience"** Keys.

#     3.  **Comments Field (Sub-Facts and 100% Capture Pool):**
#         *   **Conditionality:** A 'Comment' is **NOT mandatory** for every row.
#         *   **Content (Priority 1 - Sub-Fact/Metadata):** Use the Comment column to capture all metadata, descriptive clauses, and secondary facts.
#         *   **100% Capture:** If a sentence contains no primary Key:Value fact, it MUST be placed as a Comment in the most logically relevant row to ensure **100% of the document content is captured**.
#         *   **Non-Redundancy:** A sentence/clause should appear as a Comment only once across the entire table.
#         *   **Preserve Original Language:** Do NOT paraphrase, summarize, or alter the wording.

#     4.  **Value Field:** The 'Value' must be the most concise, clean, extracted data point.
#     5.  **Name:** This field must be divided into two parts first name and last name.
#     6.  ALL DATA MUST BE EXTRACTED without fail and comply with the above.

#     ### ⚙️ Required Output Table Headers
#     | Key | Value | Comments |
#     """

# def parse_markdown_table(markdown_text: str) -> pd.DataFrame:
#     """Parses the raw Markdown table string into a Pandas DataFrame."""
    
#     lines = [line.strip() for line in markdown_text.strip().split('\n')]
#     if not lines: return pd.DataFrame()
#     header_index, separator_index = -1, -1
    
#     for i, line in enumerate(lines):
#         if '|' in line and header_index == -1:
#             header_index = i
#         elif all(c in ('-', '|', ' ', ':') for c in line) and separator_index == -1 and header_index != -1:
#             separator_index = i
#             break
            
#     if header_index == -1 or separator_index == -1:
#         # Fallback to try to extract from a code block if the direct table failed
#         match = re.search(r'```(?:markdown)?\n(.*?)```', markdown_text, re.DOTALL)
#         if match:
#             print("-> Attempting to parse text from an unexpected code block...")
#             # Recurse with cleaner text
#             return parse_markdown_table(match.group(1)) 
        
#         print("🚨 ERROR: Could not find valid Markdown table structure in LLM output.")
#         return pd.DataFrame()
        
#     headers = [h.strip() for h in lines[header_index].strip('|').split('|')]
#     data_lines = lines[separator_index + 1:]
#     data = []
    
#     for line in data_lines:
#         if not line.strip(): continue
#         row_values = [col.strip() for col in line.strip('|').split('|')]
        
#         # Ensure we have the correct number of columns
#         if len(row_values) == len(headers):
#             data.append(row_values)

#     return pd.DataFrame(data, columns=headers)


# def run_gemini_extraction_markdown(pdf_data: BytesIO, model_name: str = "gemini-2.5-flash") -> pd.DataFrame:
#     """
#     Sends the raw PDF content to Gemini and uses the Markdown Table Parser.
#     """
    
#     if not os.getenv("GEMINI_API_KEY"):
#         print("🚨 ERROR: GEMINI_API_KEY environment variable not set.")
#         return pd.DataFrame()

#     try:
#         client = genai.Client()
#     except Exception as e:
#         print(f"🚨 ERROR: Error initializing Gemini client: {e}")
#         return pd.DataFrame()
        
#     system_prompt = create_system_prompt()
#     print(f"-> Sending document (PDF) directly to Gemini {model_name} for Markdown extraction...")

#     pdf_part = types.Part.from_bytes(
#         data=pdf_data.getvalue(),
#         mime_type='application/pdf'
#     )
    
#     response_text = None
#     try:
#         response = client.models.generate_content(
#             model=model_name,
#             contents=[system_prompt, pdf_part],
#         )
#         response_text = response.text

#         final_df = parse_markdown_table(response_text)
        
#         if not final_df.empty:
#             final_df = final_df[['Key', 'Value', 'Comments']]
        
#         return final_df

#     except Exception as e:
#         print(f"🚨 ERROR: Gemini API or Parsing Error: {e}")
#         if response_text:
#              print(f"Raw LLM output (for debugging):\n{response_text}")
#         return pd.DataFrame()


# # --- 4. MAIN EXECUTION BLOCK ---

# if __name__ == "__main__":
    
#     PDF_FILE_PATH = "C:/Users/kaneo/Downloads/Assignment/Data Input.pdf" 
#     OUTPUT_FILE = "Output.xlsx"
    
#     print("\n--- Starting AI-Powered Data Extraction Task (Non-Streamlit, Markdown Approach) ---")

#     # 1. Read the PDF file from the local path
#     try:
#         with open(PDF_FILE_PATH, 'rb') as f:
#             pdf_data = BytesIO(f.read())
#     except FileNotFoundError:
#         print(f"🚨 ERROR: Input file not found at '{PDF_FILE_PATH}'. Please check the path.")
#         exit()
#     except Exception as e:
#         print(f"🚨 ERROR: Failed to read PDF file: {e}")
#         exit()
        
#     # 2. Perform the extraction
#     final_df = run_gemini_extraction_markdown(pdf_data)
    
#     # 3. Save the output
#     if not final_df.empty:
#         # --- NEW CODE: CALL THE AUTO-FIT FUNCTION ---
#         auto_fit_excel_columns(final_df, OUTPUT_FILE)
        
#         print(f"\n✅ SUCCESS! Final structured data saved to {OUTPUT_FILE}")
#         print("-" * 50)
#         print(f"Total data points dynamically captured: {len(final_df)}")
#         print("First 5 rows of the output (Excel columns auto-fit):")
#         print(final_df.head())
#         print("-" * 50)
#     else:
#         print("\n❌ Extraction failed. Check the error messages above.")

import os
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from io import BytesIO
import re
import xlsxwriter

# --- 1. CORE FUNCTIONS ---

def auto_fit_excel_columns(df: pd.DataFrame, file_path: BytesIO):
    """Saves the DataFrame to an in-memory BytesIO object and adjusts column widths using xlsxwriter."""
    
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='DataExtraction', index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets['DataExtraction']

    for i, col in enumerate(df.columns):
        header_len = len(col)
        max_len = df[col].astype(str).str.len().max()
        
        column_width = max(max_len if pd.notna(max_len) else 0, header_len) + 5 
        column_width = min(column_width, 80) 
        
        worksheet.set_column(i, i, column_width)

    writer.close()


def create_system_prompt() -> str:
    """Generates the final, master prompt (Your final version)."""
    
    return """
    You are the **Data Extraction and Structuring Agent (DESA)**. Your task is to meticulously analyze the raw, unstructured PDF content and convert **all** explicit data into a highly organized, comprehensive Markdown Table.

    ### 🎯 Goal
    Your output MUST strictly follow the provided **SEMANTIC NAMING CONVENTIONS** and achieve **Non-Redundant 100% Capture**.

    ### 📜 Output & Creation Rules (STRICT COMPLIANCE REQUIRED)

    1.  **Output Format (CRITICAL):** Your ENTIRE output MUST be a single Markdown Table with exactly three header columns: **| Key | Value | Comments |**. NO other text, headings, explanations, or code blocks are permitted.

    2.  **Key Naming (SEMANTIC ENFORCEMENT):**
        *   **CRITICAL NAME DECOMPOSITION RULE:** The full name (e.g., "Vijay Kumar") MUST be split into two distinct rows: **Key: "First Name" and Key: "Last Name"**. This is a non-negotiable instruction.
        *   **Employment:** Keys MUST use semantic names like **"Current Job," "Previous Job,"** etc. (and their related details).
        *   **Education:** Keys MUST use semantic names like **"12th Standard," "Undergraduate Degree," "Postgraduate Degree."**
        *   **Consolidation (Education/Skills):** Subjects, Ranking/Class Position, and Honors Status **MUST** be placed in the **Comments** of the relevant degree/score row.
        *   **Skills:** Use **"Skill X"** and **"Skill X Experience"** Keys.

    3.  **Comments Field (Sub-Facts and 100% Capture Pool):**
        *   **Conditionality:** A 'Comment' is **NOT mandatory** for every row.
        *   **Content (Priority 1 - Sub-Fact/Metadata):** Use the Comment column to capture all metadata, descriptive clauses, and secondary facts.
        *   **100% Capture:** If a sentence contains no primary Key:Value fact, it MUST be placed as a Comment in the most logically relevant row to ensure **100% of the document content is captured**.
        *   **Non-Redundancy:** A sentence/clause should appear as a Comment only once across the entire table.
        *   **Preserve Original Language:** Do NOT paraphrase, summarize, or alter the wording.

    4.  **Value Field:** The 'Value' must be the most concise, clean, extracted data point.
    5.  **Name:** This field must be divided into two parts first name and last name.
    6.  ALL DATA MUST BE EXTRACTED without fail and comply with the above.

    ### ⚙️ Required Output Table Headers
    | Key | Value | Comments |
    """

def parse_markdown_table(markdown_text: str) -> pd.DataFrame:
    """Parses the raw Markdown table string into a Pandas DataFrame."""
    
    lines = [line.strip() for line in markdown_text.strip().split('\n')]
    if not lines: return pd.DataFrame()
    header_index, separator_index = -1, -1
    
    for i, line in enumerate(lines):
        if '|' in line and header_index == -1:
            header_index = i
        elif all(c in ('-', '|', ' ', ':') for c in line) and separator_index == -1 and header_index != -1:
            separator_index = i
            break
            
    if header_index == -1 or separator_index == -1:
        match = re.search(r'```(?:markdown)?\n(.*?)```', markdown_text, re.DOTALL)
        if match:
            # Attempt to parse text from an unexpected code block
            return parse_markdown_table(match.group(1)) 
        
        return pd.DataFrame()
        
    headers = [h.strip() for h in lines[header_index].strip('|').split('|')]
    data_lines = lines[separator_index + 1:]
    data = []
    
    for line in data_lines:
        if not line.strip(): continue
        row_values = [col.strip() for col in line.strip('|').split('|')]
        
        if len(row_values) == len(headers):
            data.append(row_values)

    return pd.DataFrame(data, columns=headers)


def run_gemini_extraction_markdown(pdf_data: BytesIO, client: genai.Client, model_name: str = "gemini-2.5-flash") -> pd.DataFrame:
    """
    Sends the raw PDF content to Gemini and uses the Markdown Table Parser.
    """
    
    system_prompt = create_system_prompt()
    st.info(f"-> Sending document (PDF) directly to Gemini {model_name} for Markdown extraction...")

    pdf_part = types.Part.from_bytes(
        data=pdf_data.getvalue(),
        mime_type='application/pdf'
    )
    
    response_text = None
    try:
        response = client.models.generate_content(
            model=model_name,
            contents=[system_prompt, pdf_part],
        )
        response_text = response.text

        final_df = parse_markdown_table(response_text)
        
        if not final_df.empty:
            final_df = final_df[['Key', 'Value', 'Comments']]
        
        return final_df

    except Exception as e:
        st.error(f"🚨 ERROR: Gemini API or Parsing Error: {e}")
        if response_text:
             st.code(f"Raw LLM output (for debugging):\n{response_text}")
        return pd.DataFrame()


@st.cache_resource
def get_gemini_client():
    """Initializes and returns the Gemini client using st.secrets or env var."""
    # Prioritizes st.secrets for cloud deployment, falls back to env var for local testing
    api_key = os.getenv("GEMINI_API_KEY") or st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        return None
    try:
        return genai.Client(api_key=api_key)
    except Exception as e:
        st.error(f"Error initializing Gemini client: {e}")
        return None

# --- 2. STREAMLIT APPLICATION LAYOUT ---

st.set_page_config(page_title="AI-Powered Document Structuring Demo", layout="wide")

st.title("📄 AI-Powered Document Structuring & Data Extraction Demo")
st.subheader("Transforms unstructured PDF into a structured Excel output using Gemini 2.5 Flash.")
st.markdown("---")

client = get_gemini_client()

if not client:
    st.error("🚨 **API KEY MISSING:** Please set the `GEMINI_API_KEY` in your environment variables or Streamlit secrets.")
    
uploaded_file = st.file_uploader(
    "Upload the Data Input.pdf file:", 
    type="pdf", 
    disabled=(client is None),
    help="The PDF is sent directly to the Gemini API for multimodal extraction."
)

if uploaded_file is not None and client:
    st.success(f"File uploaded: {uploaded_file.name}")
    
    if st.button("Start Extraction", type="primary"):
        
        pdf_data = BytesIO(uploaded_file.read())
        
        with st.spinner("⏳ Running Multimodal Extraction via Gemini 2.5 Flash... (Approx. 10-30 seconds)"):
            final_df = run_gemini_extraction_markdown(pdf_data, client)
        
        if not final_df.empty:
            st.success(f"✅ Extraction Complete! Dynamically captured {len(final_df)} data points.")
            st.markdown("---")
            st.header("Extracted Data Table Preview")

            # Display the result table
            st.dataframe(final_df, use_container_width=True)
            
            # Create Excel in memory with Auto-Fit columns
            excel_buffer = BytesIO()
            auto_fit_excel_columns(final_df, excel_buffer)
            excel_buffer.seek(0)

            st.download_button(
                label="Download Final Output.xlsx (Auto-Fit Columns)",
                data=excel_buffer,
                file_name="Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary"
            )
        else:
            st.error("Extraction failed. Please check the console/error messages for API or parsing issues.")

# --- INSTRUCTIONS FOR HOSTING ---
st.markdown("""
---
### 📦 Deliverables Checklist

*   **Source Code (Deliverable #1):** This `app.py` file goes into your GitHub repository.
*   **Live Demo (Deliverable #2):** Run this file locally (`streamlit run app.py`) or deploy the repository to Streamlit Community Cloud.
*   **Final Generated Output (Deliverable #3):** The `Output.xlsx` file generated by the demo.
""")