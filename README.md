# AI-Powered Document Structuring & Data Extraction Agent (DESA)

## 🎯 Assignment Objective
To design and implement an AI-backed solution that transforms content from an unstructured PDF document into a structured Excel output, strictly adhering to constraints on 100% data capture and preservation of original language.

## 🚀 Solution Architecture
This solution uses the **Gemini 2.5 Flash** model (via Google's `google-genai` API) in a multimodal context to directly process the PDF file. The LLM is instructed via a highly aggressive prompt to extract all facts into a rigid Markdown table format, bypassing brittle JSON schema issues. The Streamlit web application then parses the Markdown table and generates a final Excel file with auto-fitted columns.

## 🛠️ Local Setup & Usage

### Prerequisites
1.  **Clone the repository:**
    ```bash
    git clone https://github.com/OnkarKane/Agentic-AI
    ```
2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
3.  **Set API Key:** Set your Gemini API key as an environment variable.
    *   **macOS/Linux:** `export GEMINI_API_KEY="YOUR_KEY_HERE"`
    *   **Windows (cmd):** `set GEMINI_API_KEY="YOUR_KEY_HERE"`
    *   When setting the API KEY remove the ("")

### Running the Demo Locally 
1.  Ensure the API key is set in your current terminal session.
2.  Run the Streamlit application:
    ```bash
    streamlit run app.py
    ```
3.  The application will open in your browser, allowing you to upload `Data Input.pdf` and download the structured `Output.xlsx`.
