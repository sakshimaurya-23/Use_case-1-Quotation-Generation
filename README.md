# Use_Case 1: Quotation-Generation

A Python-based solution for extracting details from `.msg` email files, matching them with a master Excel sheet, and generating quotations using IBM WatsonX AI. This tool streamlines the quotation process, ensuring accuracy and consistency.

---

## Features

- **Email Parsing**: Extracts and preprocesses email content from `.msg` files.
- **Data Extraction**: Extracts structured details and tables from emails using IBM WatsonX AI.
- **Data Matching**: Matches extracted data with a master Excel file, appending relevant pricing details.
- **Quotation Generation**: Produces formatted quotations and detailed Excel summaries.
- **Interactive UI**: Streamlit-based interface for easy file uploads and result previews.

---

## Prerequisites

1. **Python Libraries**:
   Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. **Environment Variables**:
   - `WATSONX_URL`: URL for IBM WatsonX API.
   - `WATSONX_APIKEY`: API key for IBM WatsonX.

3. **Input Files**:
   - `.msg` files containing email data.
   - A master Excel sheet with pricing details.

---

## How to Run

1. **Prepare the Environment**:
   - Install required libraries.
   - Set environment variables for WatsonX API.

2. **Run the Application**:
   - Launch Streamlit:
     ```bash
     streamlit run app.py
     ```

3. **Use the UI**:
   - Upload `.msg` files.
   - View extracted data, matched details, and generated quotations.

---

## Key Functions

- **`extract_details_with_llm`**: Extracts structured details from email bodies.
- **`extract_table_with_llm`**: Extracts and processes tables from email content.
- **`match_with_master_excel`**: Matches extracted data with the master Excel sheet.
- **`generate_quotation_content`**: Creates formatted quotations and summaries.





