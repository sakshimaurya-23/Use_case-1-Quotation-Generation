import streamlit as st
import extract_msg as em
from dotenv import load_dotenv
import pandas as pd
import re
import json
from io import StringIO
from ibm_watsonx_ai import Credentials
from ibm_watsonx_ai.foundation_models import ModelInference
from rapidfuzz import process, fuzz
import os
import requests

load_dotenv()

credentials = Credentials(
    url=os.getenv("WATSONX_URL"),
    api_key=os.getenv("WATSONX_APIKEY")  
)
parameters_1 = {
    "decoding_method": 'greedy',
    "max_new_tokens": 1500,
    "min_new_tokens": 5,
    "temperature": 0
}
project_id = os.getenv("project_id")
model_id="meta-llama/llama-3-2-90b-vision-instruct"

model = ModelInference(
    model_id=model_id,
    credentials=credentials,
    project_id=project_id,
    params=parameters_1
)


def extract_details_with_llm(msg_body):
    """
    Extract structured details such as Subject, Sender, Receiver, and Date from the email message body.
    """
    user_query = '''
    Extract the following details from the provided email content:

    1. **Our Ref**: Extract the reference number mentioned in the message.
    2. **Date**: Extract date in the email sent date in "Day, DD Month YYYY" format, if available in the email.
    3. **To**: Extract the recipient's name and contact mentioned in the body, for example, "kindly address to Abella Jake Yabut @ 64138413".
    4. **From**: Extract the sender's name from the body of the email we can get from the very firast line in the body of mail, for example, "Hi Lionel".
    5. **Subject/Prj Name**: Extract the project name or subject line mentioned in the message, such as "SSR2024-040: GMET-EDT Capacity Uplift".

    Output the extracted details in the following structured format without any additional explanations:

    **Our Ref**: [Extracted Reference Number]  
    **Date**: [Extracted Date]  
    **To**: [Recipient's Name and Contact]  
    **From**: [Sender's Name]  
    **Subject/Prj Name**: [Extracted Subject or Project Name]  
    '''
    prompt = f"{user_query}\n\nMessage:\n{msg_body}"
    
    with st.spinner("Processing extracted details..."):
        response = model.generate(prompt=prompt)
    extracted_details = response.get("results", [{}])[0].get("generated_text", "").strip()
    return extracted_details

def parse_llm_response(field, llm_response):
    """
    Parse specific field value from LLM response.
    """
    pattern = rf"{field}: (.+)"
    match = re.search(pattern, llm_response)
    return match.group(1) if match else None

def extract_table_with_llm(msg_body):
    """
    Extract table from the message content.
    """
    user_query = '''
    Extract the table from the following content. The table may be messy, unaligned, or embedded in an HTML format.
    Use the following column headers exactly as given:
    | Req. Ref. | Project | Site | Env. | Type | Items | Qty (GiB) |

    Do not add any extra text or explanations. Output only the table in clean markdown format.
'''
   
    prompt = f"{user_query}\n\nMessage:\n{msg_body}"
    with st.spinner("Processing extracted table..."):
        response = model.generate(prompt=prompt)
    table_output = response.get("results", [{}])[0].get("generated_text", "").strip()
    return table_output


def extract_msg_content(file_path):
    """ Extract message content from .msg files. """
    msg = em.Message(file_path)
    return msg.htmlBody.decode('utf-8', errors='ignore') if msg.htmlBody else msg.body.decode('utf-8', errors='ignore')

def preprocess_message(msg_body):
    """ Clean message content. """
    return re.sub(r"(UOB EMAIL DISCLAIMER.*$|CAUTION:.*?$)", "", msg_body, flags=re.IGNORECASE | re.DOTALL).strip()

def markdown_to_dataframe(markdown_table):
    try:
        return pd.read_csv(StringIO(markdown_table.replace("|", "").strip()), sep="\s{2,}", engine='python')
    except Exception as e:
        return f"Error parsing markdown table: {e}"

def match_with_master_excel(msg_df):
    master_excel_path = '/Users/sakshimaurya/Desktop/nxgen-o2c-main/Book3.xlsx'  # Update with the correct path
    master_df = pd.read_excel(master_excel_path)

    # Clean Excel column headers and string values
    master_df.columns = master_df.columns.str.strip()
    for col in ['Req. Ref.', 'Project', 'Site', 'Env.', 'Type', 'Description']:
        master_df[col] = master_df[col].astype(str).str.strip().str.lower()

    # Clean string values in .msg table
    for key in ['Req. Ref.', 'Project', 'Site', 'Env.', 'Type', 'Items']:
        msg_df[key] = msg_df[key].astype(str).str.strip().str.lower()

    # Match Rows and Append Unit Cost and Total Cost
    results = []
    for _, msg_row in msg_df.iterrows():
        # Filter Excel rows using composite keys
        filtered_df = master_df[
            (master_df['Req. Ref.'] == msg_row['Req. Ref.']) &
            (master_df['Project'] == msg_row['Project']) &
            (master_df['Site'] == msg_row['Site']) &
            (master_df['Env.'] == msg_row['Env.']) &
            (master_df['Type'] == msg_row['Type'])
        ]

        # Apply Fuzzy Matching on 'Items' and 'Description'
        if not filtered_df.empty:
            descriptions = filtered_df['Description'].tolist()
            match, score, _ = process.extractOne(msg_row['Items'], descriptions, scorer=fuzz.partial_ratio)
            matched_row = filtered_df[filtered_df['Description'] == match].iloc[0]

            result = msg_row
            result['Unit Cost'] = matched_row['Unit Cost']
            result['Total Cost'] = matched_row['Total Cost']
            result['Quote Reference #'] = matched_row['Quote Reference #']
            result['Matching Score'] = score
            results.append(result)
        else:
            # Add placeholders if no match is found
            result = msg_row
            result['Unit Cost'] = 'N/A'
            result['Total Cost'] = 'N/A'
            result['Quote Reference #'] = 'N/A'
            result['Matching Score'] = 0
            results.append(result)


    # Create Final DataFrame
    final_df = pd.DataFrame(results)
    return final_df


def generate_quotation_content(details, final_df):
    """Generate formatted quotation content as a single page."""
    # Parse the details using regex
    detail_lines = {}
    for line in details.split("\n"):
        if ":" in line:
            key, value = line.split(":", 1)
            detail_lines[key.strip()] = value.strip()

    # Extract details with fallbacks
    our_ref = detail_lines.get("**Our Ref**", "N/A")
    date_sent = detail_lines.get("**Date**", "N/A")
    to_client = detail_lines.get("**To**", "N/A")
    subject_name = detail_lines.get("**Subject/Prj Name**", "N/A")
    from_sender = detail_lines.get("**From**", "N/A")
    valid_til = "90 days from date of quotation"

    our_ref = final_df['Quote Reference #'].iloc[0] if 'Quote Reference #' in final_df.columns and not final_df.empty else "N/A"

    # Calculate Total Investment
    total_investment = final_df['Total Cost'].replace('N/A', 0).astype(float).sum()
    total_with_gst = round(total_investment * 1.08, 2)  # Applying 8% GST

    # Prepare Excel summary data
    additional_rows = pd.DataFrame({
        "Req. Ref.": ["", ""],
        "Project": ["", ""],
        "Site": ["", ""],
        "Env.": ["", ""],
        "Type": ["", ""],
        "Items": ["Total Investments", "Total Investments incl. 8% GST"],
        "Qty (GiB)": ["", ""],
        "Total Cost": [total_investment, total_with_gst]
    })
    summary_data = pd.concat([final_df.drop(columns=['Matching Score'], errors='ignore'), additional_rows], ignore_index=True)

    # Generate the single-page quotation content
    single_page = f"""
Our Ref: {our_ref}  
Date: {date_sent}  
Valid Til: {valid_til} 

To: {to_client}  
Platform: Open System Storage
Company: United Overseas Bank Limited  

From: {from_sender}  
Subject: {subject_name}
Company: S&I Systems Private Limited   
No Of Pages: 1  

-----------------------------------------------------------------------------------------

Dear {to_client.split('@')[0] if '@' in to_client else to_client},  

Thank you for giving S&I this opportunity to propose the following offer for your new infrastructure requirements.  
We hope that you will find our proposal favorable and do feel free to call us should you require any further clarifications or information.  

Investment Summary: 
-------------------
- Total Investment: {total_investment:.2f}  
- Total Investment (Incl. GST): {total_with_gst:.2f}  

For a detailed breakdown, please download the attached Excel file.

Delivery and Payment Terms  
- Validity: Price is valid for 90 days from date of quotation.  
- Delivery: 4 to 8 weeks upon receipt of order confirmation.  

GENERAL TERMS & CONDITIONS
- Payment: S&I shall invoice the Customer in accordance with the agreed payment schedule in this agreement, and the Customer agrees to abide by the payment schedule and pay promptly in full all due invoices (for undisputed invoices) within the agreed 30 days from date of invoice.  
- Taxes: The above prices are subject to prevailing GST at the date of purchase.  
- Title of Goods: Title of goods will remain with S&I until full payment is received.  
- Governing Law: This agreement shall be governed by and interpreted in accordance with the laws of the Republic of Singapore.  

The Sales quotation is governed by the terms and conditions defined in the MSA signed between UOB and S&I dated 29th Aug 2018.  
This quotation shall not be effective until executed by the Customer (via Purchase Order and/or Statement of Work (SOW)) and accepted by S&I. Subsequent amendments or changes to the details contained in this Agreement have to be in writing and signed by both S&I and the Customer.  

We thank you and hope to hear from you soon.  

Thanks & Best regards,  
S&I Systems Private Limited  
"""
    return single_page, summary_data



def extract_msg_content(file_path):
    """ Extract message content from .msg files. """
    msg = em.Message(file_path)
    return msg.htmlBody.decode('utf-8', errors='ignore') if msg.htmlBody else msg.body.decode('utf-8', errors='ignore')

def preprocess_message(msg_body):
    """ Clean message content. """
    return re.sub(r"(UOB EMAIL DISCLAIMER.*$|CAUTION:.*?$)", "", msg_body, flags=re.IGNORECASE | re.DOTALL).strip()

def markdown_to_dataframe(markdown_table):
    try:
        return pd.read_csv(StringIO(markdown_table.replace("|", "").strip()), sep="\s{2,}", engine='python')
    except Exception as e:
        return f"Error parsing markdown table: {e}"
    


def match_with_master_excel(msg_df):
    master_excel_path = '/Users/sakshimaurya/Desktop/nxgen-o2c-main/Book3.xlsx'  # Update with the correct path
    master_df = pd.read_excel(master_excel_path)

    # Clean Excel column headers and string values
    master_df.columns = master_df.columns.str.strip()

    # Match Rows and Append Unit Cost and Total Cost
    results = []
    for _, msg_row in msg_df.iterrows():
        # Filter Excel rows using composite keys
        filtered_df = master_df[
            (master_df['Req. Ref.'].str.strip().str.lower() == str(msg_row['Req. Ref.']).strip().lower()) &
            (master_df['Project'].str.strip().str.lower() == str(msg_row['Project']).strip().lower()) &
            (master_df['Site'].str.strip().str.lower() == str(msg_row['Site']).strip().lower()) &
            (master_df['Env.'].str.strip().str.lower() == str(msg_row['Env.']).strip().lower()) &
            (master_df['Type'].str.strip().str.lower() == str(msg_row['Type']).strip().lower())
        ]

        if not filtered_df.empty:
            for _, matched_row in filtered_df.iterrows():
                result = msg_row.to_dict()
                result['Req. Ref.'] = matched_row['Req. Ref.']
                result['Project'] = matched_row['Project']
                result['Site'] = matched_row['Site']
                result['Env.'] = matched_row['Env.']
                result['Type'] = matched_row['Type']
                result['Unit Cost'] = matched_row['Unit Cost']
                result['Total Cost'] = matched_row['Total Cost']
                result['Quote Reference #'] = matched_row['Quote Reference #']
                results.append(result)
        else:
            # Add placeholders if no match is found
            result = msg_row.to_dict()
            result['Unit Cost'] = 'N/A'
            result['Total Cost'] = 'N/A'
            result['Quote Reference #'] = 'N/A'
            results.append(result)

    # Create Final DataFrame
    final_df = pd.DataFrame(results)
    return final_df


# Streamlit UI
st.set_page_config(layout="wide")
st.title("Quotation Generation")

uploaded_msg_file = st.file_uploader("Upload .msg File", type=["msg"])

if uploaded_msg_file:
    temp_msg_path = "temp.msg"
    with open(temp_msg_path, "wb") as f:
        f.write(uploaded_msg_file.read())

    try:
        # Extract and clean message body
        st.info("Extracting message content...")
        msg_body = extract_msg_content(temp_msg_path)
        clean_body = preprocess_message(msg_body)
        # st.write(msg_body)

        # Extract details
        st.info("Extracting structured details...")
        details = extract_details_with_llm( msg_body)
        table_df = markdown_to_dataframe( details)
        st.code(details)
        #st.write(details)

        # Extract table
        st.info("Extracting table...")
        extracted_table = extract_table_with_llm( msg_body)
        table_df = markdown_to_dataframe(extracted_table)
        st.subheader("Extracted Table")
        st.dataframe(table_df)

        # Match table with master Excel
        st.info("Matching table with master Excel...")
        final_table = match_with_master_excel(table_df)
        st.subheader("Matched Table")
        st.dataframe(final_table)

        if st.button("Generate Quotation", key="generate_quotation"):
            st.info("Generating final quotation...")
            single_page, excel_df = generate_quotation_content(details, final_table)
            st.subheader("Quotation")
            st.text_area("Generated Quotation", single_page, height=600)

            # Create Excel for download
            excel_path = "investment_summary.xlsx"
            excel_df.to_excel(excel_path, index=False)
            with open(excel_path, "rb") as excel_file:
                st.download_button(
                    label="Investment Summary",
                    data=excel_file,
                    file_name="investment_summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")
    finally:
        os.remove(temp_msg_path)


