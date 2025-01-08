import streamlit as st
from mailmerge import MailMerge
import pandas as pd
import os
from io import BytesIO
from zipfile import ZipFile

# Function to transform column names
def transform_column_names(col):
    # Replace `\n` in the middle of the string with `_`
    col = col.replace('\n', '_') if '\n' in col and not col.startswith('\n') else col
    
    # Replace `\n` at the start with `M__`
    col = col.replace('\n', 'M__', 1) if col.startswith('\n') else col
    
    # Remove brackets and their content, e.g., 13(2) -> 132
    col = ''.join(char for char in col if char not in '()')
    # Remove `/` and `.`
    col = col.replace('/', '').replace('.', '')
    
    return col

# Function to perform mail merge
def perform_mail_merge(word_file, excel_file):
    # Load the Word file and Excel sheet
    data = pd.read_excel(excel_file)
    data.columns = data.columns.str.replace(' ', '_')
    # Apply the transformation to column names
    data.columns = [transform_column_names(col) for col in data.columns]
    # Identify datetime columns and format them to DD-MM-YYYY
    for col in data.columns:
        if pd.api.types.is_datetime64_any_dtype(data[col]):  # Check if the column is a datetime type
            data[col] = pd.to_datetime(data[col]).dt.strftime('%d-%m-%Y')  # Format: DD-MM-YYYY
                
    # Preprocess all columns for unnecessary spaces and line breaks
    for col in data.columns:
    # Clean addresses and text columns (if column contains strings)
        if data[col].dtype == 'object':  # Check if the column is string type
            data[col] = data[col].str.strip()  # Remove leading/trailing spaces
            data[col] = data[col].str.replace(r'\s+', ' ', regex=True)  # Replace multiple spaces with a single space
            data[col] = data[col].str.replace(r'\n|\r', ' ', regex=True)  # Replace newlines and carriage returns with a space

        # List to store the generated documents
        output_files = []

        for index, row in data.iterrows():
            # Prepare a dictionary for placeholders and their corresponding values
            merge_fields = {field: str(row[field]) for field in row.index}
            with MailMerge(word_file) as document:
                # Merge fields
                document.merge(**merge_fields)

                # Save the generated document to a BytesIO object
                output_stream = BytesIO()
                document.write(output_stream)
                output_stream.seek(0)

                output_files.append((f"Document_{index + 1}.docx", output_stream))

    return output_files

# Streamlit App
st.markdown("""
    <style>
    body {
        background-color: black;
    }
    .title-style {
        font-size: 48px;
        color: white;
        background-color: #D2B48C;  
        padding: 20px;
        text-align: center;
        border-radius: 10px;
        width: 100%;
        margin: 0 auto;
    }
    </style>
    <div class="title-style">
        Letters Drafting App
    </div>
    """, unsafe_allow_html=True)

# Upload Word template file
word_file = st.file_uploader("📤 Upload Word Template (.docx)", type=["docx"])

# Upload Excel file
excel_file = st.file_uploader("📤 Upload Excel File (.xlsx)", type=["xlsx"])


if word_file and excel_file:
    if st.button("Generate Documents"):
        with st.spinner("Generating documents..."):
            try:
                # Perform mail merge
                output_files = perform_mail_merge(word_file, excel_file)
                # Create a ZIP file to download all documents
                zip_buffer = BytesIO()
                with ZipFile(zip_buffer, "w") as zip_file:
                    for filename, file_stream in output_files:
                        zip_file.writestr(filename, file_stream.getvalue())
                zip_buffer.seek(0)

                # Provide download link for the ZIP file
                st.success("Documents generated successfully!")
                st.download_button(
                    label="Download All Documents",
                    data=zip_buffer,
                    file_name="Generated_Documents.zip",
                    mime="application/zip",
                )

            except Exception as e:
                st.error(f"An error occurred: {e}")