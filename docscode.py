import streamlit as st
import pandas as pd
from docx import Document
import io

def fill_blanks_for_multiple_rows(template_content, df):
    """
    Fill blanks in a Word document for multiple rows of data from an Excel file.
    Generate a separate Word file for each row in memory.

    Parameters:
    template_content (BytesIO): Content of the Word template file.
    df (DataFrame): DataFrame containing the Excel data.

    Returns:
    list: List of in-memory Word files with their names.
    """
    # Load the template document
    template_doc = Document(template_content)

    files = []

    for index, row in df.iterrows():
        # Create a new document by reloading the template for each row
        doc = Document(template_content)

        # Get row data as a dictionary
        row_data = row.to_dict()

        # Replace placeholders in the document with row data
        for paragraph in doc.paragraphs:
            for key, value in row_data.items():
                placeholder = f"<{key}>"  # Assuming placeholders are in the format <column_name>
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, str(value))

        # Save the document to an in-memory BytesIO object
        file_name = f"filled_letter_{index + 1}.docx"
        file_stream = io.BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)  # Reset the stream position for reading
        files.append((file_name, file_stream))

    return files


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
    # Read the Excel file
    df = pd.read_excel(excel_file)

    # Display the data for confirmation
    st.subheader("Excel Data Preview")
    st.write(df)

    if st.button("Generate Letters"):
        # Process and generate Word files
        files = fill_blanks_for_multiple_rows(word_file, df)

        # Display success message
        st.success("Letters generated successfully!")

        # Provide download links for each file
        for file_name, file_stream in files:
            st.download_button(
                label=f"Download {file_name}",
                data=file_stream,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
