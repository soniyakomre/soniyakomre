import streamlit as st
import pdfplumber
import google.generativeai as genai
import os
import requests
from dotenv import load_dotenv
from docx2pdf import convert  # Import docx2pdf for conversion
import pythoncom  # For COM initialization
from gtts import gTTS
from io import BytesIO
import base64
import pandas as pd

# Load environment variables from .env file
load_dotenv()

# Extract text from PDF using pdfplumber
def extract_text_from_pdf(file):
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text += page.extract_text()
    return text

# Chat with PDF content using Gemini model
def chat_with_pdf(pdf_text, user_query):
    prompt = f"Given the following text from a PDF:\n\n{pdf_text}\n\nAnswer this question: {user_query}"
    
    # Initialize the generative AI model (using Gemini)
    model = genai.GenerativeModel("gemini-1.5-flash")

    response = model.generate_content(prompt)
    return response.text

# Convert DOCX to PDF using docx2pdf for accurate conversion
def convert_docx_to_pdf(docx_file):
    """
    Converts a .docx file to a .pdf file using docx2pdf and returns the PDF as a byte stream.
    """
    # Initialize COM before calling docx2pdf
    pythoncom.CoInitialize()  # Initialize COM
    
    # Convert DOCX to PDF
    output_pdf_path = "converted_output.pdf"  # Define the path to save the PDF
    with open("temp.docx", "wb") as temp_docx:
        temp_docx.write(docx_file.read())  # Save the uploaded DOCX temporarily
    
    # Convert the temporary DOCX file to PDF
    convert("temp.docx", output_pdf_path)
    
    # Read the converted PDF file and return as a byte stream
    with open(output_pdf_path, "rb") as pdf_file:
        pdf_data = pdf_file.read()
    
    # Clean up temporary DOCX file
    os.remove("temp.docx")
    
    return pdf_data
 #upload the  csv file
def convert_csv_to_excel(csv_file):
    df = pd.read_csv(csv_file)
    excel_path = "converted_output.xlsx"
    df.to_excel(excel_path, index=False)
    return excel_path

#upload the excel file
def convert_excel_to_csv(excel_file):
    df = pd.read_excel(excel_file)
    csv_path = "converted_output.csv"
    df.to_csv(csv_path, index=False)
    return csv_path


def text_to_speech(text, lang='en'):
    tts = gTTS(text=text, lang=lang)
    fp = BytesIO()
    tts.write_to_fp(fp)
    return fp



# Main app logic
def main():

    # Streamlit sidebar for user interaction
    st.sidebar.title("PDF 📄 Chatbot 👾")
    
    # File upload for PDF
    df = st.sidebar.file_uploader("Upload a PDF📄 file", type="pdf")
    
    if df is not None:
        # Extract text from PDF
        pdf_text = extract_text_from_pdf(df)
        
        # Display the extracted text (first 1500 characters for preview)
        st.subheader("Extracted Text:")
        st.text_area("Extracted Text", pdf_text[:1500], height=200)

        # Ask for a query based on the extracted text
        user_query = st.text_input("Ask a question about the PDF content:")
        
        if user_query:
            response = chat_with_pdf(pdf_text, user_query)
            st.write(f"Chatbot's response: {response}")
    
    # File upload for DOCX
    docx_file = st.sidebar.file_uploader("Upload a DOCX📰 file", type="docx")
    
    if docx_file is not None:
        # Convert DOCX to PDF
        pdf_data = convert_docx_to_pdf(docx_file)
        
        # Create a download link for the converted PDF
        st.write("Converted PDF is ready to download:")
        st.download_button(
            label="Download PDF",
            data=pdf_data,
            file_name="converted_output.pdf",
            mime="application/pdf"
        )
    
    # Section for general generative AI queries outside the PDF context
    st.header("Chat With Me 🤖")
    st.subheader("Enter a General Query 💬")
    user_input = st.text_input(" ")
    
    if user_input:
        # Configure the generative AI model
        genai.configure(api_key="AIzaSyANyn3ncz--b8t3_xbATjBP0QLEDYUWNXA")  # Load API key
        
        model = genai.GenerativeModel("gemini-1.5-flash")
        
        # Generate response based on the input
        response = model.generate_content(user_input)
        st.write(response.text)  # Display the response from the generative model

    # CSV to Excel conversion
    csv_file = st.sidebar.file_uploader("Upload a CSV file", type="csv")
    if csv_file:
        if st.button("Convert CSV to Excel"):
            excel_path = convert_csv_to_excel(csv_file)
            with open(excel_path, "rb") as file:
                st.download_button(
                    label="Download Excel",
                    data=file,
                    file_name="converted_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )  

    # Excel to CSV conversion
    excel_file = st.sidebar.file_uploader("Upload an Excel file", type="xlsx")
    if excel_file:
        if st.button("Convert Excel to CSV"):
            csv_path = convert_excel_to_csv(excel_file)
            with open(csv_path, "rb") as file:
                st.download_button(
                    label="Download CSV",
                    data=file,
                    file_name="converted_output.csv",
                    mime="text/csv"
                )              

   #convert to speech
    if st.button("Convert to Speech  🔊"):
    # Convert text to speech
     lang_code = 'en'

     speech_fp = text_to_speech(response.text, lang=lang_code)
     # Play the speech
     st.audio(speech_fp, format='audio/mpeg', start_time=0)
     # Download link
     b64 = base64.b64encode(speech_fp.getvalue()).decode()
     href = f'<a href="data:audio/mpeg;base64,{b64}" download="speech.mp3">Download Audio</a>'
     st.markdown(href, unsafe_allow_html=True)        


if __name__ == '__main__':
    main()
