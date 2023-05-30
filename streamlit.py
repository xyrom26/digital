import streamlit as st
import pandas as pd
import subprocess

def process_file(file):
    df = pd.read_excel(file)
    # Process the uploaded file here
    # Add your desired functionality
    return df

def run_other_python_file(name, email, data):
    # Save the data as a temporary file
    temp_file = "temp.xlsx"
    data.to_excel(temp_file, index=False)

    # Run your other Python file here
    # Pass the name, email, and temporary file path to the other file as arguments
    subprocess.call(["python", "DO_AUTOMATION.py", name, email, temp_file])



def main():
    st.title("Digital Ocean Automation")
    st.write("Upload your XLSX file and enter your name and email")

    # File Upload
    uploaded_file = st.file_uploader("Upload XLSX file", type="xlsx")

    # Name and Email Input
    name = st.text_input("Name")
    email = st.text_input("Email")

    if uploaded_file is not None:
        # Process the uploaded file
        data = process_file(uploaded_file)

        # Display success message
        st.success("File uploaded successfully!")

        # Display uploaded file
        st.write("Uploaded File:")
        st.dataframe(pd.read_excel(uploaded_file))

        # Display name and email
        st.write("Name:", name)
        st.write("Email:", email)

        # Run other Python file button
        if st.button("Start Automation"):
            run_other_python_file(name, email, data)

if __name__ == '__main__':
    main()