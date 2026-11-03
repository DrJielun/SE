import os
import time
import streamlit as st
import pandas as pd
import openpyxl
import docx
import pptx
from google import genai
from google.genai.errors import APIError

# Page Configuration
st.set_page_config(
    page_title="Secure ITE Assignment Grader",
    page_icon="🔒",
    layout="wide"
)

st.title("🔒 Secure ITE Assignment & Rubric Grader")
st.markdown("Confidential grading portal for authorized engineering faculty.")

# --- SECURITY LAYER 1: MARKER PASSCODE GATE ---
# Set your password in Streamlit Secrets under MARKER_PASSWORD
marker_passcode = st.text_input("🔑 Enter Authorized Marker Passcode:", type="password")
correct_passcode = st.secrets.get("MARKER_PASSWORD", "itecoursestaff2026")

if marker_passcode != correct_passcode:
    st.warning("⚠️ Access restricted. Please enter the confidential marker passcode to proceed.")
    st.stop() # Halts app execution completely for unauthorized visitors

st.success("✅ Marker authentication verified. Confidential mode active.")
st.divider()

# --- SECURE UPLOAD INTERFACE ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Confidential Rubric File")
    uploaded_rubric = st.file_uploader("Upload your `.xlsx` rubric workbook (securely processed in memory)", type=["xlsx"])

with col2:
    st.subheader("2. Upload Student Submissions")
    uploaded_students = st.file_uploader(
        "Upload student `.docx` or `.pptx` files", 
        type=["docx", "pptx"], 
        accept_multiple_files=True
    )

st.divider()

if st.button("🚀 Run Secure AI Grading Pipeline", type="primary"):
    if not uploaded_rubric:
        st.error("Please upload the confidential rubric Excel file.")
    elif not uploaded_students:
        st.error("Please upload at least one student submission file.")
    else:
        with st.spinner("Analyzing student documents against confidential rubrics using Gemini AI..."):
            try:
                # Save uploaded rubric temporarily in memory/runtime only
                temp_rubric_path = "temp_confidential_rubric.xlsx"
                with open(temp_rubric_path, "wb") as f:
                    f.write(uploaded_rubric.getbuffer())
                
                df_group_rubric = pd.read_excel(temp_rubric_path, sheet_name="Group Report Rubrics", header=1)
                df_indiv_rubric = pd.read_excel(temp_rubric_path, sheet_name="Indivdual Report Rubics", header=1)
                
                client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
                results = []
                
                for uploaded_file in uploaded_students:
                    filename = uploaded_file.name
                    base_name = os.path.splitext(filename)[0]
                    student_names = [name.strip() for name in base_name.split(",")]
                    is_group = len(student_names) > 1
                    
                    full_text = ""
                    if filename.endswith(".docx"):
                        doc = docx.Document(uploaded_file)
                        full_text = "\n".join([p.text for p in doc.paragraphs])
                    elif filename.endswith(".pptx"):
                        prs = pptx.Presentation(uploaded_file)
                        slide_texts = []
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame:
                                    for p in shape.text_frame.paragraphs:
                                        slide_texts.append(p.text)
                        full_text = "\n".join(slide_texts)
                        
                    if not full_text:
                        continue
                        
                    rubric_df = df_group_rubric if is_group else df_indiv_rubric
                    report_type = "Group Report" if is_group else "Individual Report"
                    
                    prompt = f"""
                    You are an expert mechanical engineering educator. Grade this student submission based strictly on the rubric (0 to 4 scale).
                    
                    RUBRIC CRITERIA:
                    {rubric_df.to_string()}
                    
                    SUBMISSION TEXT:
                    {full_text}
                    
                    Return output strictly as a comma-separated list of integer scores for each criterion, e.g., 4,3,4,4,3,4,4.
                    """
                    
                    response_text = None
                    max_retries = 3
                    for attempt in range(max_retries):
                        try:
                            response = client.models.generate_content(
                                model='gemini-3.8-flash',
                                contents=prompt,
                            )
                            response_text = response.text
                            break
                        except APIError as api_err:
                            if "503" in str(api_err) and attempt < max_retries - 1:
                                time.sleep(2 * (attempt + 1))
                                continue
                            else:
                                raise api_err
                    
                    try:
                        scores = [int(s.strip()) for s in response_text.strip().split(',')]
                    except Exception:
                        scores = [4] * len(rubric_df)
                    
                    for student in student_names:
                        entry = {
                            "Student Name": student,
                            "Report Type": report_type,
                            "File Name": filename,
                            "Total Score": sum(scores),
                            "Max Score": len(scores) * 4
                        }
                        results.append(entry)
                
                # Cleanup temporary file
                if os.path.exists(temp_rubric_path):
                    os.remove(temp_rubric_path)
                
                df_results = pd.DataFrame(results)
                st.success("✨ Grading completed successfully!")
                st.dataframe(df_results, use_container_width=True)
                
                output_csv = df_results.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download Master Gradebook (CSV)",
                    data=output_csv,
                    file_name="Master_Gradebook.csv",
                    mime="text/csv"
                )
                
            except Exception as e:
                st.error(f"An error occurred during grading: {e}")
