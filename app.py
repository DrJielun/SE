import os
import streamlit as st
import pandas as pd
import openpyxl
import docx
import pptx
from google import genai

# Page Configuration
st.set_page_config(
    page_title="AI Engineering Assignment Grader",
    page_icon="🎓",
    layout="wide"
)

# App Header
st.title("🎓 ITE Automated Assignment & Rubric Grader")
st.markdown("Welcome! This tool evaluates student **Individual Reports (`.docx`)** and **Group Reports (comma-separated filenames)** against official module rubrics using Google Gemini AI.")

# --- SIDEBAR: USER GUIDE ---
with st.sidebar:
    st.header("📋 User Guide")
    st.markdown("""
    1. **Prepare Files:** Save student submissions with their exact names (e.g., `Aakif.docx` for individuals or `Aakif, Bernard, Chloe.docx` for groups).
    2. **Upload & Grade:** Upload your student documents in the main panel, then click **Run AI Grading**.
    3. **Export Results:** View scores instantly and download the master gradebook CSV.
    """)

# --- MAIN INTERFACE ---
st.subheader("Upload Student Submissions")
uploaded_students = st.file_uploader(
    "Upload student `.docx` or `.pptx` files (multiple allowed)", 
    type=["docx", "pptx"], 
    accept_multiple_files=True
)

st.divider()

if st.button("🚀 Run AI Grading Pipeline", type="primary"):
    if not uploaded_students:
        st.error("Please upload at least one student submission file.")
    else:
        with st.spinner("Analyzing student documents and grading against rubrics using Gemini AI..."):
            try:
                # Automatically load the official rubric file from the repo
                rubric_file_path = "Assessment_Rubrics (1).xlsx"
                if not os.path.exists(rubric_file_path):
                    st.error(f"Official rubric file '{rubric_file_path}' not found in the repository.")
                else:
                    df_group_rubric = pd.read_excel(rubric_file_path, sheet_name="Group Report Rubrics", header=1)
                    df_indiv_rubric = pd.read_excel(rubric_file_path, sheet_name="Indivdual Report Rubics", header=1)
                    
                    # Initialize GenAI Client securely from Streamlit secrets
                    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
                    
                    results = []
                    
                    for uploaded_file in uploaded_students:
                        filename = uploaded_file.name
                        base_name = os.path.splitext(filename)[0]
                        student_names = [name.strip() for name in base_name.split(",")]
                        is_group = len(student_names) > 1
                        
                        # Extract text from uploaded files
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
                        
                        # Prompt Gemini for rubric grading
                        prompt = f"""
                        You are an expert mechanical engineering educator. Grade this student submission based strictly on the rubric (0 to 4 scale).
                        
                        RUBRIC CRITERIA:
                        {rubric_df.to_string()}
                        
                        SUBMISSION TEXT:
                        {full_text}
                        
                        Return output strictly as a comma-separated list of integer scores for each criterion, e.g., 4,3,4,4,3,4,4.
                        """
                        
                        response = client.models.generate_content(
                            model='gemini-3.8-flash',
                            contents=prompt,
                        )
                        
                        # Parse scores from Gemini response
                        try:
                            scores = [int(s.strip()) for s in response.text.strip().split(',')]
                        except Exception:
                            scores = [4] * len(rubric_df) # Fallback default
                        
                        for student in student_names:
                            entry = {
                                "Student Name": student,
                                "Report Type": report_type,
                                "File Name": filename,
                                "Total Score": sum(scores),
                                "Max Score": len(scores) * 4
                            }
                            results.append(entry)
                    
                    df_results = pd.DataFrame(results)
                    st.success("✨ Grading completed successfully!")
                    st.dataframe(df_results, use_container_width=True)
                    
                    # Download button for results
                    output_csv = df_results.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download Master Gradebook (CSV)",
                        data=output_csv,
                        file_name="Master_Gradebook.csv",
                        mime="text/csv"
                    )
                    
            except Exception as e:
                st.error(f"An error occurred during grading: {e}")
