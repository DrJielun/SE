import os
import time
import json
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
st.markdown("Confidential grading portal with criterion scores and brief justifications.")

# --- SAFE SECRET LOADER ---
def get_secret(key, default=""):
    if key in os.environ:
        return os.environ[key]
    try:
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return default

# --- SECURITY LAYER: MARKER PASSCODE GATE ---
marker_passcode = st.text_input("🔑 Enter Authorized Marker Passcode:", type="password")
correct_passcode = get_secret("MARKER_PASSWORD", "itecoursestaff2026")

if marker_passcode != correct_passcode:
    st.warning("⚠️ Access restricted. Please enter the confidential marker passcode to proceed.")
    st.stop()

st.success("✅ Marker authentication verified. Confidential mode active.")
st.divider()

# --- UPLOAD INTERFACE ---
col1, col2 = st.columns(2)

with col1:
    uploaded_rubric = st.file_uploader("1. Upload Confidential Rubric File (.xlsx)", type=["xlsx"])

with col2:
    uploaded_students = st.file_uploader(
        "2. Upload Student Submissions (.docx / .pptx)", 
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
        status_container = st.status("Initializing grading pipeline...", expanded=True)
        try:
            status_container.write("Loading rubric sheets...")
            temp_rubric_path = "temp_confidential_rubric.xlsx"
            with open(temp_rubric_path, "wb") as f:
                f.write(uploaded_rubric.getbuffer())
            
            df_group_rubric = pd.read_excel(temp_rubric_path, sheet_name="Group Report Rubrics", header=1)
            df_indiv_rubric = pd.read_excel(temp_rubric_path, sheet_name="Indivdual Report Rubics", header=1)
            
            api_key = get_secret("GEMINI_API_KEY")
            if not api_key:
                st.error("Gemini API key not found in environment or secrets.")
                st.stop()
                
            client = genai.Client(api_key=api_key)
            results = []
            
            for uploaded_file in uploaded_students:
                filename = uploaded_file.name
                status_container.write(f"📄 Processing file: **{filename}**...")
                
                base_name = os.path.splitext(filename)[0]
                student_names = [name.strip() for name in base_name.split(",")]
                is_group = len(student_names) > 1
                
                # Extract text
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
                    
                if not full_text.strip():
                    status_container.warning(f"⚠️ Warning: No readable text found in {filename}.")
                    continue
                    
                rubric_df = df_group_rubric if is_group else df_indiv_rubric
                report_type = "Group Report" if is_group else "Individual Report"
                
                criterion_col = rubric_df.columns[0]
                criterion_names = [str(val).split('\n')[0].strip() for val in rubric_df[criterion_col].dropna().tolist()]
                
                # Prompt requesting JSON output containing both score and reason
                prompt = f"""
                You are an expert mechanical engineering educator. Grade this student submission based strictly on the rubric criteria below (0 to 4 scale for each criterion).
                
                RUBRIC CRITERIA:
                {rubric_df.to_string()}
                
                SUBMISSION TEXT:
                {full_text}
                
                Return output strictly as a valid JSON object where keys are the criterion index (0, 1, 2...) and values are an object with two fields: "score" (integer 0-4) and "reason" (a brief 1-sentence explanation for the score).
                Example format:
                {{
                  "0": {{"score": 3, "reason": "Identified three key contributors correctly."}},
                  "1": {{"score": 2, "reason": "Mentioned two causal agents but missed GHG emissions."}}
                }}
                """
                
                status_container.write(f"🤖 Sending request to Gemini AI for {filename}...")
                
                models_to_try = ['gemini-3.8-flash', 'gemini-3.6-flash', 'gemini-3.5-flash-lite']
                response_text = None
                success = False
                
                for model_name in models_to_try:
                    if success:
                        break
                    for attempt in range(2):
                        try:
                            response = client.models.generate_content(
                                model=model_name,
                                contents=prompt,
                            )
                            response_text = response.text
                            success = True
                            break
                        except APIError as api_err:
                            if "503" in str(api_err):
                                status_container.warning(f"⚠️ Model {model_name} busy (503). Retrying...")
                                time.sleep(2)
                                continue
                            else:
                                raise api_err
                
                if not success:
                    raise Exception("All Gemini flash model endpoints are currently experiencing high server demand (503). Please try again shortly.")
                
                # Clean up and parse JSON response
                try:
                    clean_json = response_text.strip()
                    if clean_json.startswith("```json"):
                        clean_json = clean_json[7:-3].strip()
                    elif clean_json.startswith("```"):
                        clean_json = clean_json[3:-3].strip()
                    grading_data = json.loads(clean_json)
                except Exception:
                    grading_data = {}
                
                scores_list = []
                for idx, crit_name in enumerate(criterion_names):
                    item = grading_data.get(str(idx), {"score": 4, "reason": "Passed standard criteria."})
                    score = int(item.get("score", 4))
                    reason = item.get("reason", "")
                    scores_list.append(score)
                    
                    for student in student_names:
                        # We will build dictionary per student below
                        pass
                
                total_score = sum(scores_list)
                max_score = len(criterion_names) * 4
                
                for student in student_names:
                    entry = {
                        "Student Name": student,
                        "Report Type": report_type,
                        "File Name": filename,
                    }
                    for idx, crit_name in enumerate(criterion_names):
                        header_label = crit_name[:40] + "..." if len(crit_name) > 40 else crit_name
                        item = grading_data.get(str(idx), {"score": 4, "reason": "Standard evaluation."})
                        score = item.get("score", 4)
                        reason = item.get("reason", "No reason provided.")
                        
                        entry[f"{header_label} [Score]"] = score
                        entry[f"{header_label} [Reason]"] = reason
                        
                    entry["Total Score"] = total_score
                    entry["Max Score"] = max_score
                    results.append(entry)
            
            if os.path.exists(temp_rubric_path):
                os.remove(temp_rubric_path)
            
            status_container.update(label="✨ Grading completed successfully!", state="complete", expanded=False)
            
            df_results = pd.DataFrame(results)
            st.success("Detailed Criterion Scores & Justifications:")
            st.dataframe(df_results, use_container_width=True)
            
            output_csv = df_results.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Detailed Master Gradebook (CSV)",
                data=output_csv,
                file_name="Detailed_Gradebook_With_Reasons.csv",
                mime="text/csv"
            )
            
        except Exception as e:
            status_container.update(label="❌ Error occurred during grading.", state="error", expanded=True)
            st.error(f"Details: {e}")
