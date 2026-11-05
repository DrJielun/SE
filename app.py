import os
import time
import json
import io
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
st.markdown("Confidential grading portal matching your exact rubric template layout.")

# --- SAFE SECRET LOADER (Prevents Render / Streamlit crashes) ---
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
            
            # Read sheets with header=1 to capture the rubric table structure
            df_group_rubric = pd.read_excel(temp_rubric_path, sheet_name="Group Report Rubrics", header=1)
            df_indiv_rubric = pd.read_excel(temp_rubric_path, sheet_name="Indivdual Report Rubics", header=1)
            
            api_key = get_secret("GEMINI_API_KEY")
            if not api_key:
                st.error("Gemini API key not found in environment or secrets.")
                st.stop()
                
            client = genai.Client(api_key=api_key)
            
            # Build openpyxl workbook directly to match exact template layout
            wb = openpyxl.Workbook()
            wb.remove(wb.active) # Remove default blank sheet
            
            indiv_files = []
            group_files = []
            
            for f in uploaded_students:
                base_name = os.path.splitext(f.name)[0]
                student_names = [name.strip() for name in base_name.split(",")]
                if len(student_names) > 1:
                    group_files.append((f, student_names))
                else:
                    indiv_files.append((f, student_names[0]))
            
            def process_grading_sheet(sheet_title, rubric_df, file_student_pairs):
                if not file_student_pairs:
                    return
                
                ws = wb.create_sheet(title=sheet_title)
                
                # Write title row
                ws.cell(row=1, column=1, value=f"Assessment Rubrics - {sheet_title}")
                
                # Write header row (row 2)
                headers = list(rubric_df.columns)
                for col_idx, h in enumerate(headers, 1):
                    ws.cell(row=2, column=col_idx, value=h)
                
                # Append student score and reason columns side-by-side
                current_col = len(headers) + 1
                student_mapping = {}
                
                for uploaded_file, student_name in file_student_pairs:
                    ws.cell(row=2, column=current_col, value=student_name)
                    ws.cell(row=2, column=current_col + 1, value="Reason")
                    student_mapping[student_name] = (current_col, current_col + 1)
                    current_col += 2
                
                # Write rubric criteria rows (starting at row 3)
                for r_idx, row in rubric_df.iterrows():
                    excel_row = r_idx + 3
                    for c_idx, val in enumerate(row, 1):
                        ws.cell(row=excel_row, column=c_idx, value=val if pd.notna(val) else "")
                
                # Grade each file via Gemini and fill in scores & reasons
                for uploaded_file, student_name in file_student_pairs:
                    status_container.write(f"📄 Processing {sheet_title}: **{uploaded_file.name}**...")
                    
                    full_text = ""
                    filename = uploaded_file.name
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
                        continue
                        
                    criterion_col = rubric_df.columns[0]
                    criterion_names = [str(val).split('\n')[0].strip() for val in rubric_df[criterion_col].dropna().tolist()]
                    
                    prompt = f"""
                    You are an expert mechanical engineering educator. Grade this student submission based strictly on the rubric criteria below (0 to 4 scale for each criterion).
                    
                    RUBRIC CRITERIA:
                    {rubric_df.to_string()}
                    
                    SUBMISSION TEXT:
                    {full_text}
                    
                    Return output strictly as a valid JSON object where keys are the criterion index ("0", "1", "2"...) and values are an object with two fields: "score" (integer 0-4) and "reason" (a brief 1-sentence explanation for the score).
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
                    
                    try:
                        clean_json = response_text.strip()
                        if clean_json.startswith("```json"):
                            clean_json = clean_json[7:-3].strip()
                        elif clean_json.startswith("```"):
                            clean_json = clean_json[3:-3].strip()
                        grading_data = json.loads(clean_json)
                    except Exception:
                        grading_data = {}
                    
                    score_col, reason_col = student_mapping[student_name]
                    for idx, _ in enumerate(criterion_names):
                        excel_row = idx + 3
                        item = grading_data.get(str(idx), {"score": 4, "reason": "Standard evaluation."})
                        score = int(item.get("score", 4))
                        reason = item.get("reason", "No reason provided.")
                        
                        ws.cell(row=excel_row, column=score_col, value=score)
                        ws.cell(row=excel_row, column=reason_col, value=reason)

            process_grading_sheet("Individual Report Rubrics", df_indiv_rubric, indiv_files)
            process_grading_sheet("Group Report Rubrics", df_group_rubric, group_files)
            
            if not wb.sheetnames:
                ws = wb.create_sheet(title="Master Gradebook")
                ws.cell(row=1, column=1, value="No valid student files processed.")

            if os.path.exists(temp_rubric_path):
                os.remove(temp_rubric_path)
            
            status_container.update(label="✨ Grading completed successfully!", state="complete", expanded=False)
            
            output_excel = io.BytesIO()
            wb.save(output_excel)
            excel_bytes = output_excel.getvalue()
            
            st.success("Excel Master Gradebook generated successfully matching your template layout!")
            st.download_button(
                label="📥 Download Master Gradebook (.xlsx)",
                data=excel_bytes,
                file_name="Master_Gradebook_Template_Layout.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            status_container.update(label="❌ Error occurred during grading.", state="error", expanded=True)
            st.error(f"Details: {e}")
