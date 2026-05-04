import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
from io import BytesIO

# --- 1. CONFIGURATION ---
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])

@st.cache_resource
def find_working_model():
    try:
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                return m.name
    except:
        return "models/gemini-1.5-flash"
    return "models/gemini-1.5-flash"

selected_model_name = find_working_model()
model = genai.GenerativeModel(selected_model_name)

# --- 2. AI LOGIC (INTEGRATED CRITERIA) ---
def generate_advanced_plan(topic, syllabus, extra_context):
    prompt = f"""
    Topic: {topic}. Syllabus Code: {syllabus}. Context: {extra_context}.
    Generate a professional lesson plan in English. 
    Use the following EXACT markers for the document structure:

    SECTION: LESSON OBJECTIVES
    [4 points]
    SECTION: LESSON OUTCOMES
    [4 points]
    SECTION: SUCCESS CRITERIA
    [4 points]
    SECTION: PREREQUISITE
    [1 point]
    SECTION: KEYWORDS
    [6 items]
    SECTION: HOTS
    [4 main domains from Bloom's Taxonomy]
    SECTION: DIGITAL CITIZENSHIP
    [4 points on ethical tech use/Chromebooks/Canva/YouTube]

    SECTION: SUGGESTED OPENING (WAY FORWARD)
    [Hook activity and transition plan]

    SECTION: DIFFERENTIATION STRATEGIES
    - HA (Higher Achiever): [1 challenging activity]
    - MA (Medium Achiever): [1 core activity]
    - LA (Lower Achiever): [1 scaffolded activity]

    SECTION: BLENDED LEARNING (30 MINS)
    - Activity 1 & 2: [Descriptions]
    - Teacher Preparation: [Step-by-step before lesson]
    - Objectives: [3 points]
    - Student Tasks: [Step-by-step details]

    SECTION: PLENARY (EXIT TICKET)
    [2-3 minute closing activity]

    SECTION: HOMEWORK
    [Task assigned based on topic]
    """
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"System Error: {str(e)}"

# --- 3. WORD EXPORT LOGIC ---
def create_word_export(topic, syllabus, text):
    doc = Document()
    doc.add_heading(f'Advanced Lesson Plan: {topic}', 0)

    # Admin Header
    admin_table = doc.add_table(rows=3, cols=4)
    admin_table.style = 'Table Grid'
    labels = [["Week No:", "Date:"], ["Class Size:", "Day:"], ["Venue:", "Duration:"]]
    for r in range(3):
        admin_table.cell(r, 0).text = labels[r][0]
        admin_table.cell(r, 2).text = labels[r][1]
    doc.add_paragraph()

    # Parsing and Boxing ALL Sections
    sections = text.split('SECTION:')
    for section in sections:
        if not section.strip(): continue
        lines = section.strip().split('\n')
        title = lines[0].strip()
        content = "\n".join(lines[1:]).strip()
        
        doc.add_heading(title.title(), level=1)
        table = doc.add_table(rows=1, cols=1)
        table.style = 'Table Grid'
        table.cell(0, 0).text = content
        doc.add_paragraph()

    # HOD Approval
    doc.add_page_break()
    doc.add_heading("HOD Approval & Remarks", level=1)
    hod_table = doc.add_table(rows=2, cols=2)
    hod_table.style = 'Table Grid'
    hod_table.cell(0, 0).text = "Remarks:"
    hod_table.rows[1].height = Pt(50)
    hod_table.cell(1, 0).text = "Date:"; hod_table.cell(1, 1).text = "Signature:"

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 4. GUI ---
st.set_page_config(page_title="Advanced Lesson Planner", layout="wide")

st.title("🎓 Advanced Differentiated Lesson Planner")
st.info("This version includes Objectives, HOTS, Differentiation (HA/MA/LA), and Blended Learning.")

c1, c2 = st.columns(2)
with c1: u_topic = st.text_input("Lesson Topic:")
with c2: u_syllabus = st.text_input("Syllabus Code:")
u_extra = st.text_area("Extra Context (Optional):")

if st.button("🚀 GENERATE COMPLETE LESSON PLAN"):
    if u_topic and u_syllabus:
        with st.spinner("AI is integrating all criteria into your plan..."):
            result = generate_advanced_plan(u_topic, u_syllabus, u_extra)
            st.session_state['adv_plan_out'] = result
    else:
        st.warning("Please fill in the Topic and Syllabus.")

if 'adv_plan_out' in st.session_state:
    st.divider()
    st.subheader("AI Draft Preview")
    st.text_area("Content", st.session_state['adv_plan_out'], height=400)
    doc_file = create_word_export(u_topic, u_syllabus, st.session_state['adv_plan_out'])
    st.download_button("📥 Download Word (.docx)", doc_file, f"Advanced_Plan_{u_topic}.docx")

st.markdown("---")
st.caption("v3.5 | Integrated Professional Criteria | © 2026 PTES Innovation")
