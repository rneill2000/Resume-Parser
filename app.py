import streamlit as st
import tempfile
from format_resume import create_resume_doc

st.title("Branded Resume Editor")

st.markdown(
    """
    <style>
        .main {background-color: #f3f7fa;}
        .st-bb {color: #284b62;}
        .st-bb h1 {color: #0b233b;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.image("fulllogo_transparent.png", width=180)

st.header("Edit Your Resume Sections")

name = st.text_input("Name", value="")
summary = st.text_area("Summary", value="")
certifications_raw = st.text_area("Certifications (separate by | or new lines)", value="")
skills_raw = st.text_area("Skills (separate by | or new lines)", value="")
experience_raw = st.text_area("Experience (paste plain text, separate jobs by blank lines)", value="")
education_raw = st.text_area("Education (paste plain text, separate entries by blank lines)", value="")

def parse_list(raw_text):
    if '|' in raw_text:
        items = [item.strip() for item in raw_text.split('|') if item.strip()]
    else:
        items = [line.strip() for line in raw_text.strip().split('\n') if line.strip()]
    return items

def parse_experience(raw_text):
    jobs_raw = [job.strip() for job in raw_text.strip().split('\n\n') if job.strip()]
    jobs = []
    for job_raw in jobs_raw:
        lines = [line.strip() for line in job_raw.split('\n') if line.strip()]
        if len(lines) < 3:
            continue
        comp_line = lines[0]
        parts = comp_line.rsplit(' ', 3)
        if len(parts) >= 4:
            company = parts[0]
            city = parts[1].rstrip(',')
            state = parts[2]
            years = parts[3]
        else:
            company = comp_line
            city = state = years = ''
        title = lines[1]
        bullets = lines[2:]
        jobs.append({
            'company': company,
            'city': city,
            'state': state,
            'years': years,
            'title': title,
            'bullets': bullets
        })
    return jobs

def parse_education(raw_text):
    edus_raw = [edu.strip() for edu in raw_text.strip().split('\n\n') if edu.strip()]
    edus = []
    for edu_raw in edus_raw:
        lines = [line.strip() for line in edu_raw.split('\n') if line.strip()]
        if len(lines) < 2:
            continue
        univ = lines[0]
        degree = lines[1]
        edus.append({'university': univ, 'degree': degree})
    return edus

if st.button("Generate Formatted Resume"):
    certifications = parse_list(certifications_raw)
    skills = parse_list(skills_raw)
    experience = parse_experience(experience_raw)
    education = parse_education(education_raw)

    logo_path = "fulllogo_transparent.png"

    doc = create_resume_doc(
        name=name,
        summary=summary,
        certifications=certifications,
        skills=skills,
        experience=experience,
        education=education,
        logo_path=logo_path
    )

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
        doc.save(tmp_file.name)
        tmp_file.seek(0)
        st.success("Resume generated!")
        st.download_button("Download Resume DOCX", tmp_file.read(), file_name="formatted_resume.docx")
