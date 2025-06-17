import streamlit as st
import tempfile
from format_resume import create_resume_doc

def parse_experience(raw_text):
    # Split experience entries by double newlines (assumes each job separated by blank line)
    jobs_raw = [job.strip() for job in raw_text.strip().split('\n\n') if job.strip()]
    jobs = []
    for job_raw in jobs_raw:
        lines = [line.strip() for line in job_raw.split('\n') if line.strip()]
        if len(lines) < 3:
            continue  # Not enough info: company/date, title, bullets
        # First line: Company, City, State, Years (e.g. "XYZ Corp, New York, NY 2019 to Present")
        comp_line = lines[0]
        # Try to split company and years from right side
        # Assume years are last part after last tab or two spaces
        # This is a simple heuristic, can be improved later
        parts = comp_line.rsplit(' ', 3)  # last 3 parts might be city, state, years
        if len(parts) >= 4:
            company = parts[0]
            city = parts[1].rstrip(',')  # remove trailing comma
            state = parts[2]
            years = parts[3]
        else:
            # fallback
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
    # Each education entry separated by double newline
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

def parse_list(raw_text):
    # Split by lines or pipe | symbol and strip
    items = []
    if '|' in raw_text:
        items = [item.strip() for item in raw_text.split('|') if item.strip()]
    else:
        items = [line.strip() for line in raw_text.strip().split('\n') if line.strip()]
    return items

def main():
    st.title("Simplified Resume Formatter")

    name = st.text_input("Name")
    summary = st.text_area("Summary")

    certifications_raw = st.text_area("Certifications (separate by | or new lines)")
    skills_raw = st.text_area("Skills (separate by | or new lines)")
    experience_raw = st.text_area("Experience (paste plain text, separate jobs by blank lines)")
    education_raw = st.text_area("Education (paste plain text, separate entries by blank lines)")

    if st.button("Generate Formatted Resume"):
        certifications = parse_list(certifications_raw)
        skills = parse_list(skills_raw)
        experience = parse_experience(experience_raw)
        education = parse_education(education_raw)

        # Use a placeholder logo path or remove if you want
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

if __name__ == "__main__":
    main()
