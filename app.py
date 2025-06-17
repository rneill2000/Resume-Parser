import streamlit as st
import tempfile
import json
from format_resume import create_resume_doc

def parse_resume_file(uploaded_file):
    # For now, just return dummy parsed data
    return {
        "name": "Jane Doe",
        "summary": "Experienced Data Engineer with a passion for building scalable data pipelines.",
        "certifications": ["AWS Certified Solutions Architect", "Certified Scrum Master"],
        "skills": ["Python", "SQL", "AWS", "Docker"],
        "experience": [
            {
                "company": "XYZ Corp",
                "city": "New York",
                "state": "NY",
                "years": "2019 to Present",
                "title": "Data Engineer",
                "bullets": [
                    "Improved database performance",
                    "Wrote backend APIs",
                ]
            },
            {
                "company": "ABC Inc",
                "city": "Boston",
                "state": "MA",
                "years": "2017 to 2019",
                "title": "Junior Developer",
                "bullets": [
                    "Automated reporting processes",
                    "Built scalable data pipelines",
                ]
            },
        ],
        "education": [
            {
                "university": "State University",
                "degree": "Bachelor of Science in Computer Science"
            }
        ]
    }

def main():
    st.title("Resume Formatting Tool")

    uploaded_file = st.file_uploader("Upload your resume (Word or PDF)", type=["docx", "pdf"])

    if uploaded_file is not None:
        parsed_data = parse_resume_file(uploaded_file)

        st.header("Edit Resume Fields")

        name = st.text_input("Name", value=parsed_data["name"])
        summary = st.text_area("Summary", value=parsed_data["summary"])

        certifications = st.text_area("Certifications (separate by |)", value=" | ".join(parsed_data["certifications"]))
        certifications_list = [x.strip() for x in certifications.split("|")]

        skills = st.text_area("Skills (separate by |)", value=" | ".join(parsed_data["skills"]))
        skills_list = [x.strip() for x in skills.split("|")]

        experience_text = st.text_area("Experience (JSON format)", value=json.dumps(parsed_data["experience"], indent=2))
        education_text = st.text_area("Education (JSON format)", value=json.dumps(parsed_data["education"], indent=2))

        if st.button("Generate Formatted Resume"):
            try:
                experience_list = json.loads(experience_text)
                education_list = json.loads(education_text)
            except Exception as e:
                st.error(f"Error parsing JSON: {e}")
                return

            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
                doc = create_resume_doc(
                    name,
                    summary,
                    certifications_list,
                    skills_list,
                    experience_list,
                    education_list,
                    "fulllogo_transparent.png"  # Update to your logo path or handle upload
                )
                doc.save(tmp_file.name)
                tmp_file.seek(0)
                st.success("Resume generated!")
                st.download_button("Download Resume DOCX", tmp_file.read(), file_name="formatted_resume.docx")

if __name__ == "__main__":
    main()
