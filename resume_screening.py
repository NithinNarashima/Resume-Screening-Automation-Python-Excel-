import os
import docx2txt
from pdfminer.high_level import extract_text
import pandas as pd

# Extract text from docx or pdf file
def extract_text_from_file(file_path):
    if file_path.endswith(".docx"):
        return docx2txt.process(file_path)
    elif file_path.endswith(".pdf"):
        return extract_text(file_path)
    else:
        return ""

# Load keywords for each role from Excel
def load_keywords(file_path):
    df = pd.read_excel(file_path)
    keyword_dict = {}
    for _, row in df.iterrows():
        role = row['Role'].strip().lower()
        keywords = [kw.strip().lower() for kw in row['Keywords'].split(',')]
        keyword_dict[role] = keywords
    return keyword_dict

# Score a resume based on keyword matches
def score_resume(resume_text, keywords):
    resume_text = resume_text.lower()
    return sum(1 for kw in keywords if kw in resume_text)

# Go through each resume and score it for every role
def screen_resumes(resume_folder, keyword_dict):
    if not os.path.exists(resume_folder):
        print(f"❌ Folder not found: {resume_folder}")
        return pd.DataFrame()

    results = []
    for file in os.listdir(resume_folder):
        if not (file.endswith(".docx") or file.endswith(".pdf")):
            continue  # Skip unsupported files

        file_path = os.path.join(resume_folder, file)
        print(f"🔍 Reading: {file_path}")
        resume_text = extract_text_from_file(file_path)

        for role, keywords in keyword_dict.items():
            score = score_resume(resume_text, keywords)
            print(f"📊 {file} matched {score} keywords for role '{role}'")
            results.append({
                "File Name": file,
                "Role": role.title(),
                "Score": score
            })

    return pd.DataFrame(results)

# Save the results to Excel
def save_results(df, output_file="screening_results.xlsx"):
    if df.empty:
        print("⚠️ No results to save. Check if your 'resumes' folder has valid files.")
        return
    df_sorted = df.sort_values(by=["Role", "Score"], ascending=[True, False])
    df_sorted.to_excel(output_file, index=False)
    print(f"✅ Results saved to {output_file}")

# Main execution
if __name__ == "__main__":
    keyword_file = "job_keywords.xlsx"  # Must be in the same folder as this script
    resume_folder = "resumes"           # Subfolder inside current directory

    keyword_dict = load_keywords(keyword_file)
    df_results = screen_resumes(resume_folder, keyword_dict)
    save_results(df_results)
