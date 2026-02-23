import google.generativeai as genai
from groq import Groq
import pandas as pd
import json
import time
import re
import os
from dotenv import load_dotenv
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pypandoc

# Load environment variables from .env file
load_dotenv()

# 1. Configuration
AI_PROVIDER = os.getenv("AI_PROVIDER", "gemini").lower()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

# Batch Folder List
FOLDER_PATHS = [
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\CCNA Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\COBOL Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\CodeIgniter Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Cucumber Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Data Analyst Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Data Science Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Data Warehouse Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\DevSecOps Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Django Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Docker Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\ElasticSearch Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Flutter Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Golang Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\HR Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\iOS Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Java Collections Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Java Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Jenkins Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\JMeter Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Kafka Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Keras Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Kotlin Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Kubernetes Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Laravel Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Machine Learning Interview Questions",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Microservices Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\MuleSoft Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\MySQL Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Pega Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\PL SQL Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Power BI Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\PySpark Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Python Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\React Native Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\REST API Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\SAP Basis Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Scala Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\ServiceNow Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Snowflake Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Splunk Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\SQL Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Swift Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\Tableau Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\TensorFlow Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\TypeScript Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\UiPath Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\VMware Interview Questions Practice Test",
    r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\WebMethods Interview Questions Practice Test"
]

INPUT_FILE = "questions.txt"

# Initialize AI Client
model = None
groq_client = None

if AI_PROVIDER == "gemini":
    if not GEMINI_API_KEY:
        raise ValueError("GEMINI_API_KEY not found in .env file")
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel('gemini-2.5-flash')
    print("Using Gemini API for generation.")
elif AI_PROVIDER == "groq":
    if not GROQ_API_KEY:
        raise ValueError("GROQ_API_KEY not found in .env file")
    groq_client = Groq(api_key=GROQ_API_KEY)
    print("Using Groq API (Llama 3) for generation.")

columns = [
    "Question", "Question Type", 
    "Answer Option 1", "Explanation 1", 
    "Answer Option 2", "Explanation 2", 
    "Answer Option 3", "Explanation 3", 
    "Answer Option 4", "Explanation 4", 
    "Answer Option 5", "Explanation 5", 
    "Answer Option 6", "Explanation 6", 
    "Correct Answers", "Overall Explanation", "Domain"
]

def get_topic_name(folder_path):
    """Extract clean topic name from folder name."""
    folder_name = os.path.basename(folder_path.rstrip(os.sep))
    # Remove " Interview Questions Practice Test" if present or " Interview Questions"
    topic = folder_name.replace(" Interview Questions Practice Test", "").replace(" Interview Questions", "").strip()
    return topic

def load_questions(folder_path, file_name):
    questions = []
    file_path = os.path.join(folder_path, file_name)
    if not os.path.exists(file_path):
        print(f"Error: {file_path} not found.")
        return []

    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or re.match(r'^Section \d+.*', line, re.IGNORECASE):
                continue
            clean_q = re.sub(r'^(\d+[\.\)]|Question \d+:?)\s*', '', line, flags=re.IGNORECASE).strip()
            if clean_q:
                questions.append(clean_q)
    return questions

def get_ai_response(prompt):
    if AI_PROVIDER == "gemini":
        response = model.generate_content(prompt)
        return response.text.strip()
    elif AI_PROVIDER == "groq":
        chat_completion = groq_client.chat.completions.create(
            messages=[{"role": "user", "content": prompt}],
            model="llama-3.1-8b-instant",
            response_format={"type": "json_object"}
        )
        return chat_completion.choices[0].message.content.strip()

def generate_question_data(question, total_count, current_index, topic_name):
    print(f"Processing question {current_index}/{total_count}: {question[:50]}...")
    
    prompt = f"""
    You are an expert {topic_name} coach. 
    Analyze the following {topic_name} interview question and generate 6 multiple-choice options (1 correct, 5 tricky but incorrect distractors), explanations for each, the correct answer index, an overall explanation, and the domain.
    
    Question: "{question}"
    
    Respond ONLY with a valid JSON object matching this exact structure, with no markdown formatting or extra text.
    The "Correct Answers" must be a single digit (1, 2, 3, 4, 5, or 6) representing the index of the correct answer option.
    The "Question" in the JSON must NOT contain any numbering prefix.

    {{
      "Question": "The clean question text",
      "Question Type": "multiple-choice",
      "Answer Option 1": "Option 1 text",
      "Explanation 1": "Explanation why 1 is right/wrong",
      "Answer Option 2": "Option 2 text",
      "Explanation 2": "Explanation why 2 is right/wrong",
      "Answer Option 3": "Option 3 text",
      "Explanation 3": "Explanation why 3 is right/wrong",
      "Answer Option 4": "Option 4 text",
      "Explanation 4": "Explanation why 4 is right/wrong",
      "Answer Option 5": "Option 5 text",
      "Explanation 5": "Explanation why 5 is right/wrong",
      "Answer Option 6": "Option 6 text",
      "Explanation 6": "Explanation why 6 is right/wrong",
      "Correct Answers": "1",
      "Overall Explanation": "A comprehensive summary of the concept",
      "Domain": "{topic_name} topic area"
    }}
    """
    
    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            response_text = get_ai_response(prompt)
            if response_text.startswith("```json"):
                response_text = response_text[7:-3].strip()
            elif response_text.startswith("```"):
                response_text = response_text[3:-3].strip()
            return json.loads(response_text)
        except Exception as e:
            if attempt < max_retries:
                wait_time = 2 ** attempt
                print(f"Error on attempt {attempt}/{max_retries}: {e}. Retrying in {wait_time}s...")
                time.sleep(wait_time)
            else:
                print(f"Failed to process question after {max_retries} attempts: {e}")
                error_row = {col: "" for col in columns}
                error_row["Question"] = question
                error_row["Overall Explanation"] = f"ERROR GENERATING: {str(e)}"
                return error_row

def clean_all_text(text):
    if not isinstance(text, str): return text
    patterns = ["A. ", "B. ", "C. ", "D. ", "E. ", "F. ", "1. ", "2. ", "3. ", "4. ", "5. ", "6. ", "a. ", "b. ", "c. ", "d. ", "e. ", "f. "]
    for p in patterns: text = text.replace(p, "")
    return text

def merge_csv_files(folder_path, output_file):
    csv_files = [file for file in os.listdir(folder_path) if file.endswith(".csv") and file.startswith("Sheet")]
    if not csv_files: return None
    dfs = []
    for file in csv_files:
        try:
            df = pd.read_csv(os.path.join(folder_path, file))
            dfs.append(df)
        except Exception as e: print(f"Error reading {file}: {e}")
    if not dfs: return None
    merged_df = pd.concat(dfs, ignore_index=True)
    columns_to_keep = ["Question", "Question Type", "Answer Option 1", "Explanation 1", "Answer Option 2", "Explanation 2", "Answer Option 3", "Explanation 3", "Answer Option 4", "Explanation 4", "Answer Option 5", "Explanation 5", "Answer Option 6", "Explanation 6", "Correct Answers", "Overall Explanation", "Domain"]
    existing_columns = [c for c in columns_to_keep if c in merged_df.columns]
    merged_df = merged_df[existing_columns].dropna(subset=["Question", "Correct Answers", "Overall Explanation"])
    merged_df.to_csv(output_file, index=False, encoding='utf-8')
    return merged_df

def create_docx(df, title_name, output_file):
    doc = Document()
    title = doc.add_heading(title_name, 0); title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Exam Prep and Study Guide\n"); sub.alignment = WD_ALIGN_PARAGRAPH.CENTER; sub.runs[0].bold = True
    total_text = doc.add_paragraph(f"Total Questions: {len(df)}"); total_text.alignment = WD_ALIGN_PARAGRAPH.CENTER; total_text.runs[0].italic = True
    author = doc.add_paragraph("\nBy\nManish Dnyandeo Salunke"); author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()
    doc.add_heading("Preface", level=1)
    doc.add_paragraph("This book is a practical resource for technical interviews. It covers a wide range of MCQs to help reinforce your understanding.\n\n— Manish Dnyandeo Salunke")
    doc.add_page_break()
    doc.add_heading("About the Author", level=1)
    doc.add_paragraph("Manish Dnyandeo Salunke is a seasoned IT professional and author from Pune, India.")
    doc.add_page_break()
    KEYS = [("Answer Option 1", "Explanation 1"), ("Answer Option 2", "Explanation 2"), ("Answer Option 3", "Explanation 3"), ("Answer Option 4", "Explanation 4"), ("Answer Option 5", "Explanation 5"), ("Answer Option 6", "Explanation 6")]
    LABELS = ["A", "B", "C", "D", "E", "F"]
    for q_num, (_, row) in enumerate(df.iterrows(), start=1):
        doc.add_heading(f"Q{q_num}. {row['Question']}", level=1)
        if "Domain" in row: doc.add_paragraph(f"Domain: {row['Domain']}").runs[0].italic = True
        doc.add_paragraph()
        for i, (opt_col, exp_col) in enumerate(KEYS):
            if opt_col in row and not pd.isna(row[opt_col]):
                p = doc.add_paragraph(); p.add_run(f"{LABELS[i]}. ").bold = True; p.add_run(str(row[opt_col]))
        doc.add_paragraph()
        doc.add_paragraph().add_run("Correct Answer: ").bold = True; doc.paragraphs[-1].add_run(str(row["Correct Answers"]))
        if "Overall Explanation" in row: doc.add_paragraph().add_run("Explanation: ").bold = True; doc.paragraphs[-1].add_run(str(row["Overall Explanation"]))
        doc.add_paragraph(); doc.add_page_break()
    doc.add_heading("Copyright Disclaimer", level=1)
    doc.add_paragraph("© Manish Dnyandeo Salunke. All rights reserved.")
    doc.save(output_file)

def convert_docx_to_epub(docx_file, epub_file, title):
    try:
        pypandoc.convert_file(docx_file, 'epub', outputfile=epub_file, extra_args=[f"--metadata=title:{title}", "--metadata=author:Manish Dnyandeo Salunke", "--metadata=lang:en"])
    except Exception as e: print(f"EPUB error: {e}")

def process_single_folder(folder_path):
    topic_name = get_topic_name(folder_path)
    print(f"\n========================================")
    print(f"PROCESSING TOPIC: {topic_name}")
    print(f"FOLDER: {folder_path}")
    print(f"========================================\n")
    
    questions = load_questions(folder_path, INPUT_FILE)
    if not questions: return
    
    output_xlsx = f"{topic_name.replace(' ', '_')}_Generated.xlsx"
    output_path = os.path.join(folder_path, output_xlsx)
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    
    chunk_size = 80
    for i in range(0, len(questions), chunk_size):
        chunk = questions[i:i + chunk_size]
        chunk_index = (i // chunk_size) + 1
        print(f"\n--- Sheet {chunk_index} ({i+1} to {i+len(chunk)}) ---")
        chunk_rows = []
        for j, question in enumerate(chunk):
            chunk_rows.append(generate_question_data(question, len(questions), i + j + 1, topic_name))
            time.sleep(2 if AI_PROVIDER == "gemini" else 1)
            if (j + 1) % 5 == 0 and (j + 1) < len(chunk):
                print("--- Rate limit pause (3s) ---"); time.sleep(3)
        
        df = pd.DataFrame(chunk_rows, columns=columns)
        for col in df.columns: df[col] = df[col].apply(clean_all_text)
        df.to_excel(writer, sheet_name=f'Sheet{chunk_index}', index=False)
        df.to_csv(os.path.join(folder_path, f"Sheet{chunk_index}.csv"), index=False, encoding='utf-8')
    
    writer.close()
    merged_csv = os.path.join(folder_path, "Merged_Questions.csv")
    merged_df = merge_csv_files(folder_path, merged_csv)
    if merged_df is not None:
        title = os.path.basename(folder_path.rstrip(os.sep))
        docx_p = os.path.join(folder_path, "MCQ_Ebook.docx")
        epub_p = os.path.join(folder_path, "MCQ_Ebook.epub")
        create_docx(merged_df, title, docx_p)
        convert_docx_to_epub(docx_p, epub_p, title)
    print(f"\nFinished processing: {topic_name}")

def main():
    while FOLDER_PATHS:
        folder = FOLDER_PATHS[0]
        try:
            process_single_folder(folder)
        except Exception as e:
            print(f"CRITICAL ERROR in {folder}: {e}")
        
        # Remove folder from list once done
        FOLDER_PATHS.pop(0)
        
        if FOLDER_PATHS:
            print(f"\n--- Waiting 1 minute before next folder ({len(FOLDER_PATHS)} folders remaining) ---")
            time.sleep(60)

if __name__ == "__main__":
    main()
