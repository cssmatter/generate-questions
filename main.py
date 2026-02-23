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

# Folder and File paths
SOURCE_FOLDER = r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\AEM Interview Questions Practice Test"
INPUT_FILE = "questions.txt"
OUTPUT_FILE = "AEM_Interview_Questions_Generated.xlsx"

# Initialize the selected model
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
else:
    raise ValueError(f"Unknown AI_PROVIDER: {AI_PROVIDER}")

# The columns matching your template
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

def load_questions(folder_path, file_name):
    questions = []
    file_path = os.path.join(folder_path, file_name)
    
    if not os.path.exists(file_path):
        print(f"Error: {file_path} not found.")
        return []

    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            # Skip empty lines, section headers like "Section 1", "Section 2" etc.
            if not line or re.match(r'^Section \d+.*', line, re.IGNORECASE):
                continue
            
            # Remove leading numbering like "1. ", "1) ", "Question 1: " etc.
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
            messages=[{
                "role": "user",
                "content": prompt,
            }],
            model="llama-3.3-70b-versatile", # High quality free tier model
            response_format={"type": "json_object"} # Groq supports strict JSON mode
        )
        return chat_completion.choices[0].message.content.strip()

def generate_question_data(question, total_count, current_index):
    print(f"Processing question {current_index}/{total_count}: {question[:50]}...")
    
    prompt = f"""
    You are an expert AEM (Adobe Experience Manager) architect. 
    Analyze the following AEM interview question and generate 6 multiple-choice options (1 correct, 5 tricky but incorrect distractors), explanations for each, the correct answer index, an overall explanation, and the domain.
    
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
      "Domain": "AEM topic area (e.g., OSGi, JCR, Sling, Architecture)"
    }}
    """
    
    try:
        response_text = get_ai_response(prompt)
        
        # Clean up the response in case the model ignored the JSON format instruction
        if response_text.startswith("```json"):
            response_text = response_text[7:-3].strip()
        elif response_text.startswith("```"):
            response_text = response_text[3:-3].strip()
            
        return json.loads(response_text)
    except Exception as e:
        print(f"Error processing question: {e}")
        error_row = {col: "" for col in columns}
        error_row["Question"] = question
        error_row["Overall Explanation"] = f"ERROR GENERATING: {str(e)}"
        return error_row

def clean_all_text(text):
    if not isinstance(text, str):
        return text
    
    # Patterns to remove, matching the VBA script exactly
    patterns = [
        "A. ", "B. ", "C. ", "D. ", "E. ", "F. ",
        "1. ", "2. ", "3. ", "4. ", "5. ", "6. ",
        "a. ", "b. ", "c. ", "d. ", "e. ", "f. "
    ]
    
    for p in patterns:
        text = text.replace(p, "")
    
    return text

# Step 1: Merge all CSV files
def merge_csv_files(folder_path, output_file):
    csv_files = [file for file in os.listdir(folder_path) if file.endswith(".csv") and file.startswith("Sheet")]
    if not csv_files:
        print("No CSV files found to merge.")
        return None
    
    dfs = []
    for file in csv_files:
        try:
            df = pd.read_csv(os.path.join(folder_path, file))
            dfs.append(df)
        except Exception as e:
            print(f"Error reading {file}: {e}")
    
    if not dfs:
        return None

    merged_df = pd.concat(dfs, ignore_index=True)
    
    # Load and clean data (keep specific columns and drop incomplete rows)
    columns_to_keep = [
        "Question", "Question Type",
        "Answer Option 1", "Explanation 1",
        "Answer Option 2", "Explanation 2",
        "Answer Option 3", "Explanation 3",
        "Answer Option 4", "Explanation 4",
        "Answer Option 5", "Explanation 5",
        "Answer Option 6", "Explanation 6",
        "Correct Answers", "Overall Explanation", "Domain"
    ]
    
    existing_columns = [c for c in columns_to_keep if c in merged_df.columns]
    merged_df = merged_df[existing_columns]
    merged_df = merged_df.dropna(subset=["Question", "Correct Answers", "Overall Explanation"])
    
    merged_df.to_csv(output_file, index=False, encoding='utf-8')
    print(f"Merged CSV saved as '{output_file}'. Total Questions: {len(merged_df)}")
    return merged_df

# Step 2: Create Word document
def create_docx(df, folder_name, output_file):
    doc = Document()
    total_questions = len(df)

    # Title Page
    title = doc.add_heading(folder_name, 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub_heading = doc.add_paragraph("Exam Prep and Study Guide\n")
    sub_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_heading.runs[0].bold = True

    total_text = doc.add_paragraph(f"Total Questions: {total_questions}")
    total_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
    total_text.runs[0].italic = True

    author = doc.add_paragraph("\nBy\nManish Dnyandeo Salunke")
    author.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_page_break()

    # Preface
    doc.add_heading("Preface", level=1)
    content = (
        "In today's fast-paced and highly competitive tech world, interviews have evolved beyond just technical know-how. "
        "They now require clarity, speed, and confidence in answering structured questions.\n\n"
        "This book is a practical resource for aspiring system engineers, students, and IT professionals preparing for technical interviews. "
        "It covers a wide range of multiple-choice questions (MCQs), complete with correct answers and concise explanations to help reinforce your understanding.\n\n"
        "Whether you're preparing for your first job, transitioning into a new role, or simply brushing up on your skills, "
        "this book is designed to serve as a quick and effective learning tool.\n\n"
        "Thank you for choosing this book as part of your preparation journey. I hope it helps you succeed and grow in your IT career.\n\n"
        "— Manish Dnyandeo Salunke"
    )
    doc.add_paragraph(content)
    doc.add_page_break()

    # About Author
    doc.add_heading("About the Author", level=1)
    bio = (
        "Manish Dnyandeo Salunke is a seasoned IT professional, educator, and passionate author from Pune, India. "
        "With years of hands-on experience in the IT industry, Manish has contributed to various roles involving system engineering, infrastructure management, and technical support.\n\n"
        "His passion for writing and mentoring led him to create practical learning resources aimed at helping aspiring IT professionals succeed in their careers.\n\n"
        "Outside his technical pursuits, Manish enjoys storytelling, content creation, and writing books that simplify complex concepts for everyone."
    )
    doc.add_paragraph(bio)
    doc.add_page_break()

    # Questions
    OPTION_KEYS = [
        ("Answer Option 1", "Explanation 1"),
        ("Answer Option 2", "Explanation 2"),
        ("Answer Option 3", "Explanation 3"),
        ("Answer Option 4", "Explanation 4"),
        ("Answer Option 5", "Explanation 5"),
        ("Answer Option 6", "Explanation 6"),
    ]
    LABELS = ["A", "B", "C", "D", "E", "F"]

    def cell_val(row, col):
        if col not in row.index:
            return None
        val = row[col]
        if pd.isna(val) or str(val).strip() == "":
            return None
        return str(val).strip()

    for q_num, (_, row) in enumerate(df.iterrows(), start=1):
        question = cell_val(row, "Question")
        correct_answers_raw = cell_val(row, "Correct Answers")
        overall_explanation = cell_val(row, "Overall Explanation")
        domain = cell_val(row, "Domain")

        doc.add_heading(f"Q{q_num}. {question}", level=1)

        if domain:
            meta_p = doc.add_paragraph(f"Domain: {domain}")
            meta_p.runs[0].italic = True

        doc.add_paragraph()

        options = []
        for i, (opt_col, exp_col) in enumerate(OPTION_KEYS):
            opt_text = cell_val(row, opt_col)
            if opt_text is None:
                continue
            exp_text = cell_val(row, exp_col)
            options.append((LABELS[i], opt_text, exp_text))

        correct_labels = set()
        if correct_answers_raw:
            for part in str(correct_answers_raw).replace(";", ",").split(","):
                part = part.strip().upper()
                if part.isdigit():
                    idx = int(part) - 1
                    if 0 <= idx < len(LABELS):
                        correct_labels.add(LABELS[idx])
                elif part in LABELS:
                    correct_labels.add(part)

        for label, opt_text, _ in options:
            p = doc.add_paragraph()
            p.add_run(f"{label}. ").bold = True
            p.add_run(opt_text)

        doc.add_paragraph()
        ca_p = doc.add_paragraph()
        ca_p.add_run("Correct Answer: ").bold = True
        ca_p.add_run(str(correct_answers_raw) or "")

        if overall_explanation:
            exp_p = doc.add_paragraph()
            exp_p.add_run("Explanation: ").bold = True
            exp_p.add_run(overall_explanation)

        has_per_option_exp = any(exp for _, _, exp in options if exp)
        if has_per_option_exp:
            doc.add_paragraph()
            hdr_p = doc.add_paragraph()
            hdr_p.add_run("Answer Analysis:").bold = True

            for label, opt_text, exp_text in options:
                if not exp_text:
                    continue
                is_correct = label in correct_labels
                analysis_p = doc.add_paragraph()
                status = "Correct" if is_correct else "Wrong"
                analysis_p.add_run(f"{label}. [{status}] ").bold = True
                analysis_p.add_run(f"{opt_text}: ").italic = True
                analysis_p.add_run(exp_text)

        doc.add_paragraph()
        doc.add_page_break()

    # Copyright
    doc.add_heading("Copyright Disclaimer", level=1)
    text = (
        "© 2026 Manish Dnyandeo Salunke. All rights reserved.\n\n"
        "No part of this book may be reproduced, stored, or transmitted in any form or by any means—electronic, mechanical, "
        "photocopying, recording, or otherwise—without the prior written permission of the author, "
        "except for brief quotations used in reviews or educational contexts.\n\n"
        "For permissions, please contact the author directly."
    )
    doc.add_paragraph(text)

    doc.save(output_file)
    print(f"Word file '{output_file}' created successfully.")

# Step 3: Convert to EPUB
def convert_docx_to_epub(docx_file, epub_file, folder_name):
    try:
        pypandoc.convert_file(
            docx_file, 'epub', outputfile=epub_file,
            extra_args=[
                f"--metadata=title:{folder_name}",
                f"--metadata=author:Manish Dnyandeo Salunke",
                f"--metadata=lang:en"
            ]
        )
        print(f"EPUB file '{epub_file}' created successfully.")
    except Exception as e:
        print(f"EPUB conversion error: {e}")

def main():
    questions = load_questions(SOURCE_FOLDER, INPUT_FILE)
    if not questions:
        print("No questions found to process.")
        return

    print(f"Starting generation for {len(questions)} questions using {AI_PROVIDER}...")
    
    output_path = os.path.join(SOURCE_FOLDER, OUTPUT_FILE)
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    
    # Keeping chunk_size = 1 for testing as requested by the user
    chunk_size = 80
    for i in range(0, len(questions), chunk_size):
        chunk = questions[i:i + chunk_size]
        chunk_index = (i // chunk_size) + 1
        print(f"\n--- Generating Sheet {chunk_index} (Questions {i+1} to {i+len(chunk)}) ---")
        
        chunk_rows = []
        for j, question in enumerate(chunk):
            question_data = generate_question_data(question, len(questions), i + j + 1)
            chunk_rows.append(question_data)
            
            # Rate limit protection
            sleep_time = 2 if AI_PROVIDER == "gemini" else 1
            time.sleep(sleep_time) 
            
            # 5-second pause after every 5 questions
            if (j + 1) % 5 == 0 and (j + 1) < len(chunk):
                print("--- Rate limit pause (5s) ---")
                time.sleep(5)
            
        df = pd.DataFrame(chunk_rows, columns=columns)
        
        for col in df.columns:
            df[col] = df[col].apply(clean_all_text)
            
        df.to_excel(writer, sheet_name=f'Sheet{chunk_index}', index=False)
        
        csv_filename = f"Sheet{chunk_index}.csv"
        csv_path = os.path.join(SOURCE_FOLDER, csv_filename)
        df.to_csv(csv_path, index=False, encoding='utf-8')
        print(f"Exported: {csv_path}")
    
    writer.close()
    print(f"\nExcel generation complete. Excel saved in: {output_path}")

    # Start Ebook Generation
    print("\n--- Starting MCQ Ebook Generation ---")
    merged_csv = os.path.join(SOURCE_FOLDER, "Merged_Questions.csv")
    merged_df = merge_csv_files(SOURCE_FOLDER, merged_csv)
    
    if merged_df is not None:
        folder_name = os.path.basename(SOURCE_FOLDER.rstrip(os.sep))
        docx_path = os.path.join(SOURCE_FOLDER, "MCQ_Ebook.docx")
        epub_path = os.path.join(SOURCE_FOLDER, "MCQ_Ebook.epub")
        
        create_docx(merged_df, folder_name, docx_path)
        convert_docx_to_epub(docx_path, epub_path, folder_name)
    
    print("\nAll tasks completed successfully!")

if __name__ == "__main__":
    main()
