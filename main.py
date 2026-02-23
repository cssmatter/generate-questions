import google.generativeai as genai
import pandas as pd
import json
import time
import re
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# 1. Configuration
API_KEY = os.getenv("GEMINI_API_KEY")
if not API_KEY:
    raise ValueError("GEMINI_API_KEY not found in .env file")
genai.configure(api_key=API_KEY)

# Folder and File paths
SOURCE_FOLDER = r"C:\Users\manis\Udemy\certifications\Interview Practice Tests\Updated\PENDING\AEM Interview Questions Practice Test"
INPUT_FILE = "questions.txt"
OUTPUT_FILE = "AEM_Interview_Questions_Generated.xlsx"

# Use the Gemini model
model = genai.GenerativeModel('gemini-2.5-flash')

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
        response = model.generate_content(prompt)
        response_text = response.text.strip()
        
        if response_text.startswith("```json"):
            response_text = response_text[7:-3].strip()
            
        return json.loads(response_text)
    except Exception as e:
        print(f"Error processing question: {e}")
        error_row = {col: "" for col in columns}
        error_row["Question"] = question
        error_row["Overall Explanation"] = f"ERROR GENERATING: {str(e)}"
        return error_row

def main():
    questions = load_questions(SOURCE_FOLDER, INPUT_FILE)
    if not questions:
        print("No questions found to process.")
        return

    print(f"Starting generation for {len(questions)} questions...")
    
    output_path = os.path.join(SOURCE_FOLDER, OUTPUT_FILE)
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    
    chunk_size = 80
    for i in range(0, len(questions), chunk_size):
        chunk = questions[i:i + chunk_size]
        chunk_index = (i // chunk_size) + 1
        print(f"\n--- Generating Sheet {chunk_index} (Questions {i+1} to {i+len(chunk)}) ---")
        
        chunk_rows = []
        for j, question in enumerate(chunk):
            question_data = generate_question_data(question, len(questions), i + j + 1)
            chunk_rows.append(question_data)
            time.sleep(2) # API rate limit protection
            
        df = pd.DataFrame(chunk_rows, columns=columns)
        df.to_excel(writer, sheet_name=f'Sheet{chunk_index}', index=False)
    
    writer.close()
    print(f"\nSuccess! Your Excel file has been saved in: {output_path}")

if __name__ == "__main__":
    main()
