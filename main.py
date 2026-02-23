import google.generativeai as genai
import pandas as pd
import json
import time
import re

# 1. Set up your API Key here
# Get one for free at https://aistudio.google.com/
API_KEY = "AIzaSyC5u_KwikvEk63pEygVOkCprjFHoXoMa-o"
genai.configure(api_key=API_KEY)

# Use the Gemini model
model = genai.GenerativeModel('gemini-2.5-flash')

# 2. Paste your list of questions here (I've added the first 3 as an example)
questions_list = [
    "1. Which architectural layer in AEM is responsible for providing a hierarchical data store and implementing the JSR-283 specification?",
    "2. In the context of AEM's technology stack, what is the primary role of Apache Felix?",
    "3. Which design pattern does the Apache Sling framework primarily use to map HTTP request URLs to repository resources?"
    # ... Add the rest of your 80 questions here
]

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

all_rows = []

print(f"Starting generation for {len(questions_list)} questions...")
for i, question in enumerate(questions_list):
    # Remove leading numbering like "1. ", "1) ", "Question 1: " etc.
    clean_question = re.sub(r'^(\d+[\.\)]|Question \d+:?)\s*', '', question, flags=re.IGNORECASE).strip()
    
    print(f"Processing question {i+1}/{len(questions_list)}: {clean_question[:50]}...")
    
    # Create a strict prompt so the AI returns exactly the JSON we need
    prompt = f"""
    You are an expert AEM (Adobe Experience Manager) architect. 
    Analyze the following AEM interview question and generate 6 multiple-choice options (1 correct, 5 tricky but incorrect distractors), explanations for each, the correct answer index, an overall explanation, and the domain.
    
    Question: "{clean_question}"
    
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
        # Call the API
        response = model.generate_content(prompt)
        response_text = response.text.strip()
        
        # Clean up the response in case the model included markdown code blocks
        if response_text.startswith("```json"):
            response_text = response_text[7:-3].strip()
            
        # Parse the JSON and add it to our list
        question_data = json.loads(response_text)
        all_rows.append(question_data)
        
        # Sleep briefly to avoid hitting free-tier API rate limits
        time.sleep(2) 
        
    except Exception as e:
        print(f"Error processing question {i+1}: {e}")
        # Append a blank/error row so the script doesn't completely fail
        error_row = {col: "" for col in columns}
        error_row["Question"] = question
        error_row["Overall Explanation"] = f"ERROR GENERATING: {str(e)}"
        all_rows.append(error_row)

# 3. Create a Pandas DataFrame and save to Excel
df = pd.DataFrame(all_rows, columns=columns)
output_filename = "AEM_Interview_Questions_Generated.xlsx"
df.to_excel(output_filename, index=False)

print(f"\nSuccess! Your Excel file has been saved as {output_filename}")