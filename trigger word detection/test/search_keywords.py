import os
import pandas as pd
from docx import Document
import time
import re
import inflect

# Load stopwords from a text file and prepare a regex pattern
def load_stopwords(file_path):
    p = inflect.engine()
    stopwords = []
    with open(file_path, 'r') as file:
        for line in file.readlines():
            word = line.strip().lower()
            stopwords.append(word)
            plural_form = p.plural(word)
            if plural_form:
                stopwords.append(plural_form)
    # Create a regex pattern that matches these words at word boundaries
    return re.compile(r'\b(' + '|'.join(re.escape(word) for word in stopwords) + r')', re.IGNORECASE)

# Extract context around a word in text
def extract_context(full_text, match_start, window=50):
    words = full_text.split()
    word_index = len(full_text[:match_start].split())
    start = max(0, word_index - window)
    end = min(len(words), word_index + window + 1)
    return ' '.join(words[start:end])

# Process DOCX files and collect results in a dataframe
def process_docx(file_path, stopwords_regex):
    results = []
    doc = Document(file_path)
    full_text = " ".join([p.text for p in doc.paragraphs])  # Merge all text
    text_found = False  # Flag to check if any word is found
    matches = re.finditer(stopwords_regex, full_text)
    for match in matches:
        text_found = True
        word = match.group()
        start_index = match.start()  # Character index in the full text
        context = extract_context(full_text, start_index)
        results.append({
            "file_name": os.path.basename(file_path),
            "word": word,
            "loc": start_index,  # Use the character position
            "context": context
        })
    # Check if no word was found
    if not text_found:
        results.append({
            "file_name": os.path.basename(file_path),
            "word": "NO HITS",
            "loc": -1,
            "context": "NO HITS FOUND in file"
        })
    # Add an empty row after processing each document
    results.append({ "file_name": "", "word": "", "loc": "", "context": "" })
    return results

# Main function to load stopwords, process each DOCX file in a folder, and save results to CSV
def main(stopwords_file, folder_path, output_file):
    start_time = time.time()
    stopwords_regex = load_stopwords(stopwords_file)
    all_results = []
    file_count = 0

    for filename in os.listdir(folder_path):
        if filename.endswith('.docx'):
            file_path = os.path.join(folder_path, filename)
            results = process_docx(file_path, stopwords_regex)
            all_results.extend(results)
            file_count += 1

    df = pd.DataFrame(all_results)
    df.to_csv(output_file, index=False)
    elapsed_time = time.time() - start_time

    print(f"Processed {file_count} files in {elapsed_time:.2f} seconds")

if __name__ == "__main__":
#    main('stopwords.txt', '/Users/elliotli/Library/CloudStorage/Box-Box/1_WPDOCS/1 - Exams & Grades Team/Exams/Exam Administrator/Original Exams/24A Exams/Final Exams/Take-home Exams/Faculty Exam Submissions', 'keywordhits_TK.csv')
#    main('stopwords.txt', '/Users/elliotli/Library/CloudStorage/Box-Box/1_WPDOCS/1 - Exams & Grades Team/Exams/Exam Administrator/Original Exams/24A Exams/Final Exams/In-Class Exams/Faculty Exam Submissions', 'keywordhits_IC.csv')