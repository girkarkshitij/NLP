from pdfminer.high_level import extract_text
 
import spacy
import docx2txt
import nltk
import re
import os
import csv

from spacy.matcher import Matcher
from nltk.corpus import stopwords

nlp = spacy.load('en_core_web_sm')
matcher = Matcher(nlp.vocab)
 
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('maxent_ne_chunker')
nltk.download('words')
nltk.download('stopwords')

PHONE_REG = re.compile(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]')
EMAIL_REG = re.compile(r'[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+')
YEAR_REG = re.compile(r'\b[21][09][0-9][0-9]')
DIRECTORY = r"C:\Users\Kshitij\Desktop\NlpMiniProject\resume"
INFO_ROWS = []
index = 0

# Expected skill set
SKILLS_DB = [
    'MERN',
    'full stack',
    'react',
    'python',
    'java',
    'javascript',
    'SQL',
    'mongodb',
    'node',
    'express',
]

EDUCATION = [
            'BE','B.E.', 'B.E', 'BS', 'B.S', 
            'ME', 'M.E', 'M.E.', 'MS', 'M.S', 
            'BTECH', 'B.TECH', 'M.TECH', 'MTECH', 
            'SSC', 'HSC', 'CBSE', 'ICSE', 'X', 'XII'
        ]

def extract_text_from_pdf(pdf_path):
    return extract_text(pdf_path)

def extract_text_from_docx(docx_path):
    txt = docx2txt.process(docx_path)
    if txt:
        return txt.replace('\t', ' ')
    return None

def extract_year_of_graduation(text):
    year = re.findall(YEAR_REG, text)
    year.sort(key = int)
    return year[len(year)-1]
 
def extract_phone_number(resume_text):
    phone = re.findall(PHONE_REG, resume_text)
 
    if phone:
        number = ''.join(phone[0])
 
        if resume_text.find(number) >= 0 and len(number) <= 16:
            return number
    return None

def extract_names(text):
    nlp_text = nlp(text)
    
    # First name and Last name are Proper Nouns
    pattern = [{'POS': 'PROPN'}, {'POS': 'PROPN'}]
    
    matcher.add('NAME', [pattern])
    
    matches = matcher(nlp_text)
    
    for match_id, start, end in matches:
        span = nlp_text[start:end]
        return span.text

def extract_emails(resume_text):
    return re.findall(EMAIL_REG, resume_text)

def extract_skills(input_text):
    stop_words = set(nltk.corpus.stopwords.words('english'))
    word_tokens = nltk.tokenize.word_tokenize(input_text)
 
    filtered_tokens = [w for w in word_tokens if w not in stop_words]
    filtered_tokens = [w for w in word_tokens if w.isalpha()]
 
    bigrams_trigrams = list(map(' '.join, nltk.everygrams(filtered_tokens, 2, 3)))
 
    found_skills = set()
 
    for token in filtered_tokens:
        if token.lower() in SKILLS_DB:
            found_skills.add(token.lower())
 
    for ngram in bigrams_trigrams:
        if ngram.lower() in SKILLS_DB:
            found_skills.add(ngram.lower())
 
    return found_skills

def extract_education(resume_text):
    STOPWORDS = set(stopwords.words('english'))
    nlp_text = nlp(resume_text)

    nlp_text = [sent.text.strip() for sent in nlp_text.sents]

    edu = {}

    for index, text in enumerate(nlp_text):
        for tex in text.split():
            tex = re.sub(r'[?|$|.|!|,]', r'', tex)
            if tex.upper() in EDUCATION and tex not in STOPWORDS:
                return tex


def add_information_to_csv(rows):
    fields = ['name','education','graduation year','email','phone','skills']
    csv_file = "records.csv"
    with open(csv_file, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(fields)
        csvwriter.writerows(rows)

def extract_information(text, filename):
    phone = extract_phone_number(text)
    email = extract_emails(text)
    skills = extract_skills(text)
    education = extract_education(text)
    year = extract_year_of_graduation(text)
    name = extract_names(text)

    row = []
    row.append(name)
    row.append(education)
    row.append(year)
    row.append(email[0])
    row.append(phone)
    row.append(skills)
    global index
    INFO_ROWS.append(row)
    index+=1

if __name__ == '__main__':
    for filename in os.listdir(DIRECTORY):
        f = os.path.join(DIRECTORY, filename)
        if os.path.isfile(f):
            if f.endswith('.txt'):
                filevar = open(filename, 'r')
                text = filevar.read()
                filevar.close()
                extract_information(text, filename)
            if f.endswith('.docx'):
                text = extract_text_from_docx(f)
                extract_information(text, filename)
            elif f.endswith('.pdf'):
                text = extract_text_from_pdf(f)
                extract_information(text, filename)
            else:
                print("Unsupported file format")

    add_information_to_csv(INFO_ROWS)
 
