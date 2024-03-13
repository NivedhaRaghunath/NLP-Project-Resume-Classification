# -*- coding: utf-8 -*-
"""
Created on Mon Feb 19 11:48:41 2024

@author: sunfa
"""



import pandas as pd 
import pickle
import win32com.client
import nltk
from nltk.corpus import stopwords
import re
from nltk.stem import WordNetLemmatizer
nltk.download('omw')
nltk.download('omw-1.4')
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.model_selection import train_test_split
from sklearn.svm import SVC
from sklearn.metrics import classification_report
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import accuracy_score
import docx2txt
import PyPDF2
import nltk
import os

def extract_txt_docx(filepath):
    txt = docx2txt.process(filepath)
    return txt.replace('\n', '').replace("\t","")



def extract_txt_pdf(filepath):
    txt = ""
    with open(filepath, "rb") as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page_number in range(len(pdf_reader.pages)):
            txt += pdf_reader.pages[page_number].extract_text()
    return txt.replace('\n', '').replace("\t","") 

def extract_doc_document(filepath):
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(filepath)
    text = doc.Content.Text
    doc.Close()
    word.Quit()
    return text.replace('\n', '').replace("\t","") 

intern_folder = r"D:\DATA SCIENCE\PROJECT NLP\Dataset\FinalInternship resumes"

intern_data = []

for i in os.listdir(intern_folder):
    file_path = os.path.join(intern_folder, i)
    if i.endswith(".docx"):
        content = extract_txt_docx(file_path).strip()
    
    intern_data.append((content, "Intership"))

intern_df = pd.DataFrame(intern_data, columns=["Resume Content", "Category"])



react_folder = r'D:\DATA SCIENCE\PROJECT NLP\Dataset\FinalReact resumes'
react_data = []

for i in os.listdir(react_folder):
    file_path = os.path.join(react_folder, i)
    if i.endswith(".pdf"):
        content = extract_txt_pdf(file_path)
    elif i.endswith("docx"):
        content = extract_txt_docx(file_path).strip()
    else:
        content = extract_doc_document(file_path).strip()
    
    react_data.append((content, "React"))

react_df = pd.DataFrame(react_data, columns=["Resume Content", "Category"])

ps_folder = r'D:\DATA SCIENCE\PROJECT NLP\Dataset\FinalPeoplesoft resumes'


ps_data = []

for i in os.listdir(ps_folder):
    file_path = os.path.join(ps_folder, i)
    if i.endswith(".docx"):
        content = extract_txt_docx(file_path).strip()
    else:
        content = extract_doc_document(file_path).strip()
    
    ps_data.append((content, "PeopleSoft"))

ps_df = pd.DataFrame(ps_data, columns=["Resume Content", "Category"])


sql_folder = r'D:\DATA SCIENCE\PROJECT NLP\Dataset\FinalSQL Developer Lightning insight'


sql_data = []

for i in os.listdir(sql_folder):
    file_path = os.path.join(sql_folder, i)
    if i.endswith(".docx"):
        content = extract_txt_docx(file_path).strip()
    else:
        content = extract_doc_document(file_path).strip()
    
    sql_data.append((content, "SQL"))

sql_df = pd.DataFrame(sql_data, columns=["Resume Content", "Category"])

workday_folder = r'D:\DATA SCIENCE\PROJECT NLP\Dataset\Finalworkday resumes'


workday_data = []

for i in os.listdir(workday_folder):
    file_path = os.path.join(workday_folder, i)
    if i.endswith(".docx"):
        content = extract_txt_docx(file_path).strip()
    else:
        content = extract_doc_document(file_path).strip()
    
    workday_data.append((content, "Workday"))

workday_df = pd.DataFrame(workday_data, columns=["Resume Content", "Category"])


combined_data = pd.concat([react_df,ps_df,sql_df,workday_df,intern_df]).reset_index(drop=True)
    
stop_words = set(stopwords.words("english"))

lemmatizer = WordNetLemmatizer()

def clean_text(text):
    text = text.strip()
    text = text.lower()
    text = re.sub(r"http\S+","",text)
    text = re.sub(r'[^a-zA-Z\s]', '', text)
    words = text.split()
    cleaned_words = [lemmatizer.lemmatize(word) for word in words if word not in stop_words]
    cleaned_text = ' '.join(cleaned_words)

    return cleaned_text

clean_data= combined_data.copy()
clean_data["Resume Content"] = clean_data["Resume Content"].apply(clean_text)


tfidf_vectorizer = TfidfVectorizer()


X_tfidf = tfidf_vectorizer.fit_transform(clean_data['Resume Content'])

X_tfidf_df = pd.DataFrame(X_tfidf.toarray(), columns=tfidf_vectorizer.get_feature_names_out())

pickle.dump(tfidf_vectorizer,open("transform.pkl","wb"))

label_encoder = LabelEncoder()
clean_data['encoded_category'] = label_encoder.fit_transform(clean_data['Category'])

Y_labled = clean_data['encoded_category']

pickle.dump(label_encoder,open("encoder.pkl","wb"))

x = X_tfidf_df
y = Y_labled

X_train, X_test, y_train, y_test = train_test_split(x,y, test_size=0.2, random_state=42)

svm_classifier = SVC(kernel='linear')  # You can change the kernel as needed
svm_classifier.fit(X_train, y_train)

svm_predictions_test = svm_classifier.predict(X_test)

print("Classification Report for SVM:")
print(classification_report(y_test, svm_predictions_test))

svm_predictions_train = svm_classifier.predict(X_train)
accuracy_test_svm = accuracy_score(y_test, svm_predictions_test)
accuracy_train_svm = accuracy_score(y_train, svm_predictions_train)
print("Testing Accuracy for Naive Bayes:",accuracy_test_svm )
print("Training Accuracy for Naive Bayes:",accuracy_test_svm )


filename = "nlp_model.pkl"
pickle.dump(svm_classifier,open(filename,"wb"))