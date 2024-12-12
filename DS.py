import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
import requests
import pandas as pd
import os
import re
from nltk.tokenize import word_tokenize
import roman
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
import string
import nltk
nltk.download('punkt')

df = pd.read_excel('input.xlsx')

# Loop through each row and extract URL and URL_ID
urls = []
for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    urls.append((url_id, url))
    
def extract_article_text(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    article_title = soup.find('title').get_text()
    article_text = ''
    article_body = soup.find('body')

    for p in article_body.find_all('p'):
        if p.parent.name not in ['header', 'footer', 'nav']:
            article_text += p.get_text() + '\n'

    return article_title, article_text

os.makedirs('extracted_articles', exist_ok=True)

for url_id, url in urls:
    article_title, article_text = extract_article_text(url)
    
    file_name = re.sub(r'[^\w\s]', '', url_id) + '.txt'
    file_path = os.path.join('extracted_articles', file_name)

    with open(file_path, 'w',encoding='utf-8') as f:
        f.write(article_text[1099:-82])

    # print(f'Article extracted and saved: {file_name}')
    
    
    # Load stop words from .txt files in a folder
def load_stop_words(folder_path):
    stop_words = set()
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if file_name.endswith('.txt'):  # Check if it's a text file
            with open(file_path, 'r', encoding='latin1') as file:
                for line in file:
                    stop_words.add(line.strip()) 
    return stop_words

# for roman words 
def is_roman_numeral(word):
    try:
        roman.fromRoman(word)
        return True
    except roman.InvalidRomanNumeralError:
        return False
    
    
#To load text from .txt files in a folder
text_data_dict = {}
def text_data_extracted_articles(folder_path):
    
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.txt'):
            file_path = os.path.join(folder_path, file_name)
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    text_data = file.read()
            except UnicodeDecodeError:
                with open(file_path, 'r', encoding='latin1') as file:
                    text_data = file.read()
            
            # Replace Roman numerals with their equivalent integers
            words = text_data.split()
            for i in range(len(words)):
                if is_roman_numeral(words[i]):
                    words[i] = str(roman.fromRoman(words[i]))
            
                text_data = ' '.join(words)
                text_data_dict[file_name] = text_data   
    return text_data_dict
    
text_data_dict = text_data_extracted_articles('extracted_articles')


# to clean up text with stop words single texts
def clean_text_with_custom_stopwords(text, custom_stop_words):
    words = word_tokenize(text)
   
    filtered_words = [word for word in words if word.lower() not in custom_stop_words]
   
    cleaned_text = ' '.join(filtered_words)
    
    return cleaned_text
    print(cleaned_text)

 
# to clean up text with stop words multi texts
def clean_and_save_text(text, stop_words, output_folder, file_name):
    cleaned_text = clean_text_with_custom_stopwords(text,stop_words)
    output_path = os.path.join(output_folder, f"{file_name}")
    with open(output_path, 'w', encoding='utf-8') as output_file:
        output_file.write(cleaned_text)


stop_words = load_stop_words("StopWords")

folder_path="extracted_articles"

text_data_extracted_articles(folder_path)

output_folder = 'cleaned_texts'

os.makedirs(output_folder, exist_ok=True)

for file_name, text in text_data_dict.items():
    clean_and_save_text(text, stop_words, output_folder, file_name)
    print(f"Cleaned text saved as '{file_name}.txt'")
    
    
    
def load_positive_negative_stopwords(folder_path):
    
    positive_stopwords = set()
    
    negative_stopwords = set()
    
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if file_name.endswith('.txt'):
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    for word in file:
                        word = word.strip()
                        if 'positive' in file_name:
                            positive_stopwords.add(word)
                        elif 'negative' in file_name:
                            negative_stopwords.add(word)
                        
            except UnicodeDecodeError:
                with open(file_path, 'r', encoding='latin1') as file:
                    for word in file:
                        word = word.strip()
                        if 'positive' in file_name:
                            positive_stopwords.add(word)
                        elif 'negative' in file_name:
                            negative_stopwords.add(word)
                            
    return positive_stopwords, negative_stopwords

positive_stopwords, negative_stopwords = load_positive_negative_stopwords("MasterDictionary")



# Load cleaned texts
def load_cleaned_texts(folder_path):
    cleaned_texts = {}
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if file_name.endswith('.txt'):
            with open(file_path, 'r', encoding='utf-8') as file:
                text = file.read()
                cleaned_texts[file_name] = text
    return cleaned_texts

# Count positive and negative words in cleaned texts
def count_positive_negative_words(cleaned_texts, positive_stopwords, negative_stopwords):
    positive_counts = {}
    negative_counts = {}
    for file_name, text in cleaned_texts.items():
        positive_count = sum(1 for word in text.split() if word.lower() in positive_stopwords)
        negative_count = sum(1 for word in text.split() if word.lower() in negative_stopwords)
        positive_counts[file_name] = positive_count
        negative_counts[file_name] = negative_count
    return positive_counts, negative_counts


positive_stopwords, negative_stopwords = load_positive_negative_stopwords("MasterDictionary")
cleaned_texts = load_cleaned_texts("cleaned_texts")
positive_counts, negative_counts = count_positive_negative_words(cleaned_texts, positive_stopwords, negative_stopwords)




def clean_punctuations_and_save_text(text_data,output_folder, file_name):
                    sentences = sent_tokenize(text_data)
                    words = word_tokenize(text_data)
                    cleaned_words = [word.lower() for word in words if word not in string.punctuation]
                    word_string = ' '.join(cleaned_words)
                    output_path = os.path.join(output_folder, f"{file_name}")
                    with open(output_path, 'w', encoding='utf-8') as output_file:
                        output_file.write(word_string)
                        

text_data_punctuations_dict=text_data_extracted_articles('cleaned_texts')

output_folder = 'cleaned_punctuations_texts'
os.makedirs(output_folder, exist_ok=True)
        
for file_name ,text in text_data_punctuations_dict.items():
    clean_punctuations_and_save_text(text,output_folder, file_name)

def All_Score_counts(cleaned_words,text,Positive,Negative):
    
    sentences = sent_tokenize(text)
    
    positive_score = (Positive)
    
    negative_score = (Negative)
    
    polarity_score = ((positive_score) - (negative_score)) / ((positive_score + negative_score) + 0.000001)
    
    subjectivity_score = (positive_score + negative_score) / (len(cleaned_words) + 0.000001)
    
    avg_sentence_length = len(cleaned_words) / len(sentences)
    
    complex_words = [word for word in cleaned_words if len(word) > 2]  
    
    percentage_complex_words = len(complex_words) / len(cleaned_words)
    
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    
    avg_words_per_sentence = len(cleaned_words) / len(sentences)
    
    complex_word_count = len([word for word in cleaned_words if len(word) > 2])
    
    cleaned_text = text.translate(str.maketrans('', '', string.punctuation)).lower()
    
    words = cleaned_text.split()
    word_count=len(words)
    
    def syllable_count(word):
        word = word.lower()
        if word.endswith(('es', 'ed')):
            return 0  
        vowels = 'aeiou'
        count = 0
        for index, char in enumerate(word):
            if char in vowels and (index == 0 or word[index - 1] not in vowels):
                count += 1
        return max(count, 1) 
    syllable_count_per_word = [syllable_count(word) for word in cleaned_words]

    personal_pronouns = re.findall(r'\b(?:i|we|my|ours|us)\b', text.lower())
    personal_pronoun_count = len(personal_pronouns)
    
    average_word_length = sum(len(word) for word in cleaned_words) / len(cleaned_words)
    
    return positive_score,negative_score,polarity_score,subjectivity_score,avg_sentence_length,percentage_complex_words,fog_index,avg_words_per_sentence,complex_word_count,syllable_count_per_word,personal_pronoun_count,average_word_length,word_count



text_read_data_punctuations_dict=text_data_extracted_articles('cleaned_punctuations_texts')
positive_stopwords, negative_stopwords = load_positive_negative_stopwords("MasterDictionary")
def all_values():
    for file_name ,text in text_read_data_punctuations_dict.items():
        
        word_list = text.split()
        
        if positive_counts.get(file_name)==None:
            Positive=0
        else:
            Positive=positive_counts.get(file_name)
        if negative_counts.get(file_name)==None:
            Negative=0
        else:
            Negative=negative_counts.get(file_name)
        print(file_name)
        positive_score,negative_score,polarity_score,subjectivity_score,avg_sentence_length,percentage_complex_words,fog_index,avg_words_per_sentence,complex_word_count,syllable_count_per_word,personal_pronoun_count,average_word_length,word_count=All_Score_counts(word_list,text,Positive,Negative)
    # print(file_name)    
   
 # return positive_score,negative_score,polarity_score,subjectivity_score,avg_sentence_length,percentage_complex_words,fog_index,avg_words_per_sentence,complex_word_count,syllable_count_per_word,personal_pronoun_count,average_word_length
 import openpyxl

# Load the Excel workbook
def outputcsv():
    workbook = openpyxl.load_workbook('Output Data Structure.xlsx')
    sheet = workbook.active
    for file_name ,text in text_read_data_punctuations_dict.items():
        word_list = text.split()
        if positive_counts.get(file_name)==None:
            Positive=0
        else:
            Positive=positive_counts.get(file_name)
            
        if negative_counts.get(file_name)==None:
            Negative=0
        else:
            Negative=negative_counts.get(file_name)
        positive_score,negative_score,polarity_score,subjectivity_score,avg_sentence_length,percentage_complex_words,fog_index,avg_words_per_sentence,complex_word_count,syllable_count_per_word,personal_pronoun_count,average_word_length,word_count=All_Score_counts(word_list,text,Positive,Negative)
        remove=".txt"
        name=file_name.replace(remove,'')
        for i in range(2,102):
            if sheet['A'+str(i)].value == name :
                sheet['C'+str(i)] = positive_score #POSITIVE SCORE
                sheet['D'+str(i)] = negative_score  #NEGATIVE SCORE
                
                sheet['E'+str(i)] = polarity_score  #POLARITY SCORE
                
                sheet['F'+str(i)] = subjectivity_score #SUBJECTIVITY SCORE

                sheet['G'+str(i)] = avg_sentence_length  #AVG SENTENCE LENGTH

                sheet['H'+str(i)] = percentage_complex_words  #PERCENTAGE OF COMPLEX WORDS

                sheet['I'+str(i)] = fog_index #FOG INDEX

                sheet['J'+str(i)] = avg_words_per_sentence #AVG NUMBER OF WORDS PER SENTENCE

                sheet['K'+str(i)] = complex_word_count  #COMPLEX WORD COUNT

                sheet['L'+str(i)] = word_count  # WORD COUNT

                sheet['M'+str(i)] = "-"  #SYLLABLE PER WORD beacause of list

                sheet['N'+str(i)] = personal_pronoun_count #PERSONAL PRONOUNS

                sheet['O'+str(i)] = average_word_length  #AVG WORD LENGTH
        print(name)

    workbook.save('Final Output .xlsx')
    
outputcsv()