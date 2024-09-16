#importing all libraries
import pandas as pd
import os
import glob
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from string import punctuation
from nltk.corpus import wordnet
from nltk.probability import FreqDist
import re
from goose3 import Goose
import csv

#Run these Commands one-time
#nltk.download('wordnet')
#nltk.download('punkt')

#new changes
curr_dic = os.getcwd()
links = pd.read_excel(curr_dic+"/Input.xlsx")

#Scraping all the links
goose = Goose()
for i in range(100):
    try:
        article = goose.extract(links["URL"][i])
        with open(links["URL_ID"][i], 'w') as f:
            data = [article.title,article.cleaned_text]

            for line in data:
                f.write(line + '\n')
                print(line)
    except:
        with open(links["URL_ID"][i], 'w') as f:
            data=['']
            
            for line in data:
                f.write(line + '\n')
                print(line)
        
# Data to be written to the CSV file
data =[
    ['URL_ID', 'URL', 'POSITIVE SCORE','NEGATIVE SCORE','POLARITY SCORE','SUBJECTIVITY SCORE','AVG SENTENCE LENGTH','PERCENTAGE OF COMPLEX WORDS','FOG INDEX','AVG NUMBER OF WORDS PER SENTENCE','COMPLEX WORD COUNT','WORD COUNT','SYLLABLE PER WORD','PERSONAL PRONOUNS','AVG WORD LENGTH']]

# Specify the file name
filename = curr_dic+'/data.csv'

# Writing to CSV file
with open(filename, 'w', newline='') as csvfile:
    csvwriter = csv.writer(csvfile)
    csvwriter.writerows(data)
print(f"CSV file '{filename}' has been created successfully.")

#Reading text files one by one
for i in range(100):

    def read_text_file(file_path):
        try:
           with open(file_path, 'r') as file:
                content = file.read()
           return content
        except FileNotFoundError:
            print("File not found.")
            return None

    file_path = curr_dic+'/'+links["URL_ID"][i]  
    file_content = read_text_file(file_path)

# Combine all Stop Words and store it in a new txt file
    def combine_text_files(folder_path, output_file):
        if not folder_path.endswith('/'):
            folder_path += '/'
        files = glob.glob(folder_path + '*.txt')   # Using glob to get all text files in the folder

    # Open output file in append mode
        with open(output_file, 'w',encoding='iso-8859-1') as combined_file:
            for file in files:
            # Open each text file and read its content
                with open(file, 'r',encoding='iso-8859-1') as f:
                    content = f.read()
                    combined_file.write(content)
                    combined_file.write('\n')

    folder_path = curr_dic+'/StopWords'
    output_file = curr_dic+'/StopWords/combined_text.txt'

    combine_text_files(folder_path, output_file)
    print("Text files combined successfully!")

    #Combining all stop words
    stop_words = pd.read_csv(output_file,delimiter='/t',encoding='iso-8859-1',engine='python')
    all_stop_words = stop_words["SMITH  | Surnames from 1990 census > .002%.  www.census.gov.genealogy/names/dist.all.last"].tolist()

    #Deleting all the Stop Words from the Article
    def remove_stop_words(input_text, all_stop_words):
    # Tokenize input text into words
        words = input_text.split()
    # Remove custom stop words
        cleaned_words = [word for word in words if word not in all_stop_words]
    # Join the remaining words back into a single string
        cleaned_text = ' '.join(cleaned_words)
        return cleaned_text


    file_content = file_content.upper()
    input_text = file_content

    cleaned_text = remove_stop_words(input_text, all_stop_words)
    cleaned_text = cleaned_text.lower()

    #Fetching all the Positive Words
    pos_words = pd.read_csv(curr_dic+"/MasterDictionary/positive-words.txt")["a+"].tolist()

    #Fetching all the Negative Words
    neg_words =pd.read_csv(curr_dic+"/MasterDictionary/negative-words.txt",delimiter='/t', encoding='iso-8859-1',engine='python')["2-faced"].tolist()


    text = cleaned_text
# Tokenize
    tokens = word_tokenize(text)
    words_without_punctuation = [tokens for tokens in tokens if tokens not in punctuation]
    new_tokens = words_without_punctuation
    total_words = len(words_without_punctuation)

    #Complex Words analysis
    complex_words = []
    for word in new_tokens:
       
        if len(word) > 2:  
            synsets = wordnet.synsets(word)
            if len(synsets) > 2:  
                complex_words.append(word)
    complex_words_count = len(complex_words)

    #Syllable per words
    exclude_pattern = re.compile(r'\b\w*(?:es|ed)\b', re.IGNORECASE)
    vowel_count = 0
    
    for word in words_without_punctuation:
        if not re.match(exclude_pattern, word):
            vowel_count += sum(1 for char in word if char.lower() in 'aeiou')
    
    #Personal Pronouns
    pattern = r'\b(?:I|we|my|ours|us)\b'
    regex = re.compile(pattern)
    personal_p_counts = sum(1 for word in words_without_punctuation if regex.match(word.lower()))

# Tokenize into sentences
    sentences = sent_tokenize(text)
    total_sen_count = len(sentences)

    total_word_count = 0
    num_sentences = len(sentences)
    
    for sentence in sentences:
        # Tokenize sentence into words
        words = nltk.word_tokenize(sentence)
        # Increment total word count
        total_word_count += len(words)
    
    # Calculate average word count
    try:
        average_word_count = total_word_count / num_sentences
    except ZeroDivisionError:
        average_word_count = 0

    #Find occurrences of Positive words
    postive_words = [word for word in new_tokens if word in pos_words]
    pos_result = len(postive_words)


    #Find Occurrences of Negative Words
    negative_words = [word for word in new_tokens if word in neg_words]
    neg_result = len(negative_words)
    neg_result_final = neg_result*-(1)

    #Calculating Polarity
    try:
        polarity = (pos_result-neg_result)/(pos_result+neg_result) +0.000001
        print(polarity)
    except ZeroDivisionError:
        polarity = 0

    #Calculating Subjectivity
    subjectivity = (pos_result+neg_result)/(total_words+0.000001)
    
    #Calculating Avg Sentence Length
    try:
        avg_sentence_len = total_words/total_sen_count
        print(avg_sentence_len)
    except ZeroDivisionError:
        avg_sentence_len = 0

    #Calculating Percentage of complex words
    try:
        Percentage_of_complex_words = complex_words_count/total_words
        print(Percentage_of_complex_words)
    except ZeroDivisionError:
        Percentage_of_complex_words = 0

    #Calculating Fog Index
    fog_index = 0.4*(avg_sentence_len+Percentage_of_complex_words)

    #Calculating Avg number of words per sentence
    try:
        avg_no_of_words_per_sentence = total_words/total_sen_count
        print(avg_no_of_words_per_sentence)
    except ZeroDivisionError:
        avg_no_of_words_per_sentence = 0

    #Appending all data to csv file
    def append_to_csv(file_path, data):

        with open(file_path, 'a', newline='') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(data)

    file_path = curr_dic+'/data.csv'
    data_to_append = [links['URL_ID'][i], links['URL'][i], pos_result,neg_result_final,polarity,subjectivity,avg_sentence_len,Percentage_of_complex_words,fog_index,avg_no_of_words_per_sentence,complex_words_count,total_words,vowel_count,personal_p_counts,average_word_count]  # Example data to append
    append_to_csv(file_path, data_to_append)


