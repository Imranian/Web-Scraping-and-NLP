import re
import os
import openpyxl
import pandas as pd
import nltk
from nltk.tokenize import sent_tokenize, word_tokenize
from nltk.corpus import stopwords

                                                            # File
folder_path = "G:/Blackcoffer_Assignment/text_excel_files"
excel_files = [file for file in os.listdir(folder_path) if file.endswith('.xlsx')]
sorted_excel = sorted(excel_files, key=lambda x: float(os.path.splitext(x)[0]))

# Create an empty DataFrame to store the results
results_dfs = []

for excel_file in sorted_excel:
    text_path = os.path.join(folder_path, excel_file)
    # text_path = 'C:/Users/imran/OneDrive/Documents/VS code/123.0.xlsx'
    df = pd.read_excel(text_path)
    text = df['Text'].iloc[0]

    # total number of Sentences
    sentences = sent_tokenize(text)
    sentences_total = len(sentences) 

                                                                #CLEANING
    # clean the Text for calculation purpose
    replacements = {
        '!': '',
        '@': '',
        '#': '',
        '$': '',
        '%': '',
        '^': '',
        '&': '',
        '*': '',
        ',': '',
        '.': '',
    }
    for char, replacement in replacements.items():
            changed_data_1 = text.replace(char, replacement)

    # lets remove the Digits for calculation purpose
    a1 = word_tokenize(changed_data_1)
    final_text_file = ''
    for i in a1:
        if i.isalpha():
            final_text_file += i + " "

    # create a variable that holds the entire cleaned text file
    Final_changed_data = word_tokenize(final_text_file)
        

                                                                # Text analysis
    # define Positive, Negative and stop words
    with open("C:/Users/imran/Downloads/MasterDictionary-20230819T044902Z-001/MasterDictionary/positive-words.txt", 'r') as f:
        positive_words = f.read().splitlines()

    with open("C:/Users/imran/Downloads/MasterDictionary-20230819T044902Z-001/MasterDictionary/negative-words.txt", 'r') as g:
        negative_words = g.read().splitlines()

    stopwords_files = ["C:/Users/imran/Downloads/StopWords-20230819T044902Z-001/StopWords/StopWords_Auditor.txt",
                    "C:/Users/imran/Downloads/StopWords-20230819T044902Z-001/StopWords/StopWords_Currencies.txt",
                    "C:/Users/imran/Downloads/StopWords-20230819T044902Z-001/StopWords/StopWords_DatesandNumbers.txt",
                    "C:/Users/imran/Downloads/StopWords-20230819T044902Z-001/StopWords/StopWords_Generic.txt",
                    "C:/Users/imran/Downloads/StopWords-20230819T044902Z-001/StopWords/StopWords_GenericLong.txt",
                    "C:/Users/imran/Downloads/StopWords-20230819T044902Z-001/StopWords/StopWords_Geographic.txt",
                    "C:/Users/imran/Downloads/StopWords-20230819T044902Z-001/StopWords/StopWords_Names.txt"]
    stopwords_list = []

    for files in stopwords_files:
        with open(files, 'r') as file:
            stopwords_list.extend([word.strip() for word in file.readlines()])

    # define total number of syllables present in a word
    def count_syllables(word):
        return max(1, len([i for i in word.lower() if i in "aeiouy"]))

    def count_personal_pronouns(text):
        pattern = r'\b(?:I|we|my|ours|us)\b'
        matches = re.findall(pattern, text, flags=re.IGNORECASE)

        return len(matches)


    # start Analysis
    positive_score = 0
    negative_score = 0

    complex_words = []
    total_syllables = []
    total_personal_pronouns = []

    words = Final_changed_data
    cleaned_words = [word for word in words if word not in stopwords_list]
    total_words = len(words)
    total_characters = len(text)


    for word in cleaned_words:
        if word in positive_words:
            # print('p ->', word)
            positive_score += 1
                
        elif word in negative_words:
            # print('N ->', word)
            negative_score += 1
    for wd in words:
        syllables = count_syllables(wd)
        total_syllables.append(syllables)
        if syllables>2:
            complex_words.append(wd)
    for w in words:
        pronouns = count_personal_pronouns(w)
        total_personal_pronouns.append(pronouns)

                                                                # Calculation
    # formulae
    polarity_score = (positive_score - negative_score) / (positive_score + negative_score + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (total_words + 0.000001)
    average_sentence_length = total_words/sentences_total
    percentage_of_complex_words = len(complex_words)/total_words
    fog_index = (0.4)*(average_sentence_length + percentage_of_complex_words)
    average_no_of_words_per_sentence = total_words/sentences_total
    complex_word_count = len(complex_words)
    word_count = len(cleaned_words)
    avg_syllable_count_per_word = (sum(total_syllables)/len(total_syllables))
    personal_pronouns = sum(total_personal_pronouns)
    avg_word_length = total_characters/total_words


    # print(positive_score)
    # print(negative_score)
    # print(polarity_score)
    # print(subjectivity_score)
    # # print("no. of sentences = ", sentences_total)
    # print(average_sentence_length)
    # # print("total words = ", total_words)
    # # print("Total complex words = ", len(complex_words))
    # print(percentage_of_complex_words)
    # print(fog_index)
    # print(average_no_of_words_per_sentence)
    # print(complex_word_count)
    # print(word_count)
    # print(avg_syllable_count_per_word)
    # print(personal_pronouns)
    # print(avg_word_length)
    
    results_df = pd.DataFrame({
        'POSITIVE SCORE': [positive_score],
        'NEGATIVE SCORE': [negative_score],
        'POLARITY SCORE': [polarity_score],
        'SUBJECTIVITY SCORE': [subjectivity_score],
        'AVG SENTENCE LENGTH': [average_sentence_length],
        'PERCENTAGE OF COMPLEX WORDS': [percentage_of_complex_words],
        'FOG INDEX': [fog_index],
        'AVG NUMBER OF WORDS PER SENTENCE': [average_no_of_words_per_sentence],
        'COMPLEX WORD COUNT': [complex_word_count],
        'WORD COUNT': [word_count],
        'SYLLABLE PER WORD': [avg_syllable_count_per_word],
        'PERSONAL PRONOUNS': [personal_pronouns],
        'AVG WORD LENGTH': [avg_word_length],
    })
    # Append the result_df to the list of DataFrames
    results_dfs.append(results_df)

# Concatenate all DataFrames in the list into a single DataFrame
results_df = pd.concat(results_dfs, ignore_index=True)

# Read the existing Excel file into a DataFrame
output_path = "C:/Users/imran/Downloads/Output Data Structure.xlsx"
existing_df = pd.read_excel(output_path)

# Update existing_df with the calculated values from results_df
existing_df.update(results_df)

# Write the updated DataFrame to the same Excel file
existing_df.to_excel(output_path, index=False, engine="openpyxl")

