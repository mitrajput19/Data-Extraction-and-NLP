import nltk
from nltk.corpus import cmudict
import pandas as pd
from bs4 import BeautifulSoup
import requests
import os
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
import re

nltk.download('punkt')
nltk.download('cmudict')
nltk.download('stopwords')

stopwords1 = set()
for filename in os.listdir("StopWords"):
    with open(os.path.join("StopWords", filename), 'r') as file:
        stopwords1.update([line.strip() for line in file.readlines()])


def count_syllables(word):
    vowels = re.compile(r'[aeiouyAEIOUY]+')
    exceptions = re.compile(r'[a-zA-Z]+(es|ed)$')
    syllables = len(vowels.findall(word))
    if exceptions.search(word):
        return syllables - 1
    return syllables


def count_syllables_in_text(text):
    words = text.split()
    total_syllables = 0
    word_count = 0
    for word in words:
        syllable_count = count_syllables(word)
        total_syllables += syllable_count
        word_count += 1

    if word_count > 0:
        average_syllables_per_word = total_syllables / word_count
    else:
        average_syllables_per_word = 0

    return average_syllables_per_word


def complex_words(text):
    complex_word_count = sum(1 for word in text if count_syllables(word) > 1)
    return complex_word_count


def positive_score_function(text):
    positive_words = set()
    with open("MasterDictionary/positive-words.txt", 'r') as file:
        for line in file:
            positive_words.add(line.strip())
    words = text.split()
    positive_score = 0
    for word in words:
        if word.lower() in positive_words:
            positive_score += 1
    return positive_score


def negative_score_function(text):
    negative_words = set()
    with open("MasterDictionary/negative-words.txt", 'r') as file:
        for line in file:
            negative_words.add(line.strip())
    words = text.split()
    negative_score = 0
    for word in words:
        if word.lower() in negative_words:
            negative_score += 1
    return negative_score


def scrape_website(url, id):
    try:
        print(id)
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')

            # title
            title = soup.find('title').string

            # raw content
            content = ""
            try:
                content = soup.find("div", {"class": "td-post-content tagdiv-type"}).get_text()
            except AttributeError as e:
                for i in soup.find_all('p'):
                    content =content + " " + i.get_text()
            # sentences tokenize
            sentences = sent_tokenize(content)

            # word tokenize
            tokens = word_tokenize(content)
            print(len(tokens))
            filtered_tokens = [word for word in tokens if word not in stopwords1]
            filtered_text = ' '.join(filtered_tokens)

            # positive score
            positive_score = positive_score_function(filtered_text)
            # print(positive_score)

            # negative score
            negative_score = negative_score_function(filtered_text)
            # print(negative_score)

            # polarity score
            polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
            # print(polarity_score)

            # subjectivity score
            subjectivity_score = (positive_score + negative_score) / ((len(filtered_text)) + 0.000001)
            # print(subjectivity_score)

            # avg sentence length
            avg_sentence_length = len(filtered_tokens) / len(sentences)
            # print(avg_sentence_length)

            # complex word percentage
            complex_word_percentage = complex_words(filtered_tokens) / len(filtered_tokens)
            # print(complex_word_percentage)

            # fog index
            fog_index = 0.4 * (avg_sentence_length + complex_word_percentage)
            # print(fog_index)

            # AVG NUMBER OF WORDS PER SENTENCE
            # print(avg_sentence_length)

            # COMPLEX WORD COUNT
            complex_word_count = complex_words(filtered_tokens)
            # print(complex_word_count)

            # Word Count
            stop_words = set(stopwords.words('english'))
            word_count_token = [word for word in tokens if word.lower() not in stop_words]
            punctuation_list = ['!', '"', '#', '$', '%', '&', "'", '(', ')', '*', '+', ',', '-', '.', '/', ':', ';',
                                '<', '=',
                                '>', '?', '@', '[', '\\', ']', '^', '_', '`', '{', '|', '}', '~']
            word_count = [word for word in word_count_token if word not in punctuation_list]
            # print(len(word_count))

            # SYLLABLE PER WORD
            syllable_per_word = count_syllables_in_text(filtered_text)
            # print(syllable_per_word)

            # PERSONAL PRONOUNS
            personal_pronouns = ["i", "me", "my", "mine", "myself", "you", "your", "yours", "yourself", "he", "him",
                                 "his",
                                 "himself", "she", "her", "hers", "herself", "we", "us", "our", "ours", "ourselves",
                                 "they",
                                 "them", "their", "theirs", "themselves", "I", "Me", "My", "Mine", "Myself", "You",
                                 "Your", "Yours", "Yourself", "He", "Him", "His",
                                 "Himself", "She", "Her", "Hers", "Herself", "We", "Us", "Our", "Ours", "Ourselves",
                                 "They",
                                 "Them", "Their", "Theirs", "Themselves"]
            personal_pronouns_count = sum(1 for word in tokens if word in personal_pronouns)
            # print(personal_pronouns_count)

            # AVG WORD LENGTH
            word_lengths = [len(word) for word in tokens]
            avg_word_length = sum(word_lengths) / len(tokens)
            # print(avg_word_length)
            row = {"URL_ID": id, "URL": url, "POSITIVE SCORE": positive_score,
                   "NEGATIVE SCORE": negative_score, "POLARITY SCORE": polarity_score,
                   "SUBJECTIVITY SCORE": subjectivity_score, "AVG SENTENCE LENGTH": avg_sentence_length,
                   "PERCENTAGE OF COMPLEX WORDS": complex_word_percentage,
                   "FOG INDEX": fog_index, "AVG NUMBER OF WORDS PER SENTENCE": avg_sentence_length,
                   "COMPLEX WORD COUNT": complex_word_count,
                   "WORD COUNT": len(word_count), "SYLLABLE PER WORD": syllable_per_word,
                   "PERSONAL PRONOUNS": personal_pronouns_count, "AVG WORD LENGTH": avg_word_length}
            final_data.append(row)
        else:
            print(f"Failed to retrieve data from Status code: {response.status_code}")
    except Exception as e:
        print(f"An error occurred while scraping {url}: {str(e)}")


excel_file_path = 'Input.xlsx'
dataframe = pd.read_excel(excel_file_path)

columns = ["URL_ID", "URL", "POSITIVE SCORE", "NEGATIVE SCORE", "POLARITY SCORE", "SUBJECTIVITY SCORE",
           "AVG SENTENCE LENGTH", "PERCENTAGE OF COMPLEX WORDS", "FOG INDEX", "AVG NUMBER OF WORDS PER SENTENCE",
           "COMPLEX WORD COUNT", "WORD COUNT", "SYLLABLE PER WORD", "PERSONAL PRONOUNS", "AVG WORD LENGTH"]
df = pd.DataFrame(columns=columns)
final_data = []

for index, row in dataframe.iterrows():
    website_url = row['URL']
    website_url_id = row['URL_ID']
    data = scrape_website(website_url, website_url_id)

print(final_data)
for row in final_data:
    df = df._append(row, ignore_index=True)
print(df)
df.to_excel("Output.xlsx", index=False)
print("Completed Successfully")
