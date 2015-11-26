import re
from nltk import PorterStemmer
import xlrd
import numpy as np
import lda

def stem(word):
    return PorterStemmer().stem_word(word)

def remove_common(words):
    #common = set('for a the an to in and but from he she i am is'.split())
    common = set([stem(word.lower()) for word in
    #common = set([word.lower() for word in
        open("common_words.txt").read().split()])

    return [word for word in words if word not in common and
               len(word) > 3] # Change the length of words

def save_file(words, f):
    f.write(','.join(words))
    f.write('\n')

def parse_text(text):
    s = re.sub(r'\.[^.]+$', '.', text)
    if '-' in s:
        s = s[s.index('-') + 1:]

    s = s.lower()
    words = re.findall(r'[a-z]+', s)
    return words

def read_xls(fname, start_row):
    xl_workbook = xlrd.open_workbook(fname)
    xl_sheet = xl_workbook.sheet_by_index(0)

    documents = []
    for row_idx in range(start_row, xl_sheet.nrows):
        row = xl_sheet.row(row_idx)[0]
        text = remove_common(parse_text(row.value))
        documents.append(text)
    return documents

# Configuration
fname = "text.xlsx"
start_row = 1 # First Row is the column title

documents = read_xls(fname, start_row)
unique_words = set()

print "Reading and parsing the documents..."
for document in documents:
    for word in document:
        unique_words.add(stem(word))
        #unique_words.add(word)
vocab = tuple(unique_words)

print "Creating the LDA matrix..."
X = np.zeros((len(documents), len(vocab)), "int32")
for i in range(0, len(documents)):
    for word in documents[i]:
        X[i, vocab.index(stem(word))] += 1
        #X[i, vocab.index(word)] += 1

print "Performing LDA..."
model = lda.LDA(n_topics=20, n_iter=1500, random_state=1)
model.fit(X)
topic_word = model.topic_word_
n_top_words = 8

for i, topic_dist in enumerate(topic_word):
    topic_words = np.array(vocab)[np.argsort(topic_dist)][:-(n_top_words+1):-1]
    print('Topic {}: {}'.format(i, ' '.join(topic_words)))
