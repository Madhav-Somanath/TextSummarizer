import docx2txt
import re
import heapq
import nltk
import docx

text = docx2txt.process("input.docx")

# Removing square brackets and extra spaces
article_text = re.sub(r'\[[0-9]*\]', ' ', text)
article_text = re.sub(r'\s+', ' ',article_text )

# Tokenization
sentence_list = nltk.sent_tokenize(article_text)

# Removing the stopwords
stopwords = nltk.corpus.stopwords.words('english')

# Storing the word frequencies
word_frequencies = {}
for word in nltk.word_tokenize(article_text):
    if word not in stopwords:
        if word not in word_frequencies.keys():
            word_frequencies[word] = 1
        else:
            word_frequencies[word] += 1
            
# Finding the weighted frequency
maximum_frequncy = max(word_frequencies.values())

for word in word_frequencies.keys():
    word_frequencies[word] = (word_frequencies[word]/maximum_frequncy)
    
# Calculating sentence scores
sentence_scores = {}
for sent in sentence_list:
    for word in nltk.word_tokenize(sent.lower()):
        if word in word_frequencies.keys():
            if len(sent.split(' ')) < 25:
                if sent not in sentence_scores.keys():
                    sentence_scores[sent] = word_frequencies[word]
                else:
                    sentence_scores[sent] += word_frequencies[word]
                    
# Optimizing for weight of "10"
summary_sent = heapq.nlargest(10, sentence_scores, key=sentence_scores.get)

summary = ' '.join(summary_sent)
print(summary)

mydoc = docx.Document()
mydoc.add_paragraph(summary)
mydoc.save("output.docx")