import matplotlib.pyplot as plt
from textblob import TextBlob
import nltk
nltk.download('stopwords')
from nltk.corpus import stopwords

# Function for tokenization
def tokenize_text(text):
    blob = TextBlob(text)
    tokens = [token.lower() for sentence in blob.sentences for token in sentence.words]
    return tokens

# Function for stopword removal
def remove_stopwords(tokens):
    stop_words = set(stopwords.words("english"))
    filtered_words = [word for word in tokens if word not in stop_words]
    return filtered_words

# Function for sentiment analysis
def perform_sentiment_analysis(text):
    analysis = TextBlob(text)
    sentiment_polarity = analysis.sentiment.polarity
    sentiment_subjectivity = analysis.sentiment.subjectivity
    return sentiment_polarity, sentiment_subjectivity

# Function for plotting sentiment analysis results
def plot_sentiment_analysis(sentiment_polarity, sentiment_subjectivity):
    labels = ['Polarity', 'Subjectivity']
    x = [0, 1]
    width = 0.35
    
    fig, ax = plt.subplots(figsize=(6, 6))
    
    rects1 = ax.bar(x, [sentiment_polarity, sentiment_subjectivity], width, color=['y', 'g'])
    
    ax.set_ylabel('Sentiment Score')
    ax.set_title('Sentiment Analysis Results')
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    
    fig.tight_layout()
    
    plt.show()

# Read text from file
with open("speech.txt", "r", encoding="utf-8") as file:
    text = file.read()

# Perform text processing tasks
tokens = tokenize_text(text)
filtered_words = remove_stopwords(tokens)
cleaned_text = ' '.join(filtered_words)

# Perform sentiment analysis
sentiment_polarity, sentiment_subjectivity = perform_sentiment_analysis(cleaned_text)

# Print sentiment analysis results
print("Sentiment Analysis Results:")
print("Polarity: {:.2f}".format(sentiment_polarity))
print("Subjectivity: {:.2f}".format(sentiment_subjectivity))

# Plot sentiment analysis results
plot_sentiment_analysis(sentiment_polarity, sentiment_subjectivity)
