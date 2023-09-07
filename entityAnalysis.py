import spacy
from textblob import TextBlob
import pandas as pd
from tabulate import tabulate
from openpyxl import load_workbook

# Load spaCy model for NER
nlp = spacy.load("en_core_web_sm")

# Load NRC-Emotion-Lexicon dataset
emolex_df = pd.read_csv('NRC-Emotion-Lexicon.csv', encoding='utf-8')

# Read input text from file in utf-8 format
with open("speech.txt", "r", encoding="utf-8") as file:
    text = file.read()

text = text.replace("'s", '')
text = text.replace("’s", '')

# Perform named entity recognition (NER)
doc = nlp(text)
entities = {}
for entity in doc.ents:
    if entity.label_ in ["PERSON", "ORG"]:
        entity_text = entity.text.lower()
        if entity_text not in entities:
            entities[entity_text] = []

# Group related sentences for each entity
for sentence in doc.sents:
    for entity in entities.keys():
        if entity in sentence.text.lower():
            entities[entity].append(sentence.text)

# Perform sentiment analysis for each entity
emotions_label = ['Anger', 'Anticipation', 'Disgust', 'Fear', 'Joy', 'Sadness', 'Surprise', 'Trust']

def get_emotions(text):
    emotion_counts = {emotion: 0 for emotion in emotions_label}
    for word in text.split():
        word = word.lower()
        if word in emolex_df['English (en)'].str.lower().values:
            word_emotions = emolex_df[emolex_df['English (en)'].str.lower() == word]
            for emotion in emotions_label:
                emotion_counts[emotion] += word_emotions[emotion].values[0]
    return emotion_counts

entity_emotions_list = []
for entity, sentences in entities.items():
    combined_text = " ".join(sentences)
    blob = TextBlob(combined_text)
    sentiment = blob.sentiment.polarity
    sentiment_label = "Positive" if sentiment > 0 else "Negative" if sentiment < 0 else "Neutral"
    emotion_counts = get_emotions(combined_text)
    entity_emotions = {"Entity": entity, "Sentiment Polarity": sentiment, "Sentiment Label": sentiment_label}
    entity_emotions.update(emotion_counts)
    entity_emotions_list.append(entity_emotions)

# Create pandas DataFrame from list of entity emotions
df = pd.DataFrame(entity_emotions_list)
# Reorder the columns for better alignment
cols = ["Entity", "Sentiment Polarity", "Sentiment Label"] + emotions_label
df = df[cols]

# Convert the numbers to numeric values
df[emotions_label + ["Sentiment Polarity"]] = df[emotions_label + ["Sentiment Polarity"]].apply(pd.to_numeric)

# Write DataFrame to Excel file
output_file = "output.xlsx"
df.to_excel(output_file, index=False)

import openpyxl

# Load the Excel output file
output_file = "output.xlsx"
wb = openpyxl.load_workbook(output_file)
ws = wb.active

# Iterate through all columns and adjust the column width to fit the text
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter  # Get the column letter (e.g., 'A', 'B', etc.)
    for cell in col:
        try:
            # Get the length of the cell value
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    # Adjust the column width based on the maximum length of cell value
    ws.column_dimensions[column].width = max_length + 2  # Add some buffer space

# Save the modified Excel file
wb.save("output.xlsx")
print("Column widths adjusted and saved to 'output.xlsx'")

