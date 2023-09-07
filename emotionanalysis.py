from textblob import TextBlob
import matplotlib.pyplot as plt
import pandas as pd

def get_emotions(text):
    emotion_counts = [0] * 8
    for word in text.split():
        word = word.lower()
        if word in emolex_df['English (en)'].str.lower().values:
            word_emotions = emolex_df[emolex_df['English (en)'].str.lower() == word]
            for i, emotion in enumerate(emotions_label):
                emotion_counts[i] += word_emotions[emotion].values[0]
    return emotion_counts

# Read input text from file
with open('speech.txt', 'r', encoding='utf-8') as file:
    text = file.read()

# Load NRC-Emotion-Lexicon dataset
emolex_df = pd.read_csv('NRC-Emotion-Lexicon.csv', encoding='utf-8')

# Get emotion counts
emotions_label = ['Anger', 'Anticipation', 'Disgust', 'Fear', 'Joy', 'Sadness', 'Surprise', 'Trust']
emotion_counts = get_emotions(text)

# Display results
print("Emotion Analysis:")
for i, emotion in enumerate(emotions_label):
    print(emotion, ": ", emotion_counts[i])

# Plot emotions as bar chart with darker colors
def plot_emotions(emotion_counts):
    colors = ['#8B0000', '#FFA500', '#556B2F', '#8B4513', '#FFD700', '#2E8B57', '#483D8B', '#9400D3']
    plt.figure(figsize=(10, 6))
    plt.bar(emotions_label, emotion_counts, color=colors)
    plt.title("Emotion Analysis")
    plt.xlabel("Emotion")
    plt.ylabel("Count")
    plt.show()

plot_emotions(emotion_counts)
