import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from nltk.sentiment import SentimentIntensityAnalyzer
import os


def analyze_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file_path:
        return
    
    try:
        df = pd.read_excel(file_path)
        if 'Comments' not in df.columns:
            messagebox.showerror("Error", "The Excel file does not contain a 'Comments' column.")
            return
        comments = df['Comments'].dropna().tolist()

        if not comments:
            messagebox.showinfo("Analysis Result", "No comments found in the 'Comments' column.")
            return

        positive_count = 0
        negative_count = 0
        neutral_count = 0

        sia = SentimentIntensityAnalyzer()
        for comment in comments:
            sentiment_score = sia.polarity_scores(comment)['compound']
            if sentiment_score >= 0.05:
                positive_count += 1
            elif sentiment_score <= -0.05:
                negative_count += 1
            else:
                neutral_count += 1

        # Display analysis result within the same window
        result_text = f"Positive Comments: {positive_count}\nNegative Comments: {negative_count}\nNeutral Comments: {neutral_count}"
        analysis_result_label.config(text=result_text)
        
        #messagebox.showinfo("Analysis Result", f"Positive Comments: {positive_count}\nNegative Comments: {negative_count}\nNeutral Comments: {neutral_count}")
        
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

root = tk.Tk()
root.title("Sentiment Analysis from Excel")
root.geometry("300x200")
root.config(bg="black")

# Define colors
button_bg_color = "#007bff"
button_text_color = "white"

label_bg_color = "black"
label_text_color = "white"

# Set background color
root.config(bg="black")

# Add some color to the button
analyze_button = tk.Button(root, text="Select Excel File", command=analyze_excel, bg=button_bg_color, fg=button_text_color)
analyze_button.pack(pady=20)


# Label to display analysis result
analysis_result_label = tk.Label(root, text="", bg=label_bg_color, fg=label_text_color)
analysis_result_label.pack()

root.mainloop()
