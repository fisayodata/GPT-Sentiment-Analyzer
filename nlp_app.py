import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openai

# ✅ Initialize OpenAI API Key
openai.api_key = "Your API Key"

# ✅ Initialize the Tkinter App
app = tk.Tk()
app.title("GPT-4 Sentiment Analysis and Summarization")
app.geometry("600x500")

# Global data variable
data = None

# ✅ Function: Upload CSV
def upload_csv():
    global data
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        data = pd.read_csv(file_path)
        messagebox.showinfo("Success", f"CSV loaded with {len(data)} rows!")
    else:
        messagebox.showerror("Error", "Failed to load CSV")

# ✅ Function: Perform Theme Search
def theme_search():
    global data
    if data is None:
        messagebox.showerror("Error", "Please upload a CSV first!")
        return

    theme = theme_entry.get()
    keywords = keywords_entry.get().split(",")
    keywords = [kw.strip().lower() for kw in keywords]

    data['Theme_Match'] = data['Comment'].apply(lambda x: any(kw in x.lower() for kw in keywords))
    theme_count = data['Theme_Match'].sum()

    messagebox.showinfo("Theme Search", f"{theme}: Found in {theme_count} comments!")

# ✅ Function: Sentiment Analysis (Fixed API Handling)
def sentiment_analysis(theme_only=False):
    global data
    if data is None:
        messagebox.showerror("Error", "Please upload a CSV first!")
        return

    comments = data[data['Theme_Match']]['Comment'].tolist() if theme_only else data['Comment'].tolist()
    sentiments = []

    try:
        for comment in comments:
            response = openai.ChatCompletion.create(
                model="gpt-4-turbo",  
                messages=[
                    {"role": "system", "content": "You are an expert sentiment classifier."},
                    {"role": "user", "content": f"Classify this comment as Positive, Negative, or Neutral: {comment}"}
                ]
            )
            # ✅ Fixed Response Handling
            sentiment = response.choices[0].message['content'].strip()
            sentiments.append(sentiment)

        if theme_only:
            data.loc[data['Theme_Match'], 'Theme_Sentiment'] = sentiments
        else:
            data['Sentiment'] = sentiments

        # Count breakdown for sentiment
        sentiment_counts = pd.Series(sentiments).value_counts().to_dict()
        messagebox.showinfo("Sentiment Analysis", f"Results:\n{sentiment_counts}")
    except Exception as e:
        messagebox.showerror("Error", f"Error during sentiment analysis: {str(e)}")

# ✅ Function: Summarize Comments using ChatCompletion (Fixed)
def summarize_comments(theme_only=False):
    global data
    if data is None:
        messagebox.showerror("Error", "Please upload a CSV first!")
        return

    comments = data[data['Theme_Match']]['Comment'].tolist() if theme_only else data['Comment'].tolist()
    combined_text = " ".join(comments)

    try:
        response = openai.ChatCompletion.create(
            model="gpt-4-turbo",
            messages=[
                {"role": "system", "content": "You are a professional text summarizer."},
                {"role": "user", "content": f"Please summarize the following comments:\n\n{combined_text}"}
            ]
        )
        # ✅ Fixed Response Handling
        summary = response.choices[0].message['content'].strip()

        if theme_only:
            data.loc[data['Theme_Match'], 'Theme_Summary'] = summary
        else:
            data['Summary'] = summary

        messagebox.showinfo("Summary", f"Summary: {summary}")
    except Exception as e:
        messagebox.showerror("Error", f"Error during summarization: {str(e)}")

# ✅ Function: Export Results (Multi-Sheet Excel)
def export_results(theme_only=False):
    global data
    if data is None:
        messagebox.showerror("Error", "No data to export!")
        return

    export_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    try:
        with pd.ExcelWriter(export_path) as writer:
            if theme_only:
                theme_data = data[data['Theme_Match']]
                theme_data.to_excel(writer, sheet_name="Theme Comments", index=False)
                sentiment_summary = theme_data['Theme_Sentiment'].value_counts().to_frame(name='Count')
                sentiment_summary.to_excel(writer, sheet_name="Theme Sentiment Breakdown")
                theme_summary = theme_data.iloc[0]['Theme_Summary'] if 'Theme_Summary' in theme_data.columns else "No Summary Available"
                pd.DataFrame({"Theme Summary": [theme_summary]}).to_excel(writer, sheet_name="Theme Summary", index=False)
            else:
                data.to_excel(writer, sheet_name="All Comments", index=False)
                sentiment_summary = data['Sentiment'].value_counts().to_frame(name='Count')
                sentiment_summary.to_excel(writer, sheet_name="Sentiment Breakdown")
                summary = data.iloc[0]['Summary'] if 'Summary' in data.columns else "No Summary Available"
                pd.DataFrame({"Overall Summary": [summary]}).to_excel(writer, sheet_name="All Comments Summary", index=False)

        messagebox.showinfo("Export Success", f"Results exported successfully to {export_path}!")
    except Exception as e:
        messagebox.showerror("Error", f"Error exporting results: {str(e)}")

# ✅ GUI Layout with Fixed API Usage
upload_button = tk.Button(app, text="Upload CSV", command=upload_csv)
upload_button.pack(pady=10)

# Theme Entry
theme_label = tk.Label(app, text="Enter Theme:")
theme_label.pack()
theme_entry = tk.Entry(app)
theme_entry.pack(pady=5)

# Keywords Entry
keywords_label = tk.Label(app, text="Enter Keywords (comma-separated):")
keywords_label.pack()
keywords_entry = tk.Entry(app)
keywords_entry.pack(pady=5)

# Buttons for Theme Operations
theme_search_button = tk.Button(app, text="Perform Theme Search", command=theme_search)
theme_search_button.pack(pady=10)

theme_sentiment_button = tk.Button(app, text="Theme Sentiment Analysis", command=lambda: sentiment_analysis(theme_only=True))
theme_sentiment_button.pack(pady=10)

theme_summary_button = tk.Button(app, text="Theme Summarization", command=lambda: summarize_comments(theme_only=True))
theme_summary_button.pack(pady=10)

export_theme_button = tk.Button(app, text="Export Theme Results", command=lambda: export_results(theme_only=True))
export_theme_button.pack(pady=10)

# Buttons for All Comments Operations
sentiment_button = tk.Button(app, text="Sentiment Analysis (All Comments)", command=lambda: sentiment_analysis(theme_only=False))
sentiment_button.pack(pady=10)

summarize_button = tk.Button(app, text="Summarize All Comments", command=lambda: summarize_comments(theme_only=False))
summarize_button.pack(pady=10)

export_button = tk.Button(app, text="Export All Results", command=lambda: export_results(theme_only=False))
export_button.pack(pady=10)

# ✅ Run the Tkinter event loop
app.mainloop()
