import pyautogui
import pyperclip
import speech_recognition as sr
import tkinter as tk
from tkinter import messagebox, ttk
from docx import Document
from datetime import datetime

# Language codes dictionary
LANGUAGE_CODES = {
    "Hindi": "hi-IN",
    "Tamil": "ta-IN",
    "Telugu": "te-IN",
    "Bengali": "bn-IN",
    "Gujarati": "gu-IN",
    "Kannada": "kn-IN",
    "Malayalam": "ml-IN",
    "Marathi": "mr-IN",
    "Punjabi": "pa-IN",
    "Urdu": "ur-IN"
}

# Function to recognize speech and save to docx and txt files
def recognize_speech():
    recognizer = sr.Recognizer()
    selected_language = language_var.get()
    language_code = LANGUAGE_CODES.get(selected_language, "en-IN")  # Default to English (India)
    
    # Use the default microphone as the audio source
    with sr.Microphone() as source:
        label_status.config(text=f"Listening in {selected_language}...")
        audio_data = recognizer.listen(source)
        label_status.config(text="Processing...")

    try:
        # Convert speech to text in the selected language
        text = recognizer.recognize_google(audio_data, language=language_code)
        label_text.config(text=text)
        
        # Copy the recognized text to clipboard
        pyperclip.copy(text)
        label_status.config(text="Text copied to clipboard!")

        # Save text to .txt file
        with open("transcript.txt", "a", encoding="utf-8") as txt_file:
            txt_file.write(text + "\n")

        # Save text to a uniquely named .docx file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        doc_filename = f"transcript_{timestamp}.docx"
        doc = Document()
        doc.add_paragraph(text)
        doc.save(doc_filename)

        label_status.config(text=f"Text saved to {doc_filename} and transcript.txt")
        
    except sr.UnknownValueError:
        messagebox.showerror("Error", "Could not understand the audio")
        label_status.config(text="Error: Could not understand audio")
    except sr.RequestError:
        messagebox.showerror("Error", "Speech Recognition service error")
        label_status.config(text="Error: API request failed")

# GUI setup
root = tk.Tk()
root.title("Speech to Text App")
root.geometry("400x400")

# Language Selection
language_var = tk.StringVar(value="English")  # Default language
label_language = tk.Label(root, text="Select Language:")
label_language.pack(pady=5)
language_dropdown = ttk.Combobox(root, textvariable=language_var, values=list(LANGUAGE_CODES.keys()))
language_dropdown.pack(pady=5)

# GUI Elements
label_instructions = tk.Label(root, text="Press 'Start' and speak into your microphone")
label_instructions.pack(pady=10)

label_text = tk.Label(root, text="", wraplength=300, font=("Arial", 12))
label_text.pack(pady=10)

label_status = tk.Label(root, text="", fg="blue")
label_status.pack(pady=10)

button_start = tk.Button(root, text="Start", command=recognize_speech, font=("Arial", 14))
button_start.pack(pady=20)

# Run the application
root.mainloop()
