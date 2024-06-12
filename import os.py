import pdfplumber
import sqlite3
import tkinter as tk
from tkinter import messagebox
import os

# Function to search for tracking ID and open the PDF
def search_and_open_pdf():
    tracking_id = entry.get().strip()
    
    # Connect to SQLite database
    conn = sqlite3.connect('tracking_data.db')
    c = conn.cursor()
    
    # Search for the tracking ID
    c.execute("SELECT pdf_path FROM tracking WHERE tracking_id=?", (tracking_id,))
    result = c.fetchone()
    print(c)
    
    if result:
        pdf_path = result[0]
        if os.path.exists(pdf_path):
            os.startfile(pdf_path)
        else:
            messagebox.showerror("Error", f"File not found: {pdf_path}")
    else:
        messagebox.showerror("Error", "Tracking ID not found.")
    
    # Close the database connection
    conn.close()

# Create the main window
root = tk.Tk()
root.title("PDF Tracker")

# Create and place the label, entry, and button widgets
label = tk.Label(root, text="Enter Tracking ID:")
label.pack(pady=5)

entry = tk.Entry(root, width=50)
entry.pack(pady=5)

button = tk.Button(root, text="Search and Open PDF", command=search_and_open_pdf)
button.pack(pady=20)

# Start the GUI event loop
root.mainloop()
