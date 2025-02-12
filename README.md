# PDF Data Extractor

**A Python application that extracts relevant information from PDF files and saves the data to an Excel sheet for easy data analysis.**

## Features

✅ Select multiple PDF files for data extraction  
✅ Extracts relevant fields such as Date, Prepared By, Quote Number, Customer Name, Sales Data, and more  
✅ Displays extracted data in a GUI before saving  
✅ Saves extracted data in an Excel (.xlsx) file  
✅ Progress bar to show processing status  
✅ User-friendly GUI with **Tkinter**  

## Technologies Used

- Python  
- Tkinter  
- PDFPlumber (for PDF text extraction)  
- Pandas (for data storage and processing)  
- CustomTkinter (for enhanced UI elements)  

## Installation Guide

### 1️⃣ Clone the Repository
```bash
git clone https://github.com/your-github-username/pdf-data-extractor.git
cd pdf-data-extractor
```

### 2️⃣ Install Dependencies
Ensure you have Python **3.7+** installed. Then, install the required dependencies using:  
```bash
pip install -r requirements.txt
```
  
### 3️⃣ Run the Application
```bash
python pdf_extractor.py
```

## How to Use the Application

### Step 1: Select PDF Files
Click the **"Select PDF Files"** button to choose the files you want to extract data from.  

### Step 2: View Extracted Data
Click **"Check Files"** to preview the extracted data in a table format.  

### Step 3: Save Data to Excel
Click **"Save to Excel"**, choose a save location, and the extracted data will be stored in an Excel (.xlsx) file.  

### Step 4: Exit the Application
Click **"Exit"** to close the program when done.  

## Folder Structure
```
📂 pdf-data-extractor  
 ┣ 📜 pdf_extractor.py  # Main application script  
 ┣ 📜 requirements.txt  # Dependencies  
 ┣ 📜 README.md         # Project Documentation  
 ┗ 📜 .gitignore        # Git Ignore File  
```

## Contributing
Feel free to fork this repository and contribute! If you find any bugs or have feature requests, open an issue.  

## License
This project is licensed under the MIT License.

