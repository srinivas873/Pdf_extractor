import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, ttk
import pdfplumber
import pandas as pd
import customtkinter
from datetime import datetime
import os 


class PDFExtractorApp:
    def __init__(self, master):
        self.master = master
        master.title("PDF Data Extractor")

        # Layout
        self.label = tk.Label(master, text="PDF Data Extractor", font=("Arial", 14))
        self.label.pack(pady=10)

        self.select_button = tk.Button(master, text="Select PDF Files", command=self.select_files)
        self.select_button.pack(pady=5)

        

        self.save_button = tk.Button(master, text="Save to Excel", command=self.save_to_excel)
        self.save_button.pack(pady=5)

        self.progress_bar = customtkinter.CTkProgressBar(master, width=100, height=10)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=5)

        self.exit_button = tk.Button(master, text="Exit", command=master.quit)
        self.exit_button.pack(pady=20)
        

        self.check_files_button = tk.Button(master, text="Check Files", command=self.check_files)
        self.check_files_button.pack(pady=5)

        self.data = []

    def select_files(self):
        self.pdf_paths = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf")],
            multiple=True
        )
        messagebox.showinfo("Files Selected", f"{len(self.pdf_paths)} files selected.")

    def check_files(self):
        check_window = Toplevel(self.master)
        check_window.title("Check File Data")
        check_window.geometry("600x400") 
        
        select_button = tk.Button(check_window, text="Select PDF Files", command=lambda: self.display_file_data(check_window))
        select_button.pack(pady=10)

    def display_file_data(self, parent):
        file_paths = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf")],
            multiple=True
        )

        # Create Treeview widget
        tree = ttk.Treeview(parent, columns=('Date', 'Prepared by', 'Quote No', 'Customer Name','Prod Sales','Tooling Cost', 'Sales Rep', 'Type of Quote', 'Part Count'), show='headings')
        for col in tree['columns']:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor=tk.CENTER)
        
        tree.pack(fill='both', expand=True)

        if file_paths:
            for file_path in file_paths:
                text_lines = self.extract_text_from_pdf(file_path)
                tables = []
                with pdfplumber.open(file_path) as pdf:
                    for page in pdf.pages:
                        tables.extend(page.extract_tables())

                row = (
                    self.get_date(text_lines),
                    self.get_prepared_by(text_lines),
                    self.get_quote_no(text_lines),
                    self.get_customer_name(text_lines),
                    self.table_calculation(tables),
                    self.get_tooling_cost(tables),
                    self.get_sales_rep(self.get_customer_name(text_lines)),
                    self.get_type(tables, text_lines),
                    self.get_part_count(tables)
                )

                tree.insert("", tk.END, values=row)  # Insert data into the treeview
    def save_to_excel(self):
        if not hasattr(self, 'pdf_paths') or not self.pdf_paths:
            messagebox.showerror("Error", "No PDF files selected.")
            return

        excel_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Excel file"
        )
        if excel_file:
            self.process_data_and_save_to_excel(self.pdf_paths, excel_file)
            messagebox.showinfo("Success", "Data has been saved to Excel successfully.")

    def extract_text_from_pdf(self, file):
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    return text.split('\n')
        return []

    def get_quote_no(self, text_lines):
        quote_no = ''
        for line in text_lines:
            l = line.split(" ")
            len_of_line = len(l)
            if l[0] == "TO:" and l[-2] == "NO:":
                quote_no = l[-1]
            if len_of_line >= 4:
                if l[-3] == "Quotation":
                    quote_no = l[-1]

        return quote_no
        

    def get_date(self, text_lines):
        for line in text_lines:
            if "TO:" in line and "NO:" not in line:
                date_str = line.split()[-1].replace('‐', '-')
                date_formats = ["%m/%d/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d-%b-%Y"]
                for date_format in date_formats:
                    try:
                        return datetime.strptime(date_str, date_format).strftime("%Y-%m-%d")
                    except ValueError:
                        continue
                return date_str
        return ""

    def get_customer_name(self, text_lines):
        customer_name= ' '
        for line in text_lines:
            l = line.split(" ")
            len_of_line = len(l)
            if len_of_line >= 4:
                if l[-3] == "Quotation":
                    # Quote_no = l[-1]
                    customer_name = l[0]            
        return customer_name
            

    def get_tooling_cost(self, tables):
        Tool_Cost,Tool_cost,colum,Parts  = [],[],[],[]
        for table in tables:
            try:
                if 'Part #/Description' and 'Tool Cost' in table[0]:
                    for index, column in enumerate(table[0]):
                        colum.append(column)
                    if 'Tool Cost' in colum:
                        Tool_index = colum.index('Tool Cost')
                        parts_index = colum.index('Part #/Description')
                        if 'EAU' in colum:
                            for index, row in enumerate(table[1:]):
                                print(row)
                                Parts.append(row[parts_index])
                                Tool_cost.append(row[Tool_index])
                                    #Eau_index = colum.index('EAU')
                            for part, tool_cost in zip(set(Parts), Tool_cost):
                                if tool_cost is None or tool_cost == 'N/A' or tool_cost == 'Transfer'or tool_cost == 'Inrerchangable Insert' :
                                    continue
                                elif tool_cost.isalpha() == True:
                                    continue
                                else:
                                    Tool_Cost.append(int(tool_cost.replace('$', '').replace(',', '')))
                            #print(Tool_Cost)
                            return sum(Tool_Cost)

                        else:
                            for index, row in enumerate(table[1:]):
                                Parts.append(row[parts_index])
                                Tool_cost.append(row[Tool_index])
                            for part, tool_cost in zip(set(Parts), Tool_cost):
                                if tool_cost is None or tool_cost == 'N/A':
                                    continue
                                else:
                                    Tool_Cost.append(float(tool_cost.replace('$', '').replace(',', '')))
                            return sum(Tool_Cost)
                                #return sum(Tool_Cost)
            except(ValueError, TypeError, AttributeError, UnboundLocalError, IndexError) as e:
                #print(e,file)
                return 'error'

    def get_sales_rep(self, customer_name):
        sales_reps = {
            'ist': ['generac', 'miller', 'viginal', 'bernard', 'precision', 'plexus', 'rice', 'tregaskis', 'bemis', 'itw', 'serigraph','vignal'],
            'josco': ['nte', 'switchback', 'endogenex', 'psoup', 'nuwellis,', 'caztek', 'ametek'],
            'pioneer': ['emerson', 'hologic', 'duff', 'medica', 'pinnacle', 'team', 'ibm', 'mpr','bsm'],
            'ron owens': ['startek', 'specialized', 'control'],
            'jc sales': ['pentair']
        }
        for rep, customers in sales_reps.items():
            if customer_name.lower() in customers:
                return rep.upper()
        return "None"

    def get_prepared_by(self, text_lines):
        prepared_by = ''
        for line in text_lines:
            if "Prepared by:" in line:
                l = line.split(" ")
                len_of_line = len(l)
                if len_of_line >= 3:
                    if l[-2] == "Prepared":
                        prepared_by = l[-1]
                if len_of_line >= 4:
                    if l[-4] == "by:":
                        prepared_by = l[-3] + l[-2] + l[-1]
                    elif l[-4] == "Prepared":
                        prepared_by = l[-2] + l[-1]
                    elif l[-3] == "by:":
                        prepared_by = l[-1]
                    elif l[-3] == "Prepared":
                        prepared_by = l[-1]
                if len_of_line >= 7:
                    if l[-6] == "by:" and l[-7] == "Prepared":
                        prepared_by = l[-5] + l[-4] + l[-3] + l[-2] + l[-1]
        return prepared_by
    def get_type(self, tables, text_lines):
        

        notes1 = ""
        Type_of_quote = ''
        #print(text_lines)
        for line in text_lines:
            if "Type of quote" in line:
                l = line.split(" ")
                len_of_line = len(l)
                if len_of_line >= 3:
                    if l[-2] == 'quote':                            
                        Type_of_quote = l[-1]
                    elif l[-3] == 'quote':
                        Type_of_quote = l[-2]+l[-1]
        return Type_of_quote

    def extract_columns(self, table):
        columns = []
        for index, column in enumerate(table[0]):
            if column is not None:
                columns.append(column)
        return columns
    
    def table_calculation(self,tables):
        try:
            price_index,eau_index = 0,0
            indices,EAU,Parts,Parts1,S_eau,colum = [],[],[],[],[],[]
            for table in tables:
                        # print(table)
                if 'Part #/Description' in table[0] or 'Part Number/Desc.' in table[0] :
                    for index, column in enumerate(table[0]):
                                # print(column)
                        if column is not None:
                            colum.append(column)
                                    #print(colum)
                    if (any(a.startswith('EAU') for a in colum) == True and 'Piece Part Price\n@EAU' in colum) or 'Qty' in colum or 'QTY' in colum or (
                                    'EAU' in colum ):
                        if ('EAU' in colum and '@EAU' not in colum):
                                    # print('hi')
                            eau_index = colum.index('EAU')
                        elif 'Qty' in colum :
                            eau_index = colum.index('Qty')
                        elif 'QTY' in colum:
                            eau_index = colum.index('QTY')
                        elif 'EAU (MOQ)' in colum:
                            eau_index = colum.index('EAU (MOQ)')
                        """
                            elif 'EAU' in colum and any(
                                        a.startswith('Piece Part Price') for a in colum) == True:
                                return self.Special_eau_low(colum, file_path)"""
                    elif  'EAU' in colum and any(a.startswith('Piece Part Price\n@') for a in colum) :
                                #print('hi')
                        return self.Special_eau_low(colum, tables)
                    elif 'EAU' not in colum or ('EAU' not in colum and any(
                                        a.startswith('Piece Part Price') for a in colum) == True) or (
                                                     'EAU' not in colum and any(
                                                 a.startswith('UNIT PRICE') for a in colum) == True) or ('EAU' in colum and any(
                                        a.startswith('Piece Part Price') for a in colum) == True):
                        return self.Special_eau_low(colum, tables)
                    elif ('EAU' in colum and 'Piece Part Price\n@EAU' not in colum) and (
                                            'EAU' in colum and 'Piece Part\nPrice @EAU' not in colum) and any(
                                    a.startswith('Piece Part Price') for a in colum) == False:
                        return self.Special_eau2_low(eau_index, colum, tables)

                    if f'Piece Part Price\n@EAU' in colum:
                        price_index = colum.index('Piece Part Price\n@EAU')
                                # print(price_index)
                    elif f'Piece Part Price\n@EAU' not in colum:
                        price_index = eau_index + 1
                                # print(price_index)
                    try:

                        for index, row in enumerate(table[1:]):
                            len_of_row = index
                            if (row[0] != '0') or (row[0] != ''):
                                len_of_row = index
                                Parts.append(row[0])

                                if row[eau_index] == None or row[eau_index] == '':
                                    eau = 0
                                else:
                                    eau = int(row[eau_index].replace(',', ''))

                                EAU.append(eau)

                                if row[price_index] == None or row[price_index] == '':
                                    piece_part_price = 0
                                else:
                                    piece_part_price = float(row[price_index].replace('$', '').replace(',', ''))

                                Parts1.append(piece_part_price)

                                if row[0] is not None:
                                    indices.append(index)
                        if index + 1 != len(indices):
                            return self.Indices_not_same_low(EAU, Parts, Parts1, indices, len_of_row)
                        if len(Parts) == len(list(set(Parts))):
                            return self.Table1_low(EAU, Parts1)

                        elif len(Parts) != len(list(set(Parts))):
                            Parts2 = list(set(Parts))
                            return self.Table2_low(Parts2, Parts, EAU, Parts1)

                    except (ValueError, TypeError, AttributeError, UnboundLocalError, IndexError) as e:
                        print(e)

        except:
            return 'Error'
    def Special_eau2_low(self,eau_index, colum, tables):
        Part_price1 = []
        eau_index = colum.index('EAU')
        if f'Piece Part Price\n@EAU' in colum:
            price_index = colum.index('Piece Part Price\n@EAU')
            # print(price_index,"hi")
        elif f'Piece Part Price\n@EAU' not in colum:
            price_index = eau_index + 1
            # print(price_index,"bye")
        # price_index = eau_index + 1
        indices = []
        EAU = []
        Parts = []
        Parts1 = []
        S_eau = []
        colum = []
        Part_price = []
        
        for table in tables:
                    # print(table)
            if 'Part #/Description' in table[0]:
                for index, column in enumerate(table[0]):
                            # print(column)
                    if column is not None:
                        colum.append(column)
                for index, row in enumerate(table[1:]):
                            # print(row[eau_index])
                    len_of_row = index
                    if (row[0] != '0') or (row[0] != ''):
                        len_of_row = index
                        Parts.append(row[0])
                    if row[eau_index] == None or row[eau_index] == '':
                        eau = 0
                    else:
                        eau = int(row[eau_index].replace(',', ''))
                    EAU.append(eau)
                    if row[price_index] == None or row[price_index] == '':
                        piece_part_price = 0
                    else:
                        piece_part_price = float(row[price_index].replace('$', '').replace(',', ''))
                    Parts1.append(piece_part_price)
                    if row[0] is not None:
                        indices.append(index)
                if index + 1 != len(indices):
                    return self.Indices_not_same(EAU, Parts, Parts1, indices, len_of_row)
                else:
                    if 'Part #/Description' in table[0]:
                        if 'EAU' in colum:
                            eau_index = colum.index('EAU')
                        for index, row in enumerate(table[1:]):
                            eau = int(row[eau_index].replace(',', ''))
                            EAU.append(eau)
                        for index, i in enumerate(colum):
                            if f'@{eau}' in i.replace(',', '').replace(' ', ''):
                                price_index = index
                            else:
                                price_index = eau_index + 1
                        for index, row in enumerate(table[1:]):
                            Part_price.append(float(row[price_index].replace('$', '')))
                        for parts, price in zip(set(Parts), Part_price):
                            Part_price1.append(price)
                total_sum = sum(num * eau for num in Part_price1)
                return total_sum
    def Table1_low(self,EAU, Parts1):
        try:
            total_sum = 0
            result = sum(a * b for a, b in zip(EAU, Parts1))
            total_sum = result
        except (ValueError, TypeError) as e:
            return e
        return total_sum

    def Table2_low(self,Parts2, Parts, EAU, Parts1):
        EAU1 = []
        Parts11 = []
        part_eau_price_dict = {}
        if len(Parts2) != len(Parts):
            for part, eau, price in zip(Parts, EAU, Parts1):
                if part in part_eau_price_dict:
                    # Check if the current EAU is higher than what's stored
                    if eau < part_eau_price_dict[part][0]:
                        part_eau_price_dict[part] = (eau, price)
                else:
                    # Store the EAU and price for the part
                    part_eau_price_dict[part] = (eau, price)
            total_sum = sum(eau * price for eau, price in part_eau_price_dict.values())
            return total_sum
        else:
            for part, eau, price in zip(Parts, EAU, Parts1):
                if part in part_eau_price_dict:
                    if eau < part_eau_price_dict[part][0]:  # Check if the current EAU is less than the stored EAU
                        part_eau_price_dict[part] = (eau, price)
                else:
                    part_eau_price_dict[part] = (eau, price)
                # print(part_eau_price_dict[part])
            total_sum = sum(eau * price for eau, price in part_eau_price_dict.values())
            return total_sum

    def Special_eau_low(self,colum, tables):
        # print("hi")
        S_eau = []
        EAU = 0
        Part_price = []
        
        
        for table in tables:
                    # print(table[0])
                    if 'Part #/Description' in table[0]:
                        for index, column in enumerate(table[0]):
                            # print(column)
                            if column is not None:
                                if 'Piece Part Price' in column or 'UNIT PRICE' in column:
                                    # print(column.replace('Piece Part Price', '').replace('@', '').replace(',','').replace('Pcs','').strip())
                                    S_eau.append(
                                        int(column.replace('Piece Part Price', '').replace('UNIT PRICE', '').replace(
                                            '@', '').replace(',',
                                                             '').replace(
                                            'Pcs', '').replace('pcs', '').replace('(MOQ)', '').replace('EAU',
                                                                                                       '').strip()))
                                elif 'Part Price ' in column:
                                    # print('hi')
                                    S_eau.append(
                                        int(column.replace('Part Price', '').replace('@', '').replace(',', '').replace(
                                            'Pcs', '').replace('pcs', '').strip()))
                        EAU = min(S_eau)
                        for index, i in enumerate(colum):
                            if str(EAU) == i.replace('Piece Part Price', '').replace('UNIT PRICE', '').replace('@','').replace(',','').replace('Pcs', '').replace('pcs', '').replace('(MOQ)', '').replace('EAU', '').strip()or str(EAU) == i.replace(
                    'Part Price', '').replace('@', '').replace(',', '').replace(
                    'Pcs', '').replace('pcs', '').strip():
                                price_index = index

        for index, row in enumerate(table[1:]):
            Part_price.append(float(row[price_index].replace('$', '')))
        total_sum = sum(num * EAU for num in Part_price)
        return total_sum
    def Indices_not_same_low(self,EAU, Parts, Parts1, indices, len_of_row):
        Actual_eau = []
        Actual_part_price = []
        for i in range(len(indices)):
            Actual_eau.append(EAU[indices[i]])
            Actual_part_price.append(Parts1[indices[i]])
 
        total_sum = sum(a * b for a, b in zip(Actual_eau, Actual_part_price))
        #print(total_sum)
        return total_sum

    def get_part_count(self, tables):
        parts = []
        for table in tables:
            if 'Part #/Description' in table[0] or "Part Number/Desc." in table[0]:
                for row in table[1:]:
                    if row[0]:
                        parts.append(row[0])
        return len(set(parts))

    def is_valid_pdf(self, file_path):
        try:
            with pdfplumber.open(file_path) as pdf:
                return True
        except:
            return False

    def process_data_and_save_to_excel(self, pdf_paths, excel_path):
        total_files = len(pdf_paths)
        processed_files = 0

        for pdf_path in pdf_paths:
            filename = os.path.basename(pdf_path)  # Get only the file name, not the full path
            if not self.is_valid_pdf(pdf_path):
                continue

            text_lines = self.extract_text_from_pdf(pdf_path)
            tables = []
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    tables.extend(page.extract_tables())

            self.data.append((
                self.get_date(text_lines),
                self.get_prepared_by(text_lines),
                self.get_customer_name(text_lines),
                self.get_type(tables, text_lines),
                self.get_quote_no(text_lines),
                self.get_tooling_cost(tables),
                self.table_calculation(tables),
                self.get_sales_rep(self.get_customer_name(text_lines)),
                self.get_part_count(tables),
                filename.replace('.pdf','')  # Use the extracted filename instead of the full path
            ))

            processed_files += 1
            progress = processed_files / total_files
            self.progress_bar.set(progress)
            self.master.update_idletasks()

        df = pd.DataFrame(self.data, columns=['Date', "Prepared by", "Customer_name", 'Type of Quote','Quote no', 'Tooling Cost', 'Low EAU', 'Sale rep', 'Count of Parts','path'])
        df.to_excel(excel_path, index=False)
        self.data = []

if __name__ == "__main__":
    root = tk.Tk()
    root.minsize(600, 350)
    app = PDFExtractorApp(root)
    root.mainloop()
