import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import customtkinter
import os
import time


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
        self.progress_bar.set(0)  # Initialize progress bar to 0
        self.progress_bar.pack(pady=5)


        self.exit_button = tk.Button(master, text="Exit", command=master.quit)
        self.exit_button.pack(pady=20)

        self.data = []

    def select_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf")],
            multiple=True
        )
        self.pdf_paths = file_paths
        messagebox.showinfo("Files Selected", f"{len(file_paths)} files selected.")

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

    data = []
    def Quote_no(self,file):
        quote_no = ""
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                # print(page)
                text = page.extract_text()
                if text:
                    lines = text.split('\n')

                    for line in lines:
                        l = line.split(" ")
                        len_of_line = len(l)
                        if l[0] == "TO:" and l[-2] == "NO:":
                            quote_no = l[-1]
                        if len_of_line >= 4:
                            if l[-3] == "Quotation":
                                quote_no = l[-1]

                return quote_no

    def date(self,file_path):
        Date = ''
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()

                if text:
                    lines = text.split('\n')
                    for line in lines:
                        l = line.split(" ")
                        len_of_line = len(l)
                        if l[0] == "TO:" and l[-2] != "NO:":
                            Date = l[-1]

        return Date

    def Customer_name(self,file):
        customer_name = ""
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                # print(page)
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    for line in lines:
                        l = line.split(" ")
                        len_of_line = len(l)
                        if len_of_line >= 4:
                            if l[-3] == "Quotation":
                                # Quote_no = l[-1]
                                customer_name = l[0]

                return customer_name

    def tooling_cost(self,file):
        # Date = date(file)
        Parts = []
        Tool_index = 0
        Tool_cost = []
        Tool_Cost = []
        colum = []
        EAU = []
        Part_price = []
        part_eau_price_dict = {}
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
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
                                        Parts.append(row[parts_index])
                                        Tool_cost.append(row[Tool_index])
                                    #Eau_index = colum.index('EAU')
                                    for part, tool_cost in zip(set(Parts), Tool_cost):
                                        if tool_cost is None or tool_cost == 'N/A' or tool_cost == 'Transfer'or tool_cost == 'Inrerchangable Insert' :
                                            continue
                                        else:
                                            Tool_Cost.append(int(tool_cost.replace('$', '').replace(',', '')))
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
                                print(e,file)
                                return 'error'

    def sales_rep(self,file):
        sale_rep = ''
        customer_name = self.Customer_name(file)
        with pdfplumber.open(file) as pdf:
            if customer_name.lower() in ['generac', 'miller', 'viginal', 'bernard', 'precision', 'plexus', 'rice',
                                         'tregaskis', 'bemis', 'itw', 'serigraph']:
                sale_rep = 'IST'
            elif customer_name.lower() in ['nte', 'switchback', 'endogenex', 'psoup', 'nuwellis,', 'caztek', 'ametek']:
                sale_rep = 'JOSCO'
            elif customer_name.lower() in ['emerson', 'hologic', 'duff', 'medica', 'pinnacle', 'team', 'ibm', 'mpr']:
                sale_rep = 'PIONEER'
            elif customer_name.lower() in ['startek', 'specialized', 'control']:
                sale_rep = 'RON OWENS'
            elif customer_name.lower() in ['jc sales']:
                sale_rep = 'PENTAIR'
            else:
                sale_rep = 'None'
        return sale_rep
    def Prepared_by(self,file):
        prepared_by = None  # Initialize the variable to ensure it has a default value
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                # print(page)
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    for line in lines:
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

    def table_calculation(self,file_path):
        try:
            price_index = 0
            eau_index = 0
            indices = []
            EAU = []
            Parts = []
            Parts1 = []
            S_eau = []
            colum = []
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        # print(table)
                        if 'Part #/Description' in table[0]:
                            for index, column in enumerate(table[0]):
                                # print(column)
                                if column is not None:
                                    colum.append(column)
                                    #print(colum)
                            if ('EAU' in colum and 'Piece Part Price\n@EAU' in colum) or 'Qty' in colum or (
                                    'EAU' in colum and any(
                                    a.startswith('Piece Part Price @ EAU') for a in colum) == True) or (
                                    'EAU' in colum and any(
                                    a.startswith('Piece Part Price') for a in colum) == True) or (
                                    'EAU (MOQ)' in colum and any(
                                a.startswith('Piece Part Price') for a in colum) == True):
                                if ('EAU' in colum and '@EAU' not in colum):
                                    # print('hi')
                                    eau_index = colum.index('EAU')
                                elif 'Qty' in colum:
                                    eau_index = colum.index('Qty')
                                elif 'EAU (MOQ)' in colum:
                                    eau_index = colum.index('EAU (MOQ)')
                            elif 'EAU' not in colum or ('EAU' in colum and any(
                                        a.startswith('Piece Part Price') for a in colum) == True) or (
                                                     'EAU' not in colum and any(
                                                 a.startswith('UNIT PRICE') for a in colum) == True):
                                return self.Special_eau_low(colum, file_path)
                            elif ('EAU' in colum and 'Piece Part Price\n@EAU' not in colum) and (
                                            'EAU' in colum and 'Piece Part\nPrice @EAU' not in colum) and any(
                                    a.startswith('Piece Part Price') for a in colum) == False:
                                return self.Special_eau2_low(eau_index, colum, file_path)

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
                                print(e,file_path)

        except:
            return 'Error'
    def Special_eau2_low(self,eau_index, colum, file_path):
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
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
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

    def Special_eau_low(self,colum, file_path):
        # print("hi")
        S_eau = []
        EAU = 0
        Part_price = []
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
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
        if len(indices) == 1:
            for i in range(len(indices)):
                Actual_eau.append(EAU[indices[i]])
                Actual_part_price.append(Parts1[indices[i]])
        if len(indices) == 2:
            for i in range(len(indices)):
                Actual_eau.append(EAU[indices[i]])
                Actual_part_price.append(Parts1[indices[i]])
        if len(indices) == 3:
            for i in range(len(indices)):
                Actual_eau.append(EAU[indices[i]])
                Actual_part_price.append(Parts1[indices[i]])
        elif len(indices) == 4:
            for i in range(len(indices)):
                Actual_eau.append(EAU[indices[i]])
                Actual_part_price.append(Parts1[indices[i]])
        elif len(indices) == 5:
            for i in range(len(indices)):
                Actual_eau.append(EAU[indices[i]])
                Actual_part_price.append(Parts1[indices[i]])
        elif len(indices) == 6:
            for i in range(len(indices)):
                Actual_eau.append(EAU[indices[i]])
                Actual_part_price.append(Parts1[indices[i]])
        elif len(indices) == 7:
            for i in range(len(indices)):
                Actual_eau.append(EAU[indices[i]])
                Actual_part_price.append(Parts1[indices[i]])
        elif len(indices) == 8:
            for i in range(len(indices)):
                Actual_eau.append(EAU[indices[i]])
                Actual_part_price.append(Parts1[indices[i]])
        total_sum = sum(a * b for a, b in zip(Actual_eau, Actual_part_price))
        return total_sum
    def table_calculation_High_EAU(self,file_path):
        try:
            price_index = 0
            eau_index = 0
            indices = []
            EAU = []
            Parts = []
            Parts1 = []
            S_eau = []
            colum = []
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        # print(table)
                        if 'Part #/Description'  in table[0]:
                            for index, column in enumerate(table[0]):
                                if column is not None:
                                    colum.append(column)
                            if ('EAU' in colum and 'Piece Part Price\n@EAU' in colum) or 'Qty' in colum or (
                                    'EAU' in colum and any(
                                a.startswith('Piece Part Price @ EAU') for a in colum) == True) or (
                                    'EAU' in colum and any(
                                    a.startswith('Piece Part Price') for a in colum) == True) or (
                                    'EAU (MOQ)' in colum and any(
                                a.startswith('Piece Part Price') for a in colum) == True):
                                if ('EAU' in colum and '@EAU' not in colum):
                                    # print('hi')
                                    eau_index = colum.index('EAU')
                                elif 'Qty' in colum:
                                    eau_index = colum.index('Qty')
                                elif 'EAU (MOQ)' in colum:
                                    eau_index = colum.index('EAU (MOQ)')
                            elif 'EAU' not in colum or ('EAU' in colum and any(
                                        a.startswith('Piece Part Price') for a in colum) == True) or (
                                                 'EAU' not in colum and any(
                                             a.startswith('UNIT PRICE') for a in colum) == True):
                                return self.Special_eau(colum, file_path)
                            elif ('EAU' in colum and 'Piece Part Price\n@EAU' not in colum) and (
                                            'EAU' in colum and 'Piece Part\nPrice @EAU' not in colum) and any(
                                    a.startswith('Piece Part Price') for a in colum) == False:
                                return self.Special_eau2(eau_index, colum, file_path)
                            if f'Piece Part Price\n@EAU' in colum:
                                price_index = colum.index('Piece Part Price\n@EAU')
                            elif f'Piece Part Price\n@EAU' not in colum:
                                price_index = eau_index + 1
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
                                    return self.Indices_not_same(EAU, Parts, Parts1, indices, len_of_row)
                                if len(Parts) == len(list(set(Parts))):
                                    return self.Table1(EAU, Parts1)
                                elif len(Parts) != len(list(set(Parts))):
                                    Parts2 = list(set(Parts))
                                    return self.Table2(Parts2, Parts, EAU, Parts1)
                            except (ValueError, TypeError, AttributeError, UnboundLocalError, IndexError) as e:
                                print(e,file_path)
        except:
            return 'error'
    def Special_eau2(self,eau_index, colum, file_path):
        Part_price1 = []
        eau_index = colum.index('EAU')
        if f'Piece Part Price\n@EAU' in colum:
            price_index = colum.index('Piece Part Price\n@EAU')
        elif f'Piece Part Price\n@EAU' not in colum:
            price_index = eau_index + 1
        indices = []
        EAU = []
        Parts = []
        Parts1 = []
        S_eau = []
        colum = []
        Part_price = []
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if 'Part #/Description' in table[0]:
                        for index, column in enumerate(table[0]):
                            # print(column)
                            if column is not None:
                                colum.append(column)
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
    def Table1(self,EAU, Parts1):
        try:
            total_sum = 0
            result = sum(a * b for a, b in zip(EAU, Parts1))
            total_sum = result
        except (ValueError, TypeError) as e:
            return e
        return total_sum
    def Table2(self,Parts2, Parts, EAU, Parts1):
        EAU1 = []
        Parts11 = []
        part_eau_price_dict = {}
        if len(Parts2) != len(Parts):
            for part, eau, price in zip(Parts, EAU, Parts1):
                if part in part_eau_price_dict:
                    # Check if the current EAU is higher than what's stored
                    if eau > part_eau_price_dict[part][0]:
                        part_eau_price_dict[part] = (eau, price)
                else:
                    # Store the EAU and price for the part
                    part_eau_price_dict[part] = (eau, price)
            total_sum = sum(eau * price for eau, price in part_eau_price_dict.values())
            return total_sum
        else:
            for part, eau, price in zip(Parts, EAU, Parts1):
                if part in part_eau_price_dict:
                    if eau > part_eau_price_dict[part][0]:  # Check if the current EAU is less than the stored EAU
                        part_eau_price_dict[part] = (eau, price)
                else:
                    part_eau_price_dict[part] = (eau, price)
            total_sum = sum(eau * price for eau, price in part_eau_price_dict.values())
            return total_sum

    def Special_eau(self,colum, file_path):
        S_eau = []
        EAU = 0
        Part_price = []
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if 'Part #/Description' in table[0]:
                        for index, column in enumerate(table[0]):
                            if column is not None:
                                if 'Piece Part Price' in column or 'UNIT PRICE' in column:
                                    S_eau.append(int(column.replace('Piece Part Price', '').replace('UNIT PRICE', '').replace('@', '').replace(',','').replace('Pcs', '').replace('pcs', '').replace('(MOQ)', '').replace('EAU',
                                                                                                       '').strip()))
                                elif 'Part Price ' in column:
                                    S_eau.append(
                                        int(column.replace('Part Price', '').replace('@', '').replace(',', '').replace(
                                            'Pcs', '').replace('pcs', '').strip()))
            EAU = max(S_eau)
            for index, i in enumerate(colum):
                if str(EAU) == i.replace('Piece Part Price', '').replace('UNIT PRICE', '').replace('@','').replace(',','').replace('Pcs', '').replace('pcs', '').replace('(MOQ)', '').replace('EAU', '').strip()or str(EAU) == i.replace(
                    'Part Price', '').replace('@', '').replace(',', '').replace(
                    'Pcs', '').replace('pcs', '').strip():
                    price_index = index
            for index, row in enumerate(table[1:]):
                Part_price.append(float(row[price_index].replace('$', '')))
            total_sum = sum(num * EAU for num in Part_price)
        return total_sum
    def Indices_not_same(self,EAU, Parts, Parts1, indices, len_of_row):
        Actual_eau = []
        Actual_part_price = []
        if len(indices) == 1:
            Actual_eau.append(EAU[len_of_row])
            Actual_part_price.append(Parts1[len_of_row])
        if len(indices) == 2:
            Actual_eau.append(EAU[indices[1] - 1])
            Actual_eau.append(EAU[len_of_row])
            Actual_part_price.append(Parts1[indices[1] - 1])
            Actual_part_price.append(Parts1[len_of_row])
        if len(indices) == 3:
            Actual_eau.append(EAU[indices[1] - 1])
            Actual_eau.append(EAU[indices[2] - 1])
            Actual_eau.append(EAU[len_of_row])
            Actual_part_price.append(Parts1[indices[1] - 1])
            Actual_part_price.append(Parts1[indices[2] - 1])
            Actual_part_price.append(Parts1[len_of_row])
        elif len(indices) == 4:
            Actual_eau.append(EAU[indices[1] - 1])
            Actual_eau.append(EAU[indices[2] - 1])
            Actual_eau.append(EAU[indices[3] - 1])
            Actual_eau.append(EAU[len_of_row])
            Actual_part_price.append(Parts1[indices[1] - 1])
            Actual_part_price.append(Parts1[indices[2] - 1])
            Actual_part_price.append(Parts1[indices[3] - 1])
            Actual_part_price.append(Parts1[len_of_row])
        elif len(indices) == 5:
            Actual_eau.append(EAU[indices[1] - 1])
            Actual_eau.append(EAU[indices[2] - 1])
            Actual_eau.append(EAU[indices[3] - 1])
            Actual_eau.append(EAU[indices[4] - 1])
            Actual_eau.append(EAU[len_of_row])
            Actual_part_price.append(Parts1[indices[1] - 1])
            Actual_part_price.append(Parts1[indices[2] - 1])
            Actual_part_price.append(Parts1[indices[3] - 1])
            Actual_part_price.append(Parts1[indices[4] - 1])
            Actual_part_price.append(Parts1[len_of_row] - 1)
        elif len(indices) == 6:
            Actual_eau.append(EAU[indices[1] - 1])
            Actual_eau.append(EAU[indices[2] - 1])
            Actual_eau.append(EAU[indices[3] - 1])
            Actual_eau.append(EAU[indices[4] - 1])
            Actual_eau.append(EAU[indices[5] - 1])
            Actual_eau.append(EAU[len_of_row])
            Actual_part_price.append(Parts1[indices[1] - 1])
            Actual_part_price.append(Parts1[indices[2] - 1])
            Actual_part_price.append(Parts1[indices[3] - 1])
            Actual_part_price.append(Parts1[indices[4] - 1])
            Actual_part_price.append(Parts1[indices[5] - 1])
            Actual_part_price.append(Parts1[len_of_row])
        elif len(indices) == 7:
            Actual_eau.append(EAU[indices[1] - 1])
            Actual_eau.append(EAU[indices[2] - 1])
            Actual_eau.append(EAU[indices[3] - 1])
            Actual_eau.append(EAU[indices[4] - 1])
            Actual_eau.append(EAU[indices[5] - 1])
            Actual_eau.append(EAU[indices[6] - 1])
            Actual_eau.append(EAU[len_of_row])
            Actual_part_price.append(Parts1[indices[1] - 1])
            Actual_part_price.append(Parts1[indices[2] - 1])
            Actual_part_price.append(Parts1[indices[3] - 1])
            Actual_part_price.append(Parts1[indices[4] - 1])
            Actual_part_price.append(Parts1[indices[5] - 1])
            Actual_part_price.append(Parts1[indices[6] - 1])
            Actual_part_price.append(Parts1[len_of_row])
        elif len(indices) == 8:
            Actual_eau.append(EAU[indices[1] - 1])
            Actual_eau.append(EAU[indices[2] - 1])
            Actual_eau.append(EAU[indices[3] - 1])
            Actual_eau.append(EAU[indices[4] - 1])
            Actual_eau.append(EAU[indices[5] - 1])
            Actual_eau.append(EAU[indices[6] - 1])
            Actual_eau.append(EAU[indices[7] - 1])
            Actual_eau.append(EAU[len_of_row])
            Actual_part_price.append(Parts1[indices[1] - 1])
            Actual_part_price.append(Parts1[indices[2] - 1])
            Actual_part_price.append(Parts1[indices[3] - 1])
            Actual_part_price.append(Parts1[indices[4] - 1])
            Actual_part_price.append(Parts1[indices[5] - 1])
            Actual_part_price.append(Parts1[indices[6] - 1])
            Actual_part_price.append(Parts1[indices[7] - 1])
            Actual_part_price.append(Parts1[len_of_row])
        total_sum = sum(a * b for a, b in zip(Actual_eau, Actual_part_price))
        return total_sum
    

    def File_details(self,pdf_path):
        Modified = time.ctime(os.path.getmtime(pdf_path))
        return Modified

    def process_data_and_save_to_excel(self, pdf_paths, excel_path):
        total_files = len(pdf_paths)
        processed_files = 0
        for pdf_path in pdf_paths:
            self.data.append((self.date(pdf_path), self.Prepared_by(pdf_path), self.Customer_name(pdf_path),
                              self.Quote_no(pdf_path), self.tooling_cost(pdf_path), self.table_calculation(pdf_path),self.table_calculation_High_EAU(pdf_path),
                              self.sales_rep(pdf_path),self.File_details(pdf_path)))
            processed_files += 1
            progress = processed_files / total_files
            self.progress_bar.set(progress)
            self.master.update_idletasks()  # Update progress bar on screen
        df = pd.DataFrame(self.data, columns=['Date', "Prepared by", "Customer_name", 'Quote no', 'Tooling Cost', 'Low EAU','High_EAU', 'Sale rep','Modified'])
        df.to_excel(excel_path, index=False)
        self.data = []  # Clear data after saving

if __name__ == "__main__":
    root = tk.Tk()
    root.minsize(600,350)
    app = PDFExtractorApp(root)
    root.mainloop()
