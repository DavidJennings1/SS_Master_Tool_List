'''Parse files in choosen folder and create spreadsheet containing
cutting tool usage data and file count by programmer.
Note - Toolist file location is hard coded.'''

import os
import re
from collections import Counter
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog  # noqa: F401
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, Font, NamedStyle


class Parse_Files:
    ''' Class to parse through files in specified folder, search
    pattern, and return match'''

    # Class variables

    new_dict = {}
    single_list = {}
    result_dict = {}
    machine = ''

    def __init__(self, folder, pattern):
        self.folder = folder
        self.pattern = pattern
        os.chdir(folder)
        files = os.listdir()
        for item in files:
            if os.path.isdir(item):
                continue
            bin_file = Parse_Files.is_binary(item)
            if bin_file:
                continue
            with open(item, 'r') as f:
                file_contents = f.read()
                match = pattern.findall(file_contents)
                Parse_Files.result_dict[item] = set(match)
        Parse_Files.usage_count(Parse_Files.result_dict)

    def is_binary(file_name):
        ''' Fuction tries to open file as test and returns boolean'''
        try:
            with open(file_name, 'tr') as check_file:
                check_file.read()
                return False
        except UnicodeDecodeError:
            return True

    def usage_count(parsed_data):
        ''' Returns dictionary with tool number as key and
        number of times used as value'''
        all_file_tool_list = []
        for key, value in parsed_data.items():
            for item in value:
                all_file_tool_list.append(item)
        x = (s.strip('T') for s in all_file_tool_list)
        y = (s.replace('T', '') for s in x)
        new_tool_list = []
        for item in y:
            if item == '0' or item == '239' or int(item) > 300:
                continue
            new_tool_list.append(int(item))
        new_tool_list.sort()
        temp_dict = Counter(new_tool_list)
        for k, v in temp_dict.items():
            Parse_Files.new_dict[k] = v
        Parse_Files.single_use(Parse_Files.new_dict)

    def single_use(new_dict):
        '''Returns dictionary with tool number as key and
        file name as value for tools used in only one file'''
        single = []
        for keys, values in Parse_Files.new_dict.items():
            if values == 1:
                single.append(keys)
        for tnum in single:
            for k, v in Parse_Files.result_dict.items():
                for i in v:
                    if i == r'T{}'.format(tnum):
                        Parse_Files.single_list[tnum] = k

    def get_ct_number(in_data):
        '''Gets CT number from master tool lis file.'''
        Parse_Files.machine = choose_machine_combo.get()
        tl = 'C:/Users/djennings/Google Drive/King Machine Cutting Tool List.xlsx'
        wb = load_workbook(filename=tl)
        sh1 = wb.active
        t_data = {}
        for key in in_data:
            for cell1, cell2, cell3, cell4 in zip(sh1['A'], sh1['E'],
                                                  sh1['C'], sh1['J']):
                if cell1.value == key and cell2.value == Parse_Files.machine:
                    t_data[key] = (cell3.value, cell4.value)
        return t_data

    def extract_programmer():
        '''Retrieves programmer stats from files in folder'''
        os.chdir(Parse_Files.folder_selected)
        files = os.listdir()
        dave_count = 0
        john_count = 0
        no_name = 0
        for item in files:
            if os.path.isdir(item):
                continue
            bin_file = Parse_Files.is_binary(item)
            if bin_file:
                continue
            with open(item, 'r') as f:
                file_data = f.read()
                if 'DAVE' in file_data:
                    dave_count += 1
                elif 'JOHN' in file_data:
                    john_count += 1
                else:
                    no_name += 1
        dave_percent = (dave_count / (dave_count + john_count + no_name) * 100)
        john_percent = (john_count / (dave_count + john_count + no_name) * 100)
        no_name_percent = (no_name / (dave_count + john_count + no_name) * 100)
        return (dave_percent, john_percent, no_name_percent,
                dave_count, john_count, no_name)


root = tk.Tk()
root.title('Extract Tool List From Machine Library')


def choose_folder(event):
    '''Opens folder selection dialog'''
    Parse_Files.folder_selected = tk.filedialog.askdirectory()
    file_listbox.insert(tk.END, Parse_Files.folder_selected)
    folder_pick.unbind('<ButtonRelease-1>')
    process.bind('<ButtonRelease-1>', extract)


def extract(event):
    '''Used to kick off processing after button selected.'''
    pattern = re.compile(r'T\d+')
    Parse_Files(Parse_Files.folder_selected, pattern)
    write_to_spreadsheet()


def max_length(eval_string):
    '''Function takes dictionary and returns length of longest key or value'''
    string_length = 0
    for keys, values in eval_string.items():
        if len(str(keys)) > len(str(values)) and (len(str(keys)) >
                                                  string_length):
            string_length = len(str(keys))
        elif len(str(values)) > string_length:
            string_length = len(str(values))
    return(string_length)

# --------------------Combo Box


label1 = tk.Label(root, text='Choose Machine')
label1.grid(column=0, row=0)

choose_machine_combo = ttk.Combobox(root, width=19)
choose_machine_combo.set('MC12')
choose_machine_combo['values'] = ('MC12', 'MH13', 'MH06')
choose_machine_combo.bind("<<>ComboboxSelected>")
choose_machine_combo.grid(column=1, row=0)

# --------------------Listbox

file_listbox = tk.Listbox(root, bg='light blue', width=80)
file_listbox.grid(column=0, row=4, columnspan=5, sticky=tk.E+tk.W)

# --------------------Buttons

folder_pick = tk.Button(root, text="Select Folder", relief=tk.RAISED,
                        width=16, bd=2, padx=10, pady=6)
folder_pick.bind('<ButtonRelease-1>', choose_folder)
folder_pick.grid(column=1, row=1)

process = tk.Button(root, text="Process Data", relief=tk.RAISED,
                    width=16, bd=2, padx=10, pady=6)
process.grid(column=1, row=2)

# --------------------Menu Bar
menubar = tk.Menu(root)
root.config(menu=menubar)
sub_menu = tk.Menu(menubar, tearoff=False)
menubar.add_cascade(label='File', menu=sub_menu)
# sub_menu.add_command(label='Open', command=open_file)


def write_to_spreadsheet():
    '''Formats and writes data to spreadsheet.'''
    wb = Workbook()
    highlight = NamedStyle(name="highlight")
    highlight.font = Font(bold=False, size=11)
    bd = Side(style='thin', color="000000")
    highlight.border = Border(left=bd, top=bd, right=bd, bottom=bd)
    wb.add_named_style(highlight)  # Register named style
    sh1 = wb.active
    sh1.title = 'Tool Usage Frequency'
    sh1.append(['Tool Number', 'Times Used', 'CT Number', 'Description'])
    sh1['A1'].font = Font(bold=True, size=11)
    sh1['A1'].border = Border(left=bd, top=bd, right=bd, bottom=bd)
    sh1['A1'].alignment = Alignment(horizontal='center')
    sh1['B1'].font = Font(bold=True, size=11)
    sh1['B1'].border = Border(left=bd, top=bd, right=bd, bottom=bd)
    sh1['B1'].alignment = Alignment(horizontal='center')
    sh1['C1'].font = Font(bold=True, size=11)
    sh1['C1'].border = Border(left=bd, top=bd, right=bd, bottom=bd)
    sh1['C1'].alignment = Alignment(horizontal='center')
    sh1['D1'].font = Font(bold=True, size=11)
    sh1['D1'].border = Border(left=bd, top=bd, right=bd, bottom=bd)
    sh1['D1'].alignment = Alignment(horizontal='center')
    rnum = 2
    sh1.column_dimensions['A'].width = (11.95)
    sh1.column_dimensions['B'].width = (10.6)
    sh1.column_dimensions['C'].width = (10)
    col_width2 = 0
    sh1_ct_data = Parse_Files.get_ct_number(Parse_Files.new_dict)
    for keys, values in Parse_Files.new_dict.items():
        sh1_tool_list_data = sh1_ct_data[keys]
        sh1.cell(row=rnum, column=1).value = int(keys)
        sh1.cell(row=rnum, column=2).value = int(values)
        sh1.cell(row=rnum, column=1).style = 'highlight'
        sh1.cell(row=rnum, column=2).style = 'highlight'
        sh1_ct_num, sh1_description = sh1_tool_list_data
        sh1.cell(row=rnum, column=3).value = sh1_ct_num
        sh1.cell(row=rnum, column=3).style = 'highlight'
        sh1.cell(row=rnum, column=4).value = sh1_description
        sh1.cell(row=rnum, column=4).style = 'highlight'
        if len(str(sh1_description)) > col_width2:
            col_width2 = len(str(sh1_description))
        rnum += 1
    sh1.column_dimensions['D'].width = (col_width2 * 1.125)
    # ----------------------------------------------------------
    sh2 = wb.create_sheet(title='Single Use List')
    wb.active = 2
    sh2.append(['Tool Number', 'Program Number', 'CT Number', 'Description'])
    sh2['A1'].font = Font(bold=True)
    sh2['A1'].border = Border(left=bd, top=bd, right=bd, bottom=bd)
    sh2['A1'].alignment = Alignment(horizontal='center')
    sh2['B1'].font = Font(bold=True)
    sh2['B1'].border = Border(left=bd, top=bd, right=bd, bottom=bd)
    sh2['B1'].alignment = Alignment(horizontal='center')
    sh2['C1'].font = Font(bold=True, size=11)
    sh2['C1'].border = Border(left=bd, top=bd, right=bd, bottom=bd)
    sh2['C1'].alignment = Alignment(horizontal='center')
    sh2['D1'].font = Font(bold=True, size=11)
    sh2['D1'].border = Border(left=bd, top=bd, right=bd, bottom=bd)
    sh2['D1'].alignment = Alignment(horizontal='center')
    rnum = 2
    col_width = (max_length(Parse_Files.single_list))
    sh2.column_dimensions['A'].width = (11.95)
    sh2.column_dimensions['B'].width = (col_width * 1.125)
    sh2.column_dimensions['C'].width = (10)
    col_width3 = 0
    for keys, values in Parse_Files.single_list.items():
        sh2_tool_list_data = sh1_ct_data[keys]
        sh2_ct_num, sh2_description = sh2_tool_list_data
        sh2.cell(row=rnum, column=1).value = int(keys)
        sh2.cell(row=rnum, column=2).value = (values)
        sh2.cell(row=rnum, column=1).style = 'highlight'
        sh2.cell(row=rnum, column=2).style = 'highlight'
        sh2.cell(row=rnum, column=3).value = sh2_ct_num
        sh2.cell(row=rnum, column=3).style = 'highlight'
        sh2.cell(row=rnum, column=4).value = sh2_description
        sh2.cell(row=rnum, column=4).style = 'highlight'
        if len(str(sh2_description)) > col_width3:
            col_width3 = len(str(sh2_description))
        rnum += 1
    sh2.column_dimensions['D'].width = (col_width3 * 1.125)
    # ----------------------------------------------------------
    sh3 = wb.create_sheet(title='Programmer Stats')
    wb.active = 3
    prog_stat = Parse_Files.extract_programmer()
    dp, jp, np, dc, jc, nc = prog_stat
    sh3.append(['Programmer', '# Programmed', '% Programmed'])
    sh3['A1'].font = Font(bold=True, size=11)
    sh3['A1'].border = Border(left=bd, top=bd, right=bd, bottom=bd)
    sh3['A1'].alignment = Alignment(horizontal='center')
    sh3['B1'].font = Font(bold=True)
    sh3['B1'].border = Border(left=bd, top=bd, right=bd, bottom=bd)
    sh3['B1'].alignment = Alignment(horizontal='center')
    sh3['C1'].font = Font(bold=True)
    sh3['C1'].border = Border(left=bd, top=bd, right=bd, bottom=bd)
    sh3['C1'].alignment = Alignment(horizontal='center')
    sh3.cell(row=2, column=1).value = 'Dave'
    sh3.cell(row=2, column=1).style = 'highlight'
    sh3.cell(row=2, column=3).value = dp
    sh3.cell(row=2, column=3).style = 'highlight'
    sh3.cell(row=2, column=2).value = dc
    sh3.cell(row=2, column=2).style = 'highlight'
    sh3.cell(row=3, column=1).value = 'John'
    sh3.cell(row=3, column=1).style = 'highlight'
    sh3.cell(row=3, column=3).value = jp
    sh3.cell(row=3, column=3).style = 'highlight'
    sh3.cell(row=3, column=2).value = jc
    sh3.cell(row=3, column=2).style = 'highlight'
    sh3.cell(row=4, column=1).value = 'No Name'
    sh3.cell(row=4, column=1).style = 'highlight'
    sh3.cell(row=4, column=3).value = np
    sh3.cell(row=4, column=3).style = 'highlight'
    sh3.cell(row=4, column=2).value = nc
    sh3.cell(row=4, column=2).style = 'highlight'
    sh3.column_dimensions['A'].width = (11)
    sh3.column_dimensions['B'].width = (15)
    sh3.column_dimensions['C'].width = (16)
    save_name = (('{}/{} Tool Usage Data.xlsx').format
                 (Parse_Files.folder_selected, Parse_Files.machine))
    wb.save(save_name)
    file_listbox.insert(tk.END, 'Operation Complete')
    file_listbox.see(tk.END)
    os.startfile(save_name)


root.mainloop()
