import glob
import os
from datetime import datetime
import re
import json
#import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import webbrowser
import ast

class Book:
    def __init__(self,ptitle,pauthor):
        self.title = ptitle
        self.author = pauthor
        self.records = []
    def __str__(self):
        string = self.title+'|'+self.author[0:20]+'\n'
        for rec in self.records:
            string += '-  ' +rec.category+'|'+str(rec.page)+'|'+str(rec.start_loc)+'|'+str(rec.end_loc)+'|'+str(rec.date_time)+'|'+rec.note+'|'+rec.content[0:20]+'\n'
        return string
class Record:
    def __init__(self, ptitle,pauthor,pcategory,ppage,pstart_loc,pend_loc,pdate_time,pcontent):
        self.title = ptitle
        self.author = pauthor
        self.category = pcategory
        self.page = int(ppage)
        self.start_loc = int(pstart_loc)
        self.end_loc = int(pend_loc)
        self.date_time = pdate_time
        self.content = pcontent
        self.note = ""
    def __str__(self):
        return self.title[0:20]+'|'+self.author[0:20]+'|'+self.category+'|'+str(self.page)+'|'+str(self.start_loc)+'|'+str(self.end_loc)+'|'+str(self.date_time)+'|'+self.note+'|'+self.content[0:20]

def load_config():
    global config
    try:
        with open('config.json', 'r') as f:
            config = json.load(f)
    except OSError as e:
        print('Config file not found. Created with Default Values.')
        config = {'path': 'M:\\Calibre Book Library_2\\Kindle\\',
              'mapping' : {'section':['section','part'],'chapter':['chapter'],'heading':['heading','#'],\
                           'subheading':['subheading','##'],'important':['important','###'],\
                           'chapsummary':['chapsummary'],\
                           'reserved':['introduction','appendix','foreword','dedication','toc','conclusion',\
                                       'acknowledgments','notes','references','summary','preface']},\
                  'f_pg_no' : True,'f_pg_no_heading' : True,'f_loc_no' : True,'f_export_all' : False}
        
        with open('config.json', 'w') as f:
            json.dump(config, f)

def update_config():
    with open('config.json', 'w') as f:
        json.dump(config, f)

def mapping_values(mapping_):
    headings = []
    for lst in list(mapping_.values()):
        for word in lst:
            headings.append(word)
    return headings
    
def shorten_path(path):
    parts = path.split('\\')
    print(parts)
    for i in range(len(parts)-1):
        part = parts[i]
        if len(part) > 20:
            parts[i] = part[0:20]+'~'
            print(part)
    
    if len(parts[-1]) == 0:
        parts[-1] = parts[-2][:]
        
    if len(parts[-1]) > 25:
        if parts[-1].endswith('.txt'):
            ext = '.txt'
        else:
            ext=''
        parts[-1] = parts[-1][0:25]+'~'+ext
        print(part)
    
    short = parts[0]+'\\'+parts[1]+'\\..\\'+parts[-1]
    return short

def browse():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_path
    print('browse clicked')
    filename = filedialog.askdirectory(title = 'Select Clippings Folder')
    filename = filename.replace('/','\\')+'\\'
    if(filename != '\\'):
        folder_path.set(shorten_path(filename))
        config['path'] = filename
        update_config()
        print(filename)
    else:
        print('no change')
        
def open_help():
    webbrowser.open("https://pages.ramialhalabi.com/kindle-highlights-to-markdown")

def load_file():
    
    path = config['path']
    list_of_files = []
    
    list_of_files = glob.glob(path+'*.txt',recursive=True) #Check folder
    
    if len(list_of_files) == 0:
        list_of_files = glob.glob(path+'**\\*.txt',recursive=True) #Check Subfolders

    if len(list_of_files) == 0:
        messagebox.showerror(title='Error',message='Not able to find a .txt file in the folder specified')
        return []
    
    latest_file = max(list_of_files, key=os.path.getctime)
    
    print ('Latest File: '+latest_file)
    
    display_latest_file(latest_file)    
    
    data = open(latest_file, encoding="utf8")
    #lines = data.readlines()
    lines = [line.rstrip('\n') for line in data]
    data.close()

    record = [] #3 lines of related data
    records = [] #list of records above
    for line in lines:
        if line != '==========':
            if line != '':
                record.append(line)
        else:
            records.append(record)
            record = []

    records_list = [] #list of Record objects
    for record in records:
        index = record[0].rfind(' (')
        author = record[0].rstrip()[index+2:-1]
        title = record[0] = record[0][0:index]
        meta = record[1][7:].split(' | ')
        content = record[2]
        content = content.replace('$','\$')

        page = 0
        if meta[0].find('on page') != -1: #if book has pages
            page = meta[0][meta[0].rfind(' ')+1:]

            if meta[1].find('Location') == -1: #if book does not have Locations, use page values as locations
                meta.insert(1,"Location: " + str(page))
                if page.find('-') != -1: # if page number is an interval, take the first number
                    page = page.split('-')[0]
        else:    
            #print(meta[0][meta[0].find('Location'):])
            location = meta[0][meta[0].find('Location'):]
            meta.insert(1,location)

        category = meta[0].split(' ')[0]
        locations = meta[1].split(' ')[1]

        if locations.find('-') != -1:
            start_loc = locations.split('-')[0]
            end_loc = locations.split('-')[1]
        else:
            start_loc = end_loc = locations

        date_time_temp = meta[2][9:]
        date_time = datetime.strptime(date_time_temp,'%A, %B %d, %Y %I:%M:%S %p')

        records_list.append(Record(title,author,category,page,start_loc,end_loc,date_time,content))

    records = records_list

    book_titles = []
    books = []
    for record in records:
        if record.title not in book_titles:
            book_titles.append(record.title)
            new_book = Book(record.title,record.author)
            new_book.records.append(record)
            books.append(new_book)
        else:
            for book in books:
                if record.title == book.title:
                    book.records.append(record)

    for book in books:
        book.records.sort(key=lambda x: x.end_loc)
        book.records.sort(key=lambda x: x.start_loc)
        #print(book)
        n = len(book.records)
        for i in range(n):
            records = book.records
            rec = records[i]
            if rec.category == 'Note':
                if i < n-1:
                    rec2 = records[i+1]
                    if rec2.page == rec.page and rec.start_loc == rec2.end_loc: #abs(rec2.start_loc - rec.start_loc < 2):
                        #highlight after this note is a related to this highlight

                        if i > 0: #check if record before this note is also matching this highlight
                            rec3 = records[i-1]
                            if rec3.page == rec.page and rec.start_loc == rec3.end_loc: #abs(rec3.start_loc - rec.start_loc) < 2:
                                #find distance between starts and ends
                                d2 = abs(rec2.start_loc - rec.start_loc) + abs(rec2.end_loc - rec.end_loc)
                                d3 = abs(rec3.start_loc - rec.start_loc) + abs(rec3.end_loc - rec.end_loc)
                                if d2 < d3:
                                    rec2.note = rec.content
                                    rec.content = ''
                                    continue
                                else:
                                    rec3.note = rec.content
                                    rec.content = ''
                                    continue
                        rec2.note = rec.content
                        rec.content = ''
                        continue
                if i > 0:
                    rec3 = records[i-1]
                    if rec3.page == rec.page and rec.start_loc == rec3.end_loc: #abs(rec3.start_loc - rec.start_loc) < 2:
                        #note is related to this highlight
                        rec3.note = rec.content
                        rec.content = ''
                        continue

                ############################################# Less strict rules
                tolerance = 3
                
                if i < n-1:
                    rec2 = records[i+1]
                    if rec2.page == rec.page and \
                    (abs(rec2.start_loc - rec.start_loc < tolerance) or abs(rec2.end_loc - rec.end_loc < tolerance)):
                        #highlight after this note is related to this highlight

                        if i > 0: #check if record before this note is also matching this highlight
                            rec3 = records[i-1]
                            if rec3.page == rec.page and\
                            (abs(rec3.start_loc - rec.start_loc) < tolerance or\
                             abs(rec3.end_loc - rec.end_loc) < tolerance):
                                #find distance between starts and ends
                                d2 = abs(rec2.start_loc - rec.start_loc) + abs(rec2.end_loc - rec.end_loc)
                                d3 = abs(rec3.start_loc - rec.start_loc) + abs(rec3.end_loc - rec.end_loc)
                                if d2 < d3:
                                    rec2.note = rec.content
                                    rec.content = ''
                                    continue
                                else:
                                    rec3.note = rec.content
                                    rec.content = ''
                                    continue
                        rec2.note = rec.content
                        rec.content = ''
                        continue
                if i > 0:
                    rec3 = records[i-1]
                    if rec3.page == rec.page and \
                    (abs(rec3.start_loc - rec.start_loc) < tolerance or\
                     abs(rec3.end_loc - rec.end_loc) < tolerance):
                        #note is related to this highlight
                        rec3.note = rec.content
                        rec.content = ''
                        continue
 
        n = len(book.records)
        i = 0
        while i < n:
            if book.records[i].content == '':
                del book.records[i]
                i-=1
                n-=1
            i+=1
    
    #print('-------------------------')        
    #for book in books:
    #    print(book.title)
    
    display_book_options(books)
    
    return books

def display_latest_file(latest_file):
    global latest_file_var
    latest_file_var.set(shorten_path(latest_file))
    
    #win.grid_rowconfigure(2, minsize=20)

    latest_file_txt = Label(master=win, text="Lastest File: ", font=("Calibri", 14)) # Create a text label
    latest_file_txt.grid_remove()
    latest_file_txt.grid(row=3, column=1,sticky='W', padx=10)

    loaded_file_lbl = Label(master=win,textvariable=latest_file_var,font=("Calibri", 10), bg ="LightGrey")
    loaded_file_lbl.grid_remove()
    loaded_file_lbl.grid(row=4, column=1 , columnspan=3,sticky='W',padx=20)
    
    config_button = Button(master=win, text="Configure Headings",width = 16, command=lambda: configure_headings())
    config_button.grid(row=4, column=3,padx=10)
    
    pass

def display_book_options(books):
    win.grid_rowconfigure(5, minsize=20)
    
    book_list_txt = Label(master=win, text="Select Books to Export to MD: ", font=("Calibri", 14)) # Create a text label
    book_list_txt.grid(row=6, column=1,columnspan=3,sticky='W', padx=10)
    
    frame = Frame(master=win)
    scrollbar = Scrollbar(frame, orient=VERTICAL)
    listbox = Listbox(frame, yscrollcommand=scrollbar.set, selectmode='multiple')
    scrollbar.config(command=listbox.yview)
 
    frame.grid_columnconfigure(1, minsize=555)
    frame.grid_columnconfigure(2, minsize=10)
    scrollbar.grid(row=0,column=2, sticky='NS')
    listbox.grid(row=0, column=1,sticky='NSEW')
    frame.grid(row=7,column=1,columnspan=3,sticky='NSEW',padx=(15,10))
    
    for book in books:
        listbox.insert(END,book.title)
    
    def select_all():
        listbox.select_set(0, END)
    def clear_all():
        listbox.selection_clear(0, END) 
    
    select_all_button = Button(text="Select All", command=select_all)
    select_all_button.grid(row=6, column=2,padx=10)
    
    clear_all_button = Button(text="Clear Selection", command=clear_all)
    clear_all_button.grid(row=6, column=3,padx=10)

    
    win.grid_rowconfigure(7, minsize=20)
    
    export_options_txt = Label(text="Export Options: ", font=("Calibri", 14)) # Create a text label
    export_options_txt.grid(row=8, column=1,columnspan=2,sticky='W', padx=10)
    
    cb_page_no_heading = Checkbutton( text="Show heading page number", variable=bool_page_no_heading)
    cb_page_no = Checkbutton( text="Show highlight/note page number", variable=bool_page_no)
    cb_loc_no = Checkbutton( text="Show highlight/note location number", variable=bool_loc_no)
    cb_export_all = Checkbutton( text="Export all selected books into one file", variable=bool_export_all)
    
    cb_page_no_heading.grid(row=9,column=1,columnspan=2,sticky='W', padx=10)
    cb_page_no.grid(row=10,column=1,columnspan=2,sticky='W', padx=10)
    cb_loc_no.grid(row=11,column=1,columnspan=2,sticky='W', padx=10)
    cb_export_all.grid(row=12,column=1,columnspan=2,sticky='W', padx=10)
    
    try:
        if config['f_pg_no_heading']:
            cb_page_no_heading.select()
            
        if config['f_pg_no']:
            cb_page_no.select()

        if config['f_loc_no']:
            cb_loc_no.select()

        if config['f_export_all']:
            cb_export_all.select()
    except:
        print("Error setting checkboxes")
        pass
        
    export_button = Button(text="Export", command= lambda : export(books,listbox.curselection()))
    export_button.grid(row=8, column=3,padx=10, pady=10)
    pass

def export(books,selected):
        # selected = listbox.curselection()
        if len(selected) == 0:
            messagebox.showerror(title='Error',message='No Books Selected')
            return
        os.makedirs('.\Markdown_files', exist_ok=True)
        
        config['f_pg_no_heading'] = bool_page_no_heading.get() == 1
        config['f_pg_no'] = bool_page_no.get() == 1
        config['f_loc_no'] = bool_loc_no.get() == 1
        config['f_export_all'] = bool_export_all.get() == 1
        
        
        
        if bool_export_all.get() == 1:
            export_all(books,selected)
        else:
            export_books(books,selected)
            
        messagebox.showinfo(title='Done',message='Files Exported Successfully!')
        update_config()
        
def print_rec_to_file(rec,md):
    note = rec.note
    content = rec.content
    page = rec.page
    frm = rec.start_loc
    to = rec.end_loc
    date_time = rec.date_time
    string = ''
    if note != '': # Note is not empty, i.e. this is a note or a heading
        note_l = note.lower()
        mapping_values_list = mapping_values(mapping)
        if note.lower().startswith(tuple(mapping_values_list)):
            # Check if note starts with any of the heading words or symbols
            
            if content == '': ## Unmatched heading
                return
            
            if note_l.startswith(tuple(mapping['section'])):
                string += "---  \n### " + note + ": "+ content + '  \n\n'

            elif note_l.startswith(tuple(mapping['important'])):
                string += "> **_" + content + '_**  \n'    # Bold Italic Text
                if page == 0 or not config['f_pg_no_heading']:
                    string += '\n'
                
            elif note_l.startswith(tuple(mapping['subheading'])):
                string += "> **" + content + '**  \n'    # Bold Text 
                if page == 0 or not config['f_pg_no_heading']:
                    string += '\n'

            elif note_l.startswith(tuple(mapping['heading'])):
                string += "##### " + content + '  \n\n'  

            elif note_l.startswith(tuple(mapping['chapter'])):
                string += "#### " + note + ": "+ content + '  \n\n'

            elif note_l.startswith(tuple(mapping['chapsummary'])):
                string += "##### Chapter Summary  \n>" + content + '  \n'
                if page == 0 or not config['f_pg_no_heading']:
                    string += '\n'

            else:
                string += "#### " + content + '  \n\n'

            if page != 0 and config['f_pg_no_heading']:
                if note_l.startswith(tuple(mapping['important'])):
                    string += '> _~Page:'+ str(page) + '~_  \n\n'  
                else:
                    string += '_~Page:'+ str(page) + '~_  \n\n'

                
        else: ## not a heading, just a note
            string += '> **Note:** ' + note + "  \n> **Highlight:** "+ content +'  \n'
            
            if page != 0 and config['f_pg_no']: #page number exists and page number display enabled
                
                if config['f_loc_no']:          #location number display enabled
                    string += '> _~Page:'+ str(page) + '~_        _~Location:' + str(frm) + '-' + str(to) + '~_  \n\n'
                else:                           #location number display disabled, only show page
                    string += '> _~Page:'+ str(page) + '~_  \n\n'
                    
            else:            #page number doesn't exist or page number display disabled
                if config['f_loc_no']:          #location number display enabled
                    string += '> _~Location:' + str(frm) + '-' + str(to) + '~_  \n\n'
                else:
                    string += '\n'
    
    else: #Only highlight, no note or heading
        string += '> ' + content + '  \n'
        
        if page != 0 and config['f_pg_no']: #page number exists and page number display enabled
                
            if config['f_loc_no']:          #location number display enabled
                string += '> _~Page:'+ str(page) + '~_        _~Location:' + str(frm) + '-' + str(to) + '~_  \n\n'
            else:                           #location number display disabled, only show page
                string += '> _~Page:'+ str(page) + '~_  \n\n'

        else:            #page number doesn't exist or page number display disabled
            if config['f_loc_no']:          #location number display enabled
                string += '> _~Location:' + str(frm) + '-' + str(to) + '~_  \n\n'
            else:
                string += '\n'

    md.write(string)

def export_books(books,index_list):
    for i in range(len(books)):
        if i in index_list:
            book = books[i]
            md = open(".\Markdown_files\Highlights_"+re.sub(r"[^a-zA-Z0-9]+", ' ', book.title)+".md","w+", encoding = 'utf8')
            md.write("# Kindle Highlights\n")
            md.write("## **"+book.title+"**\n")
            md.write("###### *"+book.author+"*\n")
            md.write("---\n")
            for rec in book.records:
                print_rec_to_file(rec, md)
            md.close()

def export_all(books,index_list):
    md = open(".\Markdown_files\Highlights_All.md","w+", encoding = 'utf8')
    md.write("# Kindle Highlights\n")
    for i in range(len(books)):
        if i in index_list:
            book = books[i]
            md.write("---\n")
            md.write("## **"+book.title+"**\n")
            md.write("###### *"+book.author+"*\n")
            md.write("---\n")
            for rec in book.records:
                print_rec_to_file(rec, md)
    md.close()
    pass

def close(root):
    config['f_pg_no_heading'] = bool_page_no_heading.get() == 1
    config['f_pg_no'] = bool_page_no.get() == 1
    config['f_loc_no'] = bool_loc_no.get() == 1
    config['f_export_all'] = bool_export_all.get() == 1
    update_config()
    root.destroy()

def configure_headings():
    win2 = Toplevel()
    win2.title('Configure Headings')
    #win2.geometry("600x400")
    win2.grab_set() # disable main window usage
    win2.grid_columnconfigure(1, minsize=100)
    win2.grid_columnconfigure(2, minsize=440)
    
    categories = []
    labels = []
    entries = []
    text1 = 'Refer to documentation for information on editing (Particularly, using repeated characters)'
    Label(master=win2, text = text1, anchor=W, font=("Calibri", 10))\
    .grid(row=0, column=1 , columnspan=2,sticky='EW', padx=(20,0),pady=(10,10))
    
    i=0
    
    for category in mapping:
        categories.append(category)
        
        labels.append(Label(master=win2, text = category.upper()+':', anchor=W, width = 14, font=("Calibri", 12), bg ="LightGrey"))
        labels[i].grid(row=i+1, column=1 , columnspan=1,sticky='W', padx=(20,0),pady=2)
        
        entries.append(Entry(master=win2))
        entries[i].grid(row=i+1, column=2,sticky='EW',padx=(0,20))
        entries[i].insert(0, str(mapping[category]))
        
        i += 1
    
    win2.grid_rowconfigure(i+1, minsize=30)
    
    i+=1
        
    cancel_button= Button(master=win2, text="Cancel", width = 8, command=lambda: win2.destroy())
    cancel_button.grid(row=i+1, column=1,sticky='W', padx=(40,0),pady=(0,20))
    
    save_button = Button(master=win2, text="Save", width = 8, command=lambda: save(categories,entries,win2))
    save_button.grid(row=i+1, column=2, sticky='E', padx=(0,40),pady=(0,20))
    
    pass

def save(categories,entries,win2):
    lists = [] #empty list that will hold the new typed-in headings
    entry = ''
    try:
        for entry in entries:
            x = ast.literal_eval(entry.get())
            x = [n.strip() for n in x]
            lists.append(x)
    except Exception as e:
        messagebox.showerror(title='Error',message='Error parsing entry: \n'+ entry.get() +'\n\n Details:' + str(e) )
    
    for i in range(len(categories)):
        mapping[categories[i]] = lists[i] 
        
    config['mapping'] = mapping #on save
    win2.destroy()
    pass

global config

load_config()

mapping = config['mapping']

# TKinter Setup
win = Tk()
win.title('Kindle Highlights to Markdown')
win.geometry("600x500")
win.protocol('WM_DELETE_WINDOW', lambda: close(win))

win.grid_columnconfigure(1, minsize=300)
win.grid_columnconfigure(2, minsize=150)
win.grid_columnconfigure(3, minsize=150)

#Flags
bool_page_no_heading = IntVar()
bool_page_no = IntVar()
bool_loc_no = IntVar()
bool_export_all = IntVar()


path_label_txt = Label(master=win, text="Clippings Folder: ", font=("Calibri", 14)) # Create a text label
path_label_txt.grid(row=0, column=1,sticky='W', padx=10)

folder_path = StringVar()
folder_path.set(shorten_path(config['path']))

chosen_path_lbl = Label(master=win, textvariable=folder_path,font=("Calibri", 10), bg ="LightGrey")
chosen_path_lbl.grid(row=1, column=1 , columnspan=2,sticky='W', padx=20)

browse_button = Button(master=win, text="Browse", command=lambda: browse())
browse_button.grid(row=2, column=2,padx=10)

load_button = Button(master=win, text="Load Latest File",width = 16, command=lambda: load_file())
load_button.grid(row=2, column=3,padx=10)

help_button = Button(master=win, text="How To Use App",width = 16, command=lambda: open_help())
help_button.grid(row=0, column=3,padx=10)

latest_file_var = StringVar()

win.mainloop()

