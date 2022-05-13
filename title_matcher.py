#!/usr/bin/python3
import pandas as pd
import recordlinkage as rl
from recordlinkage.index import Full
from os import path

from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog

pd.set_option("max_colwidth", 200)

class Title_Matcher:

    def __init__(self, master):

        master.title('Fuzzy Matches on Titles')
        master.resizable(True, True)
        master.configure(background='#FFFFFF')

        self.file_io = None
        self.df1 = StringVar()
        self.df2 = StringVar()
        self.threshold = StringVar()

        

        ttk.Label(master, text='========== File ==========').pack(ipady=2)
        #ttk.Label(self.frame_content, text='Email:').grid(row=0, column=1, padx=5, sticky='sw')
        
        
        self.file_path = ttk.Entry(
            master, width=150, font=('Arial', 10))
        self.file_path.pack(fill=X,anchor='e',ipady=1.5)
        ttk.Button(master, text='Choose File . . .',
                   command=self.open_file).pack(anchor='w',fill=X)
        # self.progress_bar = ttk.Progressbar(master,orient=HORIZONTAL,length=800)
        # self.progress_bar.pack()
        # self.progress_bar.config(mode='indeterminate')
        # self.progress_bar.start()
        self.scale = ttk.Spinbox(master, textvariable = self.threshold,from_=10, to=100)
        self.scale.pack()
        self.scale.set(50)
        self.df1_combobox = ttk.Combobox(master,textvariable = self.df1)
        self.df2_combobox = ttk.Combobox(master,textvariable = self.df2)
        self.df1_combobox.pack(fill=X)
        self.df2_combobox.pack(fill=X)
        ttk.Label(master, text='========== Results ==========').pack()
        self.text_output = Text(master, width=150, height=10, font=('Arial', 10))
        self.text_output.pack(fill=BOTH,expand=True)
        ttk.Button(master, text='Clear',
                   command=self.clear).pack(side=RIGHT,anchor='e',fill=X,expand=TRUE)
        self.submit_button = ttk.Button(master, text='Submit',
                   command=self.submit)
        self.submit_button.pack(side=LEFT,anchor='w',fill=X,expand=TRUE)


    def submit(self):
        #self.progress_bar.start()
        ##self.df1_combobox.set('SpringAlmaOutput')
        ##self.df2_combobox.set('SpringBookstoreList')

        selected_threshold= self.threshold.get()
        input_file = self.file_path.get()
        
        #df1 = pd.read_excel(input_file,header=2,sheet_name=self.df1_combobox.get())
        df1 = pd.read_excel(input_file,sheet_name=self.df1_combobox.get())
        df2 = pd.read_excel(input_file,sheet_name=self.df2_combobox.get())
        indexer = rl.Index()
        indexer.add(Full())

        pairs = indexer.index(df1,df2,)
        print(len(pairs))

        comparer = rl.Compare()
        comparer.string('Title','Long Title',threshold=float(selected_threshold)/100,label='Title')

        potential_matches = comparer.compute(pairs, df1,df2)
        matches = potential_matches[potential_matches.sum(axis=1)> 0].reset_index()
        #print(matches)

        accumulated = matches.loc[:,['level_0','level_1']].merge(df1.loc[:,['Title','ISBN','Online Location']], left_on='level_0',right_index=True)
        accumulated = accumulated.merge(df2.loc[:,['Long Title','Internal ID','Section Code','Instructor Name']], left_on='level_1',right_index=True)
        accumulated.head()
        accumulated.to_excel('{}-{}.xlsx'.format(path.basename(self.file_io.name),selected_threshold),index=False,columns=['Internal ID','Long Title', 'Title','Online Location','Section Code','Instructor Name'])
        dfStyler = accumulated.style.set_properties(**{'text-align': 'left'})
        dfStyler.set_table_styles([dict(selector='th', props=[('text-align', 'left')])])
        self.text_output.delete(1.0, 'end')
        self.text_output.insert(END, accumulated.to_string(index=False,columns=['Internal ID','Long Title', 'Title']))
        #self.progress_bar.stop()

        #messagebox.showinfo(title='Explore California Feedback',message='Comments Submitted!')

    def clear(self):
        self.file_path.delete(0, 'end')
        #self.entry_email.delete(0, 'end')
        self.text_output.delete(1.0, 'end')
        self.df1_combobox['values']=""
        self.df2_combobox['values']=""
        self.df1_combobox.set('')
        self.df2_combobox.set('')

#not able to get this to work, box always blank
#but works if done in the open file dialog method
    def set_data_frame_combos(self):
        xl = pd.ExcelFile(self.file_io.name)
        self.df1_combobox['values']=xl.sheet_names

    def open_file(self):
        self.submit_button['state'] = 'disabled'
        self.file_io = filedialog.askopenfile(initialdir = "./")
        #self.progress_bar.start()
        self.file_path.delete(0, 'end')
        # not able to get this work
        #self.set_data_frame_combos()
        self.file_path.insert(END, self.file_io.name)
        #set the comboboxes
        xl = pd.ExcelFile(self.file_io.name)
        self.df1_combobox['values']=xl.sheet_names
        self.df1_combobox.set('Set this to Alma Sheet . . .')
        self.df2_combobox.set('Set this to Book Store Sheet . . .')
        self.df2_combobox['values']=xl.sheet_names
        #self.progress_bar.stop()
        self.submit_button['state'] = 'normal'



def main():

    root = Tk()
    # root.option_add('*tearOff',False)
    matchGUI = Title_Matcher(root)
    root.mainloop()


if __name__ == "__main__":
    main()
