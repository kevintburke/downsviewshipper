#K@D SHIP SPREADSHEET PARSER

import pandas as pd
from tkinter import messagebox, filedialog
import tkinter as tk
from ollama import chat
import subprocess
import ast
import datetime
import re
import os

class DownsviewSHIPPER:
    def __init__(self, root):
        self.root = root
        self.root.title("K@D SHIP Spreadsheet Processor")

        tk.Label(root, text="This program will split and reformat spreadsheets for the K@D SHIP workflow based on material type, workflow, and action required").grid(column=1,row=0,pady=10)

        #Prompt user to load file
        tk.Label(root, text="Please select a file to process (in .xlsx format ONLY)").grid(column=1,row=1,pady=10)
        tk.Button(root, text="Load", command=self.getfile).grid(column=1,row=2,pady=5)

        #Use textboxes to display the column headers
        tk.Label(root,text="Select the correct field to draw each of the following elements (lists will populate after loading file):").grid(column=1,row=3,pady=10)
        self.utmms_box = tk.Listbox(root, exportselection=0, height=5)
        self.uomms_box = tk.Listbox(root, exportselection=0, height=5)
        self.title_box = tk.Listbox(root, exportselection=0, height=5)
        self.callno_box = tk.Listbox(root, exportselection=0, height=5)
        self.barcode_box = tk.Listbox(root, exportselection=0, height=5)
        self.description_box = tk.Listbox(root, exportselection=0, height=5)
        self.type_box = tk.Listbox(root, exportselection=0, height=5)
        self.workflow_box = tk.Listbox(root, exportselection=0, height=5)
        self.missing_box = tk.Listbox(root, exportselection=0, height=5)
        self.loaned_box = tk.Listbox(root, exportselection=0, height=5)
        self.itemid_box = tk.Listbox(root, exportselection=0, height=5)

        tk.Label(root, text="UofT MMS ID").grid(column=0,row=4)
        self.utmms_box.grid(column=0,row=5,padx=10)
        tk.Label(root, text="uOttawa MMS ID").grid(column=1,row=4)
        self.uomms_box.grid(column=1,row=5,padx=10)
        tk.Label(root, text="Title").grid(column=2,row=4)
        self.title_box.grid(column=2,row=5,padx=10)
        tk.Label(root, text="Call Number").grid(column=0,row=6)
        self.callno_box.grid(column=0,row=7,padx=10)
        tk.Label(root, text="Barcode").grid(column=1,row=6)
        self.barcode_box.grid(column=1,row=7,padx=10)
        tk.Label(root, text="Description").grid(column=2,row=6)
        self.description_box.grid(column=2,row=7,padx=10)
        tk.Label(root, text="Material Type").grid(column=0,row=8)
        self.type_box.grid(column=0,row=9,padx=10)
        tk.Label(root, text="Workflow 1/2").grid(column=1,row=8)
        self.workflow_box.grid(column=1,row=9,padx=10)
        tk.Label(root, text="Item Missing").grid(column=2,row=8)
        self.missing_box.grid(column=2,row=9,padx=10)
        tk.Label(root, text="Item on Loan?").grid(column=0,row=10)
        self.loaned_box.grid(column=0,row=11,padx=10)
        tk.Label(root, text="Item ID").grid(column=1,row=10)
        self.itemid_box.grid(column=1,row=11,padx=10)

        tk.Label(root, text="Click here to process the spreadsheet. Note that this process may run slowly. Check the terminal for status updates.").grid(column=1,row=12,pady=10)
        tk.Button(root, text="RUN", command=self.process_sheet).grid(column=1,row=13,pady=10)

        #Create df for inputs
        self.df = None

    def getfile(self):
        self.utmms_box.delete(0, tk.END)
        self.uomms_box.delete(0, tk.END)
        self.title_box.delete(0, tk.END)
        self.callno_box.delete(0, tk.END)
        self.barcode_box.delete(0, tk.END)
        self.description_box.delete(0, tk.END)
        self.type_box.delete(0, tk.END)
        self.workflow_box.delete(0, tk.END)
        self.missing_box.delete(0, tk.END)
        self.loaned_box.delete(0, tk.END)
        self.itemid_box.delete(0, tk.END)
        file = filedialog.askopenfilename()
        try:
            df = pd.read_excel(file)
            headers = list(df.columns.values)
            converters = {}
            for header in headers:
                converters[header] = str
            self.df = pd.read_excel(file, converters=converters)
            #Populate textboxes for user to select columns to process
            for column in list(self.df.columns.values):
                self.utmms_box.insert(tk.END, column)
                self.uomms_box.insert(tk.END, column)
                self.title_box.insert(tk.END, column)
                self.callno_box.insert(tk.END, column)
                self.barcode_box.insert(tk.END, column)
                self.description_box.insert(tk.END, column)
                self.type_box.insert(tk.END, column)
                self.workflow_box.insert(tk.END, column)
                self.missing_box.insert(tk.END, column)
                self.loaned_box.insert(tk.END, column)
                self.itemid_box.insert(tk.END, column)
        except Exception as e:
            messagebox.showerror(title="ERROR",detail=e)
        #Add Enumeration and Chronology fields
        self.df["Enum A"] = None
        self.df["Enum B"] = None
        self.df["Chron I"] = None
        self.df["Chron J"] = None

    def process_sheet(self):
        #Get user selections
        utmms_column = self.utmms_box.get(self.utmms_box.curselection())
        uomms_column = self.uomms_box.get(self.uomms_box.curselection())
        title_column = self.title_box.get(self.title_box.curselection())
        callno_column = self.callno_box.get(self.callno_box.curselection())
        barcode_column = self.barcode_box.get(self.barcode_box.curselection())
        description_column = self.description_box.get(self.description_box.curselection())
        type_column = self.type_box.get(self.type_box.curselection())
        workflow_column = self.workflow_box.get(self.workflow_box.curselection())
        missing_column = self.missing_box.get(self.missing_box.curselection())
        loaned_column = self.loaned_box.get(self.loaned_box.curselection())
        itemid_column = self.itemid_box.get(self.itemid_box.curselection())
        #Check for input in every column
        if utmms_column == None or uomms_column == None or title_column == None or callno_column == None or barcode_column == None or description_column == None or type_column == None or workflow_column == None or missing_column == None or itemid_column == None:
            messagebox.showerror(title="Missing Inputs",message="Error: Please select a column for every input")
            return None
        #Eliminate all columns except user selections to simplify output steps
        self.df = self.df[[utmms_column,uomms_column,title_column,callno_column,barcode_column,description_column,type_column,workflow_column,missing_column,loaned_column,itemid_column]]
        self.df = self.df.loc[self.df[loaned_column].isnull()]
        self.df = self.df[[utmms_column,uomms_column,title_column,callno_column,barcode_column,description_column,type_column,workflow_column,missing_column,itemid_column]]
        #Add columns for Library Code and Location Code and force dtype string for all columns
        self.df = self.df.reindex(columns = self.df.columns.tolist() + ["Library Code","Location Code","Enum A","Enum B","Chron I","Chron J"])
        self.df["Library Code"] = "OTTAWA"
        self.df["Location Code"] = "TRANSFER"
        self.df = self.df.astype(str)
        print(f'€€€€€€€€€€€€€€€\nself.df is...\n{self.df}\n€€€€€€€€€€€€€€€€€€€€€€€€€€€€€\n')
        #Eliminate missing titles and extract found titles
        df_missing = self.df.loc[self.df[missing_column].notnull()]
        print("MISSING\n",df_missing,df_missing.columns.values,"\n€€€€€€€€€€€€€€€€€€\n")
        #Create found titles list by doing left merge of original and missing item dfs
        df_found = pd.merge(self.df,df_missing,on=barcode_column,how="left_anti",suffixes=("","%"))
        print(df_found,df_found.columns.values)
        #Drop duplicate columns added in merge
        drop_columns = list(df_missing.columns.values)
        for column in drop_columns:
            if column == barcode_column:
                continue
            column_new = column + "%"
            df_found = df_found.drop(labels=column_new,axis=1)
        print("FOUND\n",df_found,df_found.columns.values,"\n€€€€€€€€€€€€€€€€€€\n")
        #Parse descriptions using local AI model for all found items here, before splitting monos/serials and workflows 1/2
        ##Consider revising path to make more robust? Or require user to select prompt file?
        print(os.getcwd())
        with open("downsviewshipper_prompt.txt", "r",encoding="utf-8") as fh:
            prompt = fh.read()
        print(f'Prompt read: ',prompt)
        #Check that ollama is running and prompt user to activate if not
        print("Checking ollama...")
        try:
            chat(model="qwen2.5-coder",messages=[{'role':'user','content':'Hello'}])
            print("Ollama is running.")
        except Exception as e:
            print(e)
            messagebox.showwarning(title="Ollama Error",message="Error. Please ensure ollama is running.")
            return None
        for index, row in df_found.iterrows():
            success = False
            while success == False:
                print("Calling ollama...")
                structured_data = chat(model="qwen2.5-coder",messages=[{'role':'user','content':prompt + str(row[description_column]),},]).message.content
                #Parse to remove ```python``` and ```json``` to reduce amount of errors and repeat calls req'd
                ##Add error checking to reject any outputs that contain text not in input?
                if structured_data.startswith("`"):
                    print("Cleaning structured data string: ", structured_data)
                    structured_data = re.sub(r'```python\n?(\{[^\}]+\})\n?```|```json\n?(\{[^\}]+\})\n?```',r'\1',structured_data)
                    print("Structured data cleaned to: ", structured_data)
                #Try parsing dictionary for required content
                try:
                    parsed_data = ast.literal_eval(structured_data)
                    parsed_data["Enum A"]
                    parsed_data["Enum B"]
                    parsed_data["Chron I"]
                    parsed_data["Chron J"]
                    print(f'Parse successful: {parsed_data}')
                    success = True
                except:
                    print("Parse failed",structured_data)
                    continue
            for i in ["Enum A","Enum B","Chron I","Chron J"]:
                row[i] = str(parsed_data[i])
            print(f'€€€€€€€€€€€€€€€€€€\nINDEX IS CURRENTLY {index}/{len(df_found)}\n€€€€€€€€€€€€€€€€€€')
            df_found.loc[index] = row
        #Create regex for matching material types to serials
        serial = re.compile(r'.*[Ii]ssue.*|.*[Ll]ooseleaf.*')
        #Replace with ISSUE or BOOK based on existing material type
        for index, row in df_found.iterrows():
            if re.match(serial, str(row[type_column])):
                row["Material Type"] = "ISSUE"
                df_found.loc[index] = row
            else:
                row["Material Type"] = "BOOK"
                df_found.loc[index] = row
        df_serials = df_found[df_found[type_column].isin(["ISSUE"])]
        print("FOUND SERIALS\n",df_serials,"\n€€€€€€€€€€€€€€€€€€\n")
        #Create monos df using left_anti merge from df_found and df_serials
        df_monos = pd.merge(df_found,df_serials,on=barcode_column,how="left_anti",suffixes=("","%"))
        print(df_monos,df_monos.columns.values)
        #Drop extra columns from merge as above
        drop_columns = list(df_serials.columns.values)
        for column in drop_columns:
            if column == barcode_column:
                continue
            column_new = column + "%"
            df_monos = df_monos.drop(labels=column_new,axis=1)
        print("FOUND MONOS\n",df_monos,df_monos.columns.values,"\n€€€€€€€€€€€€€€€€€€\n")
        #Split workflow 1 and 2
        df_serials_wf1 = df_serials.loc[df_serials[workflow_column].str.endswith("1")]
        print("SERIALS WF 1\n",df_serials_wf1,"\n€€€€€€€€€€€€€€€€€€\n")
        df_serials_wf2 = df_serials.loc[df_serials[workflow_column].str.endswith("2")]
        print("SERIALS WF 2\n",df_serials_wf2,"\n€€€€€€€€€€€€€€€€€€\n")
        df_monos_wf1 = df_monos.loc[df_monos[workflow_column].str.endswith("1")]
        print("MONOS WF 1\n",df_monos_wf1,"\n€€€€€€€€€€€€€€€€€€\n")
        df_monos_wf2 = df_monos.loc[df_monos[workflow_column].str.endswith("2")]
        print("MONOS WF 2\n",df_monos_wf2,"\n€€€€€€€€€€€€€€€€€€\n")
        #Create dataframes for outputs and set default values for library and location code
        missing_output = df_missing[[itemid_column,barcode_column,uomms_column,title_column,missing_column]]
        serials_output_workflow1 = df_serials_wf1[[utmms_column,title_column,"Library Code","Location Code",callno_column,barcode_column,"Material Type",description_column,"Enum A","Enum B","Chron I","Chron J"]]
        serials_output_workflow2 = df_serials_wf2[[uomms_column,title_column,"Library Code","Location Code",callno_column,barcode_column,"Material Type",description_column,"Enum A","Enum B","Chron I","Chron J"]]
        monos_output_workflow1 = df_monos_wf1[[utmms_column,title_column,"Library Code","Location Code",callno_column,barcode_column,"Material Type",description_column,"Enum A","Enum B","Chron I","Chron J"]]
        monos_output_workflow2 = df_monos_wf2[[uomms_column,title_column,"Library Code","Location Code",callno_column,barcode_column,"Material Type",description_column,"Enum A","Enum B","Chron I","Chron J"]]
        outputs = [("Missing Titles",missing_output),("Serials - Workflow 1",serials_output_workflow1),("Serials - Workflow 2",serials_output_workflow2),("Monos - Workflow 1",monos_output_workflow1),("Monos - Workflow 2",monos_output_workflow2)]
        currentdate = datetime.datetime.today().strftime('%m%d%H%M')
        for output in outputs:
            if len(output[1]) > 0:
                messagebox.showinfo(title="Save Output",message=f'Select a location and name for the following output: {output[0]}')
                file = filedialog.asksaveasfilename(initialfile=f'{output[0]}_{currentdate}.xlsx',defaultextension=".xlsx")
                output[1].astype(str).to_excel(file,index=False)
            else:
                print("Empty dataframe. Continuing...")
        
def main():
    root = tk.Tk()
    app = DownsviewSHIPPER(root)
    root.mainloop()

if __name__ == "__main__":
    main()