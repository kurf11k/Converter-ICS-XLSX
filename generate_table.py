import ics
import pandas as pd
import tkinter as tk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import os
from tkinter import font

def get_day(x):
    return {
        "1": "PO",
        "2": "UT",
        "3": "ST",
        "4": "ČT",
        "5": "PA",
        "6": "SO",
        "7": "NE"
    }[x]
    
def event_to_dict(event):
    return {
        'Zápas': event.name,
        'Datum': event.begin.date().strftime('%d.%m.%Y'),
        'Den': get_day(event.begin.date().strftime('%u')),
        'Čas': event.begin.time().strftime('%H:%M')
    }
    
def set_time_zone(datetime):
    datetime = datetime.tz_localize("Europe/London")
    datetime = datetime.tz_convert("Europe/Prague")
    return datetime


def load_file(path_to_file):
    with open(path_to_file, 'r', encoding="utf8") as f:
        data = ""
        for line in f:
            if line.startswith("STATUS"):          
                newline = line.replace("\n", "CONFIRMED\n") 
            else:    
                newline= line
            data += newline
        return data

def parse_home_away_teams(df):
    df[["Domácí", "Hosté"]] = df["Zápas"].str.split(":", expand=True)
    
    words_to_remove = ("Sokol", "Tělovýchovná", "jednota", "Tělocvičná", "FK", "SK", ",", "z.s.", "Sportovní", "klub", "Spartak")
    cols_where_remove = ("Domácí", "Hosté")

    for w in words_to_remove:
        for col in cols_where_remove:
            df[col] = df[col].str.replace(w, "", regex=True)

    for col in cols_where_remove:    
        df[col] = df[col].str.strip()
        df[col] = df[col].str.replace("  ", " ")
        
    return df

def parse_date_time(df):
    df["DatumČas"] = df["Datum"] + " " + df["Čas"]
    df["DatumČas"] = pd.to_datetime(df["DatumČas"], format="%d.%m.%Y %H:%M")
    df = df.sort_values("DatumČas")

    df["DatumČas"] = df["DatumČas"].apply(set_time_zone)
    df["Zóna"] = df["DatumČas"].dt.strftime('%z')
    df = df.astype({"Zóna": "int"})
    df["Zóna"] = df["Zóna"] / 100 - 1
    df["DatumČas"] += pd.to_timedelta(df["Zóna"], unit='h')
    df["Čas"] = df["DatumČas"].dt.strftime('%H:%M')
    
    return df        

def generate_file_to_table(path_to_file): 
    data = load_file(path_to_file)
            
    icsFile = ics.Calendar(data)
    events = [event_to_dict(event) for event in icsFile.events]

    df = pd.DataFrame(events)

    df = parse_home_away_teams(df)
    df = parse_date_time(df)
    
    cols_to_remove = ["Zápas", "DatumČas", "Zóna"]
    df = df.drop(cols_to_remove, axis=1)
    
    df["Výsledek"] = ""
    
    if not os.path.exists("outputs"):
        os.mkdir("outputs")
    
    output_file = path_to_file.split("/")[-1]
    output_file = "./outputs/" + output_file.replace("ics", "xlsx")
    
    df.to_excel(output_file, index=False, startrow = 2, startcol = 2)
    

root = tk.Tk()
root.title('Table generator')
root.resizable(False, False)
root.geometry('600x130')


def select_file():
    filetypes = [("Calendar files", '*.ics')]

    filename = fd.askopenfilename(
        title='Open a file',
        #initialdir='/',
        filetypes=filetypes)

    path_to_file_text.delete("1.0", "end")
    path_to_file_text.insert("1.0", filename)

def on_click_generate():
    path = path_to_file_text.get("1.0", "end-1c")
    try:
        generate_file_to_table(path)
        message = "Succesfuly exported to outputs folder."
        icon="info"
        path_to_file_text.delete("1.0", "end")
    
    except Exception as ex:
        message="Cannot export\n" + str(ex)
        icon="warning"
    
    showinfo(
        title="Information",
        message=message,
        icon=icon
    )


my_font = font.Font(size=13, weight="bold")
# open button
open_file_button = tk.Button(
    root,
    text='Open a File',
    command=select_file,
    width=15,
    height=8,
    font=my_font
)

generate_button = tk.Button(
    root,
    text='Export',
    command=on_click_generate,
    width=15,
    height=8,
    font=my_font
)

path_to_file_label = tk.LabelFrame(root, text="Path to file")
path_to_file_text = tk.Text(path_to_file_label, height=1)

path_to_file_label.pack(side="top", pady=5, padx=5)
path_to_file_text.pack(side="top", pady=5, padx=10)
generate_button.pack(side="right", padx=10, pady=10)
open_file_button.pack(side="right", padx=10, pady=10)

# run the application
root.mainloop()
