import tkinter as tk
from tkinter import messagebox
import datetime
from tkinter import filedialog
import win32com.client
import pytz
import time
import pyperclip
import os


class MainWindow:
    def __init__(self, master):
        self.master = master
        master.title('Urlaubseintrag')

        self.create_widgets()
        
    def create_widgets(self):
        self.name_label = tk.Label(self.master, text='Ihr Nachname:')
        self.name_entry = tk.Entry(self.master)
        self.output_label = tk.Label(self.master, text='Wo soll die .txt Datei gespeichert werden?:')
        self.output_entry = tk.Entry(self.master)
        self.browse_button = tk.Button(self.master, text='Durchsuchen', command=self.browse)
        self.instructions = tk.Button(self.master, text='Anleitung', command=self.show_instructions)
        self.information = tk.Button(self.master, text='Allgemeine Information', command=self.show_information)
        self.run_script = tk.Button(self.master, text='Skript starten', command=self.run_script)
        
        self.name_label.grid(row=0, column=0, padx=5, pady=5, sticky='W')
        self.name_entry.grid(row=0, column=1, padx=5, pady=5)
        self.output_label.grid(row=1, column=0, padx=5, pady=5, sticky='W')
        self.output_entry.grid(row=1, column=1, padx=5, pady=5)
        self.browse_button.grid(row=1, column=2, padx=5, pady=5)
        self.instructions.grid(row=2, column=0, padx=5, pady=5)
        self.information.grid(row=2, column=1, padx=5, pady=5)
        self.run_script.grid(row=2, column=2, padx=5, pady=5)
    
    def show_instructions(self):
        messagebox.showinfo('Instructions', 'Sie müssen Outlook und die Seite mit Ihren Urlaubsdaten geöffnet haben:\n\n'
                            '1. Geben Sie Ihren Nachnamen an.\n'
                            '2. Wählen Sie aus wo die .txt Datei gespeichert werden soll\n'
                            "3. Drücken Sie den 'Skript starten'-Button\n" 
                            '4. Sie haben danach 7 Sekunden um auf die Seite mit Ihren Urlaubsdaten zu wechseln\n'
                            '5. Warten bis die Meldung erscheint, dass der Skript gelaufen ist\n'  
                            "Das war's auch schon - Ihre Urlaubsdaten sind in Ihrem Outlook Kalender mit dem Betreff 'Urlaub IhrNachname' :-)")
        
    def show_information(self):
        messagebox.showinfo("Allgemeine Information", "Dieses Skript ist ein Bedienfeld-basiertes Tool, das es dem Benutzer ermöglicht, seine Urlaubsdaten in seinen Outlook-Kalender einzutragen. Der Benutzer muss lediglich seinen Namen eingeben und einen Speicherort auswählen, wo das Skript die erstellte .txt-Datei speichern soll. Durch Klicken auf den 'Skript starten' Button werden die Urlaubsdaten automatisch in den Outlook-Kalender eingetragen.")

def self_plug(self):
    messagebox.showinfo("About", "Dieses Skript wurde erstellt von Atakan Taşdirek. Die Sourcecode können Sie in meinem Github Account einsehen (@Dtakan). ")

    def browse(self):
        folder_selected = filedialog.askdirectory()
        self.output_entry.insert(0, folder_selected)

    def run_script(self):
        name = self.name_entry.get()
        output_dir = self.output_entry.get()
        file_path = os.path.join(output_dir, name + '.txt')
        
        time.sleep(7)

        win32com.client.Dispatch('WScript.Shell').SendKeys('^(a)')

        #press CONTROL + C twice to copy everything to the clipboard
        win32com.client.Dispatch('WScript.Shell').SendKeys('^(c)')
        win32com.client.Dispatch('WScript.Shell').SendKeys('^(c)')

        #get text from the clipboard
        text = pyperclip.paste()

        file_name = self.output_entry.get()

        # Check if file with the name exists
        counter = 1
        while os.path.exists('{}{}.txt'.format(file_name, counter)):
            counter += 1

        #create a new .txt file and insert the text from the clipboard
        with open('{}{}.txt'.format(file_name, counter), 'w', encoding='utf-8-sig') as f:
            f.write(text)

        #process the text from the .txt file
        with open('{}{}.txt'.format(file_name, counter), 'r', encoding='utf-8-sig') as f:
            lines = f.readlines()


        # process the text from the website
        valid_dates = []
        lines = text.splitlines()
        for line in lines:
            try:
                # split the line into start and end dates
                dates = [d for d in line.strip().split(' ') if d.count('.') == 2]
                if len(dates) >= 2:
                    start_date_str, end_date_str = dates[:2]
                    start_date = datetime.datetime.strptime(start_date_str, '%d.%m.%Y').date()
                    end_date = datetime.datetime.strptime(end_date_str, '%d.%m.%Y').date()
                    valid_dates.append((start_date, end_date))
            except ValueError:
                # ignore invalid date
                pass

        with open(file_path, 'w') as f:
            f.write(text)
        outlook = win32com.client.Dispatch('Outlook.Application')
        mapi = outlook.GetNamespace('MAPI')
        calendar = mapi.GetDefaultFolder(9)

        for start_date, end_date in valid_dates:
            appointment = calendar.Items.Add(1)
            timezone = pytz.timezone('Europe/Berlin') # or the appropriate time zone for your location
            start_datetime = pytz.utc.localize(datetime.datetime.combine(start_date, datetime.time(hour=1, minute=0, second=0)))
            appointment.Start = start_datetime
            appointment.End = datetime.datetime.combine(end_date, datetime.time(hour=23, minute=59, second=59))
            appointment.Subject = 'Vacation ' + self.name_entry.get()
            appointment.BusyStatus = 3
            appointment.Save()
        
        messagebox.showinfo('Erledigt!', 'Ihre Urlaubsdaten wurden in Outlook eingetragen und ihre Urlaubsdaten sind gespeichert in: ' + file_path)

root = tk.Tk()
app = MainWindow(root)
root.mainloop()