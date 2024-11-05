"""
Project: DOPI (Document Organizer for PDFs and Images)
Author: Artur Ikkert
Version: 1.0
Date: 05.11.2024

Description:
------------
DoPI ist ein Dokumentenorganisator, der PDF- und Bilddateien effizient verwaltet und organisiert.
Es enthält Funktionen zur Texterkennung (OCR) in Bildern, Dokumenten-Tagging und Filtermöglichkeiten,
um Dokumente leichter zu durchsuchen und abzurufen.

Requirements:
-------------
- Python 3.x
- Tesseract OCR (Portable-Version im Repository integriert)
- Zusätzliche Python-Bibliotheken (siehe requirements.txt)
"""

import shutil
import json
from customtkinter import *
import cv2
import pytesseract
from pypdf import PdfReader
import sqlite3
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

# Reference to the local Tesseract directory
pytesseract.pytesseract.tesseract_cmd = os.path.join(os.path.dirname(__file__), 'tesseract', 'tesseract.exe')

# Global variables
file = ""
search = ""
target_folder = ""


def first_path():
    if not target_folder:
        messagebox.showinfo("Pfad", "Vor der ersten Benutzung muss ein Speicherpfad für die Dateien "
                                    "ausgewählt werden")
        save_path()


# Loading and saving the path in the configuration file
def load_path():
    global target_folder
    if os.path.exists("config.json"):
        with open("config.json", "r") as jsonfile:
            data = json.load(jsonfile)
            target_folder = data.get("Storage path", "")
    return target_folder


def save_path():
    path = filedialog.askdirectory()
    if path:
        with open("config.json", "w") as jsonfile:
            json.dump({"Storage path": path}, jsonfile)
        global target_folder
        target_folder = path
        path_entry.configure(state="normal")
        path_entry.delete("0", END)
        path_entry.insert("0", path)
        path_entry.configure(state="readonly")
        message_text = "Speicherpfad gesetzt"
        message.configure(state="normal")
        message.delete(1.0, END)
        message.insert(END, message_text)
        message.configure(text_color="#75F94D", state="disabled")
        create_database()
        read_data()


# Initialisation of the database
def create_database():
    connection = sqlite3.connect(os.path.join(target_folder, "DOPI.db"))
    cursor = connection.cursor()
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS documents (
        id INTEGER PRIMARY KEY,
        name TEXT UNIQUE,
        keyword1 TEXT,
        keyword2 TEXT,
        date TEXT, 
        content TEXT)
    ''')
    connection.commit()
    connection.close()


# Insert and update data
def insert_data(name, keyword1, keyword2, date, content, segment_archiv):
    connection = sqlite3.connect(os.path.join(target_folder, "DOPI.db"))
    cursor = connection.cursor()
    cursor.execute('SELECT * FROM documents WHERE name = ?', (name,))
    existing_data = cursor.fetchone()

    # Check whether the file already exists
    if existing_data:
        result = messagebox.askyesno("Konflikt",
                                     f"Ein Dokument mit dem Namen\n'{name}'\n"
                                     f"existiert bereits. Möchten Sie die Inhalte ersetzen?")
        if not result:
            connection.close()
            return

    # Data is only overwritten if the copy process was successful
    archiviert = archive(file, target_folder, segment_archiv)
    if archiviert:
        # If the file already exists, the data is overwritten (Upsert (Update or Insert))
        cursor.execute('''
                INSERT INTO documents (name, keyword1, keyword2, date, content)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(name) DO UPDATE SET 
                    keyword1 = excluded.keyword1,
                    keyword2 = excluded.keyword2,
                    date = excluded.date,
                    content = excluded.content
                ''', (name, keyword1, keyword2, date, content))
        connection.commit()
        connection.close()
        read_data()
        search_document(search_field_entry)
    else:
        message_text = "\nDaten wurden nicht überschrieben!"
        message.insert(END, message_text)
        message.configure(text_color="#EB3324")
        message.configure(state="disabled")


# Open and scan image
def open_image():
    global file
    file = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg")])
    if file:
        image = cv2.imread(file)
        image_file_name = os.path.basename(file)
        name_entry.configure(state="normal")
        name_entry.delete("0", END)
        name_entry.insert("0", image_file_name)
        name_entry.configure(state="disabled")
        text = pytesseract.image_to_string(image)
        return text


# Open PDF and extract text
def open_pdf():
    global file
    file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file:
        # PDF mit PyPDF öffnen
        reader = PdfReader(file)
        pdf_file_name = os.path.basename(file)
        name_entry.configure(state="normal")
        name_entry.delete("0", END)
        name_entry.insert("0", pdf_file_name)
        name_entry.configure(state="disabled")
        text = ""
        for page in reader.pages:
            text += page.extract_text() + ("\n"*2)
        return text


# Archive file
def archive(source, target_folder, segment_archiv):
    copied = False
    if segment_archiv == "Copy":
        try:
            shutil.copy(source, target_folder)
            message.configure(state="normal")
            message_text = f"Datei erfolgreich kopiert"
            message.delete(1.0, END)
            message.insert(END, message_text)
            message.configure(text_color="#75F94D", state="disabled")
            copied = True
        except Exception as e:
            message.configure(state="normal")
            message_text = f"Fehler beim Kopieren der Datei: {e}"
            message.delete(1.0, END)
            message.insert(END, message_text)
        return copied
    else:
        try:
            # Extract file name from the source path and create complete target path
            filename = os.path.basename(source)
            target_path = os.path.join(target_folder, filename)
            if os.path.exists(target_path):
                os.remove(target_path)
            shutil.move(source, target_folder)
            message.configure(state="normal")
            message_text = f"Datei erfolgreich verschoben"
            message.delete(1.0, END)
            message.insert(END, message_text)
            message.configure(text_color="#75F94D", state="disabled")
            copied = True
        except Exception as e:
            message.configure(state="normal")
            message_text = f"Fehler beim verschieben der Datei: {e}"
            message.delete(1.0, END)
            message.insert(END, message_text)
        return copied


# Event for Segmented Button
def segment_event(file_type):
    scan_button.configure(command=handle_pdf_scan if file_type == "PDF" else handle_img_scan)


# Event for Image and PDF scanning
def handle_img_scan():
    text = open_image()
    if text:
        content_textbox.delete(1.0, END)
        content_textbox.insert(END, text)


def handle_pdf_scan():
    text = open_pdf()
    if text:
        content_textbox.delete(1.0, END)
        content_textbox.insert(END, text)


# Empty keyword fields
def clear():
    keyword1_entry.delete("0", END)
    keyword2_entry.delete("0", END)
    date_entry.delete("0", END)


# Read out data and insert into the treeview
def read_data():
    connection = sqlite3.connect(os.path.join(target_folder, "DOPI.db"))
    cursor = connection.cursor()
    cursor.execute("SELECT * FROM documents ORDER BY id DESC")
    data = cursor.fetchall()
    # Delete previous data in the treeview
    for row in tree.get_children():
        tree.delete(row)

    count = 0
    for row_data in data:
        single_row = []
        for column_data in row_data:
            column_data = str(column_data).replace("\n", "")
            single_row.append(column_data)
        if count % 2 == 0:
            tree.insert("", "end", values=single_row, tags=("evenrow", ))
        else:
            tree.insert("", "end", values=single_row, tags=("oddrow",))
        count += 1
    connection.close()


# Search document
def search_document(event):
    global search
    search = search_field_entry.get().lower()
    if search:
        connection = sqlite3.connect(os.path.join(target_folder, "DOPI.db"))
        cursor = connection.cursor()
        cursor.execute('''SELECT * FROM documents WHERE LOWER(name) LIKE ? OR LOWER(keyword1) LIKE ? OR 
                          LOWER(keyword2) LIKE ? OR LOWER(date) LIKE ? OR LOWER(content) LIKE ? 
                          ORDER BY id DESC''', (f'%{search}%',) * 5)
        data = cursor.fetchall()
        for row in tree.get_children():
            tree.delete(row)

        count = 0
        for row_data in data:
            single_row = []
            for column_data in row_data:
                # Write content in one line and shorten
                column_data = str(column_data).replace("\n", "")[0:300]
                single_row.append(column_data)
            if count % 2 == 0:
                tree.insert("", "end", values=single_row, tags=("evenrow",))
            else:
                tree.insert("", "end", values=single_row, tags=("oddrow",))
            count += 1
        connection.close()
    else:
        read_data()


# Show popup for complete content (Double-Click Event)
def show_popup(content, x, y):
    popup = CTkToplevel(root)
    popup.overrideredirect(True)  # Remove window frames
    popup.geometry(f"{int(x)}+{int(y)}")
    popup.focus_force()  # Sets the focus on the pop-up

    # List of column names (without ID)
    column_names = ["Name", "Stichwort 1", "Stichwort 2", "Datum", "Inhalt"]
    # Formatted content: Column name + cell value for each row
    formatted_content = ""
    for row in content:
        for i, value in enumerate(row):
            # Add column name and value
            if value is not None:
                formatted_content += f"{column_names[i]}: {value}\n"
            # Paragraph after Column 3 (Date)
            if i == 3:
                formatted_content += "\n"

    text_box = tk.Text(popup, width=70, height=18, wrap="word", font=("Arial", 13), background="#2B2B2B",
                       foreground="#DCE4EE", border="0")
    scrollbar = CTkScrollbar(popup, command=text_box.yview, fg_color="#2B2B2B")
    scrollbar.pack(side="right", fill="y")
    text_box.configure(yscrollcommand=scrollbar.set)
    text_box.insert("end", formatted_content)
    text_box.configure(state="disabled", spacing1=7)
    text_box.pack(fill="both", expand=True)
    if search:
        highlight_text(text_box, search)

    def close_popup(event):
        popup.destroy()
    popup.bind("<FocusOut>", close_popup)


# Highlight searched text
def highlight_text(popup_text_box, search):
    start = "1.0"  # Start position: first line, first character
    while True:
        start = popup_text_box.search(search, start, stopindex="end", nocase=True)
        if not start:
            break
        end = f"{start}+{len(search)}c"  # "c" = characters. For the calculation of the end position
        popup_text_box.tag_add("highlight", start, end)
        start = end  # Updates start to the end position so that the next search begins after it.
    popup_text_box.tag_config("highlight", background="#144870", foreground="#DCE4EE", font=("Arial", 13, "bold"))


# Open popup for the double-clicked line
def on_row_click(event):
    # Identify the clicked line
    selected_item = tree.focus()
    if selected_item:
        row_values = tree.item(selected_item)['values']
        connection = sqlite3.connect(os.path.join(target_folder, "DOPI.db"))
        cursor = connection.cursor()
        cursor.execute("SELECT name, keyword1, keyword2, date, content FROM documents WHERE name = ?",
                       (row_values[1],))
        content = cursor.fetchall()
        # Position of the popup
        x, y = table_frame.winfo_rootx() + 290, table_frame.winfo_rooty() + 24
        connection.close()
        show_popup(content, x, y)


# Open file
def open_file():
    selected_item = tree.focus()
    if not selected_item:
        messagebox.showwarning("Keine Auswahl", "Bitte eine Zeile auswählen.")
        return
    item_values = tree.item(selected_item)['values']
    file_path = os.path.join(target_folder, item_values[1])

    if os.path.exists(file_path):
        try:
            os.startfile(file_path)
        except Exception as e:
            messagebox.showerror("Fehler", f"Die Datei konnte nicht geöffnet werden: {e}")
    else:
        messagebox.showerror("Datei nicht gefunden", "Die angegebene Datei existiert nicht.")


# Delete file
def delete():
    selected_item = tree.focus()
    item_values = tree.item(selected_item)['values']
    result = messagebox.askyesno("Konflikt",
                                 f"Möchten Sie\n '{item_values[1]}'\n wirklich löschen?")
    if not result:
        return
    connection = sqlite3.connect(os.path.join(target_folder, "DOPI.db"))
    cursor = connection.cursor()
    cursor.execute('DELETE FROM documents WHERE name = ?', (item_values[1],))
    connection.commit()
    connection.close()

    file_path = os.path.join(target_folder, item_values[1])
    if os.path.exists(file_path):
        try:
            os.remove(file_path)
        except Exception as e:
            messagebox.showerror("Fehler", f"Die Datei konnte nicht gelöscht werden: {e}")
    else:
        messagebox.showerror("Datei nicht gefunden",
                             f"Die Datei\n'{item_values[1]}' \nwurde bereits gelöscht."
                             f"\nDie Datenbankeinträge wurden entfernt.")
    read_data()
    search_document(search_field_entry)


def button_state(event):
    selected_item = tree.focus()
    open_button.configure(state="normal")
    delete_button.configure(state="normal")
    if not selected_item:
        open_button.configure(state="disabled")
        delete_button.configure(state="disabled")


# Creating the GUI
root = CTk()
root.title("Dokumenten-Manager")
app_width, app_height = 1280, 800
set_appearance_mode("dark")
root.resizable(False, False)
# Place window in the centre of the screen
root.geometry(f"{app_width}x{app_height}+{(root.winfo_screenwidth() - app_width) // 2}+"
              f"{(root.winfo_screenheight() - app_height) // 2}")

# Tab groups
tabview = CTkTabview(master=root)
tabview.pack(fill="both", expand=True)
tabview.add("Dokument scannen")
tabview.add("Übersicht")
for button in tabview._segmented_button._buttons_dict.values():
    button.configure(width=200)

# Frame for the sidebar
sidebar_frame = CTkFrame(master=tabview.tab("Dokument scannen"), border_width=5, border_color="#343638")
sidebar_frame.pack(fill="y", padx=(120, 20), pady=(50, 130), side="left")

# Frame for Dokument scannen
scan_frame = CTkFrame(master=tabview.tab("Dokument scannen"))
scan_frame.pack(fill="both", expand=True, padx=20, pady=(50, 0), side="left")

# Segment Button
segment_type = CTkSegmentedButton(master=sidebar_frame, values=["PDF", "IMG"], command=segment_event)
segment_type.set("PDF")
segment_type.pack(fill="x", padx=20, pady=10)

segment_archiv = CTkSegmentedButton(master=sidebar_frame, values=["Copy", "Move"])
segment_archiv.set("Copy")
segment_archiv.pack(fill="x", padx=20, pady=10)

# Button for scanning
scan_button = CTkButton(master=sidebar_frame, text="Öffnen", corner_radius=32, command=handle_pdf_scan)
scan_button.pack(padx=20, pady=10)

# Button for clearing the keywords
clear_button = CTkButton(master=sidebar_frame, text="Stichwörter leeren", corner_radius=32, command=clear)
clear_button.pack(padx=20, pady=10)

# Button for inserting into the database
insert_button = CTkButton(master=sidebar_frame, text="Daten speichern", corner_radius=32,
                          command=lambda: [insert_data(name_entry.get(), keyword1_entry.get(),
                                                       keyword2_entry.get(), date_entry.get(),
                                                       content_textbox.get("1.0", END),
                                                       segment_archiv.get())])
insert_button.pack(padx=10, pady=(282, 10))

# Button to change the path
path_button = CTkButton(master=sidebar_frame, text="Pfad ändern", corner_radius=32, command=save_path)
path_button.pack(padx=10, pady=(10, 0))

# Input fields for the texts
name_label = CTkLabel(master=scan_frame, text="Dokumentenname")
name_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
name_entry = CTkEntry(master=scan_frame, width=600, placeholder_text="Wird automatisch ausgefüllt...")
name_entry.configure(state="disabled")
name_entry.grid(row=1, column=1, padx=10, pady=10)

keyword1_label = CTkLabel(master=scan_frame, text="Stichwort")
keyword1_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
keyword1_entry = CTkEntry(master=scan_frame, width=600)
keyword1_entry.grid(row=2, column=1, padx=10, pady=10)

keyword2_label = CTkLabel(master=scan_frame, text="Stichwort")
keyword2_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")
keyword2_entry = CTkEntry(master=scan_frame, width=600)
keyword2_entry.grid(row=3, column=1, padx=10, pady=10)

date_label = CTkLabel(master=scan_frame, text="Datum")
date_label.grid(row=4, column=0, padx=10, pady=10, sticky="w")
date_entry = CTkEntry(master=scan_frame, placeholder_text="TT.MM.JJJJ", width=600)
date_entry.grid(row=4, column=1, padx=10, pady=10)

content_label = CTkLabel(master=scan_frame, text="Inhalt")
content_label.grid(row=5, column=0, padx=10, pady=10, sticky="w")

content_textbox = CTkTextbox(master=scan_frame, border_width=2, height=300, width=600, wrap="word", fg_color="#343638")
content_textbox.grid(row=5, column=1, padx=10, pady=10)

path_label = CTkLabel(master=scan_frame, text="Speicherpfad")
path_label.grid(row=6, column=0, padx=10, pady=10, sticky="w")
path_entry = CTkEntry(master=scan_frame, placeholder_text="Speicherpfad", width=600)
path_entry.grid(row=6, column=1, padx=10, pady=10)

# Hidden message Box
display_text = tk.StringVar()
message = CTkTextbox(master=scan_frame, border_width=0, height=110, width=600, fg_color="#2B2B2B", state="disabled")
message.grid(row=7, column=1, padx=10, pady=10)

# Tab 2 ---------------------------------------------------------------------------------------------------------------

# Frame for the buttons and the search field
top_frame = CTkFrame(master=tabview.tab("Übersicht"))
top_frame.pack(fill="x", padx=20, pady=(20, 0))

# Frame for the Treeview table
table_frame = CTkFrame(master=tabview.tab("Übersicht"))
table_frame.pack(expand=True, fill="both", padx=(20, 0), pady=20)

# Creating the treeview table
tree = ttk.Treeview(master=table_frame, columns=[f"Col{i}" for i in range(0, 6)], show="headings", selectmode="browse",
                    displaycolumns=(1, 2, 3, 4, 5))
tree.column(1, minwidth=150, width=200)
tree.column(2, minwidth=150, width=200)
tree.column(3, minwidth=150, width=200)
tree.column(4, minwidth=70, width=70)
tree.column(5, minwidth=200, width=540)
tree.pack(expand=False, fill="both", side="left")

# Headings of the table
tree.heading(1, text="Name")
tree.heading(2, text="Stichwort")
tree.heading(3, text="Stichwort")
tree.heading(4, text="Datum")
tree.heading(5, text="Inhalt")

# Create CTkScrollbar for the Treeview-Table
ctk_textbox_scrollbar = CTkScrollbar(master=table_frame, command=tree.yview, fg_color="#2B2B2B")
ctk_textbox_scrollbar.pack(fill="y", side="left")
# Connect the textbox scroll event with the CTkScrollbar
tree.configure(yscrollcommand=ctk_textbox_scrollbar.set)

# Changing the style of the table
style = ttk.Style()
style.theme_use("alt")
# Config the treeview colors
style.configure("Treeview", background="#343638", foreground="#DCE4EE", rowheight=50, fieldbackground="#343638")
style.configure("Treeview.Heading", background="#1F6AA5", foreground="#DCE4EE", rowheight=75)
# Change selected color
style.map("Treeview", background=[("selected", "#144870")])
style.map("Treeview.Heading", background=[("active", "#144870")])
# Create striped row tags
tree.tag_configure("oddrow", background="#404245")
tree.tag_configure("evenrow", background="#343638")

# Button for opening the file from the selected line
open_button = CTkButton(master=top_frame, text="Datei öffnen", corner_radius=32, command=open_file, state="disabled")
open_button.grid(row=0, column=0, padx=5, pady=10, sticky="w")

# Button for deleting the file from the selected line
delete_button = CTkButton(master=top_frame, text="Datei löschen", corner_radius=32, command=delete, state="disabled")
delete_button.grid(row=0, column=1, padx=5, pady=10, sticky="w")

# Search field for the database
search_field_entry = CTkEntry(master=top_frame, placeholder_text="Suche...")
search_field_entry.grid(row=0, column=2, padx=(100, 0), pady=10, sticky="ew")

# Grid configuration for the Top_Frame
top_frame.grid_rowconfigure(0, weight=1)
top_frame.grid_columnconfigure(0, weight=0)
top_frame.grid_columnconfigure(1, weight=0)
top_frame.grid_columnconfigure(2, weight=1)

# Bind events
tree.bind("<Double-1>", on_row_click)
search_field_entry.bind("<KeyRelease>", search_document)
tree.bind("<<TreeviewSelect>>", button_state)

# Initialisation
load_path()
first_path()
path_entry.insert("0", target_folder)
path_entry.configure(state="readonly")
read_data()

# Start GUI loop
root.mainloop()
