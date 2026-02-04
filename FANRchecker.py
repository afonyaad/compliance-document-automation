#!/usr/bin/env python
# coding: utf-8

# In[1]:


import FreeSimpleGUI as sg
import pandas as pd
import webbrowser
from pathlib import Path
import shutil

# GUI Builder

def build_gui() -> sg.Window:
    sg.theme("SystemDefault")
    layout = [
        [sg.Text("Excel file:"),
         sg.Input(key="-EXCEL-", enable_events=True),
         sg.FileBrowse(file_types=[("Excel Files", "*.xlsx"), ("Legacy Excel Files", "*.xls")])],

        [sg.Text("Datasheets folder:"),
         sg.Input(key="-FOLDER-"),
         sg.FolderBrowse()],

        [sg.Text("Destination folder:"),
         sg.Input(key="-DEST-"),
         sg.FolderBrowse()],

        [sg.Button("Check Missing"), sg.Button("Copy Datasheets"), sg.Button("Open Link"), sg.Button("Exit")],
        [sg.Listbox(values=[], size=(85, 20), key="-MISSING-", enable_events=True)],
    ]
    return sg.Window("FANR Datasheet Checker", layout, finalize=True)

# Excel Reader

def read_materials_from_excel(excel_path: Path) -> list[str]:
    """Read unique, non-empty product codes from column C."""
    engine = "xlrd" if excel_path.suffix.lower() == ".xls" else "openpyxl"
    df = pd.read_excel(excel_path, engine=engine)
    return (
        df.iloc[:, 2]
          .dropna()
          .astype(str)
          .str.strip()
          .unique()
          .tolist()
    )

# Missing Checker

def find_missing_datasheets(folder: Path, codes: list[str]) -> list[str]:
    """Return codes for which there is no file ending with '{code}.pdf' in folder."""
    missing = []
    for code in codes:
        pattern = f"{code}.pdf"
        exists = any(
            f.is_file() and f.name.lower().endswith(pattern.lower())
            for f in folder.iterdir()
        )
        if not exists:
            missing.append(code)
    return missing

# Main

def main():
    window = build_gui()
    codes = []
    missing = []

    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "Exit"):
            break

        if event == "Check Missing":
            excel_path = Path(values["-EXCEL-"])
            folder_path = Path(values["-FOLDER-"])

            if not excel_path.exists():
                sg.popup_error("Excel file not found.")
                continue
            if not folder_path.exists():
                sg.popup_error("Datasheets folder not found.")
                continue

            try:
                codes = read_materials_from_excel(excel_path)
            except Exception as e:
                sg.popup_error(f"Error reading Excel: {e}")
                continue

            missing = find_missing_datasheets(folder_path, codes)
            display = [f"{code}: https://www.se.com/ae/en/product/{code}/" for code in missing] or ["All datasheets are present."]
            window["-MISSING-"].update(display)

        elif event == "Copy Datasheets":
            # copy all existing datasheets for codes to destination
            folder_path = Path(values["-FOLDER-"])
            dest_path = Path(values["-DEST-"])
            if not folder_path.exists() or not dest_path:
                sg.popup_error("Please select both FZE folder and destination.")
                continue
            if not dest_path.exists():
                try:
                    dest_path.mkdir(parents=True)
                except Exception as e:
                    sg.popup_error(f"Cannot create destination: {e}")
                    continue

            copied = []
            not_found = []
            for code in codes:
                # find first matching file
                pattern = f"{code}.pdf"
                matches = [f for f in folder_path.iterdir() if f.is_file() and f.name.lower().endswith(pattern.lower())]
                if matches:
                    src = matches[0]
                    dst = dest_path / src.name
                    try:
                        shutil.copy2(src, dst)
                        copied.append(code)
                    except Exception as e:
                        not_found.append(code)
                else:
                    not_found.append(code)

            msg = [f"Copied: {', '.join(copied)}"] if copied else []
            if not_found:
                msg.append(f"Missing for copy: {', '.join(not_found)}")
            sg.popup("\n".join(msg) if msg else "No codes to copy.")

        elif event == "Open Link":
            selected = values.get("-MISSING-")
            if selected and ": " in selected[0]:
                url = selected[0].split(": ", 1)[1]
                webbrowser.open(url)

    window.close()

if __name__ == "__main__":
    main()


# In[ ]:




