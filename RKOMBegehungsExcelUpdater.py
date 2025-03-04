import pandas as pd
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog as fd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

from planio.planio_queries import (
    get_begehungsdaten
)

api_key = "547fa8685b2e0a692cdeb33aeda305edd27b0155"

def main():
    # def handle_boolean(value):
    #     if isinstance(value, bool):  # If it's already a boolean, return it as is
    #         return value
    #     elif isinstance(value, str):  # If it's a string
    #         # Convert to lowercase and check if it's 'true' or 'false'
    #         value = value.strip().lower()  # Strip to remove extra spaces and convert to lowercase
    #         if value == 'wahr':
    #             return True
    #         elif value == 'falsch':
    #             return False
    def select_file():
        output_label.configure(text="",bg="white")
        file_label.configure(bg="white")
        filename.configure(bg="white")
        root.configure(bg="white")
        filename.configure(text=fd.askopenfilename())
    def update_file(filename):
        # Function to copy the formatting from one column to another
        

        if filename.cget("text").endswith(".xlsx"):
            output_label.configure(text="Fehler (Zieldatei muss geschlossen sein)",bg="orange")
            file_label.configure(bg="orange")
            filename.configure(bg="orange")
            root.configure(bg="orange")
            begehungsdaten = get_begehungsdaten(api_key,'131')
            print(begehungsdaten)
            df = pd.read_excel(filename.cget("text"))
            # Iterate through each row
            for index, row in df.iterrows():
                # You can access individual values by column name
                if(not pd.isnull(row['Strasse'])):
                    zusatz = ""
                    if(not pd.isnull(row['Hhnrzusatz'])):
                        zusatz = row['Hhnrzusatz']
                    adressen = [(str(row['Gfrgebaeudeid']) + " - " + row['Strasse'] + " " +str(int(row['Hhnr'])) + zusatz + ", " + str(int(row['Plz'])) + " " + row['Ort']),(str(row['Gfrgebaeudeid']) + " - " + row['Strasse'] + " " +str(int(row['Hhnr'])) + zusatz + " " + str(int(row['Plz'])) + " " + row['Ort']),(str(row['Gfrgebaeudeid']) + " - " + row['Strasse'] + " " +str(int(row['Hhnr'])) + " " + zusatz + ", " + str(int(row['Plz'])) + " " + row['Ort']),(str(row['Gfrgebaeudeid']) + " - " + row['Strasse'] + " " +str(int(row['Hhnr'])) + " " + zusatz + " " + str(int(row['Plz'])) + " " + row['Ort'])]
                    for adresse in adressen:
                        if not begehungsdaten[begehungsdaten["address"] == adresse].head(1).empty:
                            planio_row = begehungsdaten[begehungsdaten["address"] == adresse].head(1)
                    
                    if not planio_row.empty:
                        df.loc[index, "Status"] = planio_row["status"].iloc[0]
                        if not pd.isnull(planio_row["protokoll"].iloc[0]):
                            df.loc[index, "Potokoll versandt"] = planio_row["protokoll"].iloc[0]
                        df.loc[index, "Sachstand"] = planio_row["sachstand"].iloc[0]
                        if not pd.isnull(row["Erschließung-Bemerkung"]):
                            df.loc[index, "Erschließung-Bemerkung"] = row["Erschließung-Bemerkung"] + " + " + planio_row["bemerkung"].iloc[0]
                        else:
                            df.loc[index, "Erschließung-Bemerkung"] = planio_row["bemerkung"].iloc[0]
                        # if handle_boolean(row["Nutzungsvereinbarung"]) and planio_row["status"].iloc[0] == "Wartend":
                        #     print("change" + str(planio_row["issue_id"].iloc[0]))
                            


            # Load the existing workbook
            book = load_workbook(filename.cget("text"))
            sheet = book.worksheets[0]
            def copy_column_formatting(sheet, source_col, target_col, start_row=1, end_row=sheet.max_row):
                for row in range(start_row, end_row + 1):
                    # Copy the cell formatting from the source column to the target column
                    source_cell = sheet.cell(row=row, column=source_col)
                    target_cell = sheet.cell(row=row, column=target_col)

                    # Copy font
                    target_cell.font = Font(name=source_cell.font.name, size=source_cell.font.size, 
                                bold=source_cell.font.bold, italic=source_cell.font.italic, 
                                color=source_cell.font.color, underline=source_cell.font.underline)
                    # Copy fill (background color)
                    if source_cell.fill.start_color:
                        target_cell.fill = PatternFill(start_color=source_cell.fill.start_color, 
                                                    end_color=source_cell.fill.end_color, 
                                                    fill_type=source_cell.fill.fill_type)
                    else:
                        target_cell.fill = PatternFill(fill_type='none')  # If no fill, apply no fill
                    # Copy alignment
                    target_cell.alignment = Alignment(horizontal=source_cell.alignment.horizontal,
                                           vertical=source_cell.alignment.vertical,
                                           text_rotation=source_cell.alignment.text_rotation,
                                           wrap_text=source_cell.alignment.wrap_text,
                                           shrink_to_fit=source_cell.alignment.shrink_to_fit)
                    # Copy borders
                    target_cell.border = Border(
            left=source_cell.border.left,
            right=source_cell.border.right,
            top=source_cell.border.top,
            bottom=source_cell.border.bottom
        )

            # Specify the column range you want to copy formatting from (for example, columns 1 to 4)
            start_col = 1  # Column A
            end_col = 26    # Column D (you can adjust this)
            new_start_col = 2 # Start from column E to extend formatting

            for col in range(start_col, end_col + 1):
                copy_column_formatting(sheet, start_col, col + (new_start_col - start_col))
            # Write the column names (headers) to the first row
            for col_num, column_name in enumerate(df.columns, start=1):
                sheet.cell(row=1, column=col_num, value=column_name)

            # Write the data to the sheet, starting from row 2
            for row_num, row in enumerate(df.values, start=2):
                for col_num, value in enumerate(row, start=1):
                    sheet.cell(row=row_num, column=col_num, value=value)

            # Save the modified workbook
            book.save(filename.cget("text"))

            output_label.configure(text="Fertig",bg="lightgreen")
            file_label.configure(bg="lightgreen")
            filename.configure(bg="lightgreen")
            root.configure(bg="lightgreen")
        else:
            output_label.configure(text="Datei nicht kompatibel",bg="yellow")
            file_label.configure(bg="yellow")
            filename.configure(bg="yellow")
            root.configure(bg="yellow")

    def on_resize(event):
        file_label.configure(wraplength=root.winfo_width()-50)
        output_label.configure(wraplength=root.winfo_width()-50)
        filename.configure(wraplength=root.winfo_width()-50)

    # Create the main window
    root = tk.Tk()
    window_width = 400
    window_height = 150
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int((screen_height - window_height) / 2)
    position_left = int((screen_width - window_width) / 2)
    root.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")
    root.configure(bg="white")
    root.title("RKOM Begehungs Excel Updater")
    # Create a label widget
    root.bind("<Configure>",on_resize)
    file_label = tk.Label(root, text='Gewählte Datei:',bg="white",wraplength = window_width-50)
    file_label.pack()
    filename = tk.Label(root,text="Keine",bg="white",wraplength = window_width-50) #Keine
    filename.pack()
    
    file_button = ttk.Button(root,text="Datei auswählen",command = lambda: select_file())
    file_button.pack()
    update_button = ttk.Button(root,text="Updaten",command = lambda: update_file(filename))
    update_button.pack()
    
    output_label = tk.Label(root, text='',bg="white",wraplength = window_width-50)
    output_label.pack()
    # Start the GUI event loop
    root.mainloop()

main()