import json
from openpyxl import load_workbook
from tkinter import Tk, filedialog, messagebox, Button, Label, StringVar
import os

def extract_data(sheet):
    data = {
        "PalletID": sheet['A18'].value,
        "Details": {
            "Length": sheet['B18'].value,
            "Width": sheet['C18'].value,
            "Delivery": sheet['B20'].value,
            "Notched": sheet['B21'].value,
            "HT": sheet['B22'].value,
            "Type": sheet['B23'].value,
            "Build": {
                "TOP": {
                    "material": sheet['B27'].value,
                    "board_number": sheet['C27'].value,
                    "board_h": sheet['D27'].value,
                    "board_w": sheet['E27'].value,
                    "board_l": sheet['F27'].value,
                    "pulled_price": sheet['G27'].value,
                    "board_cost": sheet['I27'].value,
                    "total_cost": sheet['J27'].value,
                    "nails": sheet['K27'].value,
                    "points": sheet['L27'].value
                },
                "MIDDLE": {
                    "material": sheet['B29'].value,
                    "board_number": sheet['C29'].value,
                    "board_h": sheet['D29'].value,
                    "board_w": sheet['E29'].value,
                    "board_l": sheet['F29'].value,
                    "pulled_price": sheet['G29'].value,
                    "board_cost": sheet['I29'].value,
                    "total_cost": sheet['J29'].value,
                    "nails": sheet['K29'].value,
                    "points": sheet['L29'].value
                },
                "BOTTOM": {
                    "material": sheet['B31'].value,
                    "board_number": sheet['C31'].value,
                    "board_h": sheet['D31'].value,
                    "board_w": sheet['E31'].value,
                    "board_l": sheet['F31'].value,
                    "pulled_price": sheet['G31'].value,
                    "board_cost": sheet['I31'].value,
                    "total_cost": sheet['J31'].value,
                    "nails": sheet['K31'].value,
                    "points": sheet['L31'].value
                }
            }
        }
    }
    return data

def main_app():
    root = Tk()
    root.title("Excel to JSON Converter")
    root.geometry("400x220")
    root.resizable(False, False)

    selected_file = StringVar()
    selected_file.set("No file selected.")

    def select_file():
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file_path:
            selected_file.set(file_path)
        else:
            selected_file.set("No file selected.")

    def convert_to_json():
        file_path = selected_file.get()
        if not os.path.isfile(file_path):
            messagebox.showerror("Error", "Please select a valid Excel file first.")
            return
        try:
            workbook = load_workbook(file_path, data_only=True)
            sheet = workbook.active
            data = extract_data(sheet)
            json_path = filedialog.asksaveasfilename(
                title="Save JSON File",
                defaultextension=".json",
                filetypes=[("JSON Files", "*.json")]
            )
            if json_path:
                with open(json_path, 'w') as f:
                    json.dump(data, f, indent=4)
                messagebox.showinfo("Success", "JSON file created successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

    Label(root, text="Step 1: Select your Excel file").pack(pady=(20, 5))
    Button(root, text="Select Excel File", command=select_file, width=20).pack()
    Label(root, textvariable=selected_file, wraplength=350, fg="blue").pack(pady=(5, 15))
    Label(root, text="Step 2: Convert to JSON").pack()
    Button(root, text="Convert to JSON", command=convert_to_json, width=20).pack(pady=(5, 10))

    root.mainloop()

if __name__ == "__main__":
    main_app()