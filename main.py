import tkinter as tk
from tkinter import filedialog
import openpyxl
from bs4 import BeautifulSoup


def choose_html_file():
    html_file = filedialog.askopenfilename(filetypes=[("HTML files", "*.html")])
    html_file_entry.delete(0, tk.END)
    html_file_entry.insert(0, html_file)


def process_html_to_excel():
    html_file = html_file_entry.get()
    excel_file = excel_file_entry.get()

    try:
        with open(html_file, 'r', encoding='utf-8') as file:
            content = file.read()
    except FileNotFoundError:
        result_label.config(text="Error: The HTML file not found.")
        return

    soup = BeautifulSoup(content, 'html.parser')

    try:
        workbook = openpyxl.load_workbook(excel_file)
        worksheet = workbook.active
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        worksheet = workbook.active

    table = soup.find('table')
    if table:
        for row in table.find_all('tr'):
            question_cell = row.find('td')
            if question_cell:
                question_text = question_cell.get_text().strip()
                answer_cell = question_cell.find_next('td')
                if answer_cell:
                    answer_text = answer_cell.get_text().strip()
                    if answer_text != "No":
                        worksheet.append([question_text, answer_text])

    workbook.save(excel_file)
    result_label.config(text=f"Data from HTML file appended to {excel_file} successfully.")


# Create the GUI
root = tk.Tk()
root.title("HTML to Excel Converter")

tk.Label(root, text="Select HTML File:").pack()
html_file_entry = tk.Entry(root)
html_file_entry.pack()
tk.Button(root, text="Browse", command=choose_html_file).pack()

tk.Label(root, text="Enter Excel File Name:").pack()
excel_file_entry = tk.Entry(root)
excel_file_entry.pack()

tk.Button(root, text="Convert", command=process_html_to_excel).pack()

result_label = tk.Label(root, text="")
result_label.pack()

root.mainloop()
