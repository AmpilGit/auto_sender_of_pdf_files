import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2image import convert_from_path
import pytesseract
import re
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from tkinter import scrolledtext
import openpyxl
import imaplib
import time
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\tesseract.exe'

# Глобальная переменная для хранения пути к папке вывода
result_path_folder_var = ""
def extract_text_from_pdf(file_path):
    images = convert_from_path(file_path, first_page=1, last_page=1)
    full_text = ""
    try:
        text = pytesseract.image_to_string(images[0], lang='rus+eng')
        full_text += text + "\n"
    except Exception as e:
        full_text += "NoN"
    return full_text
def load_data_from_excel(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:  
            data.append((str(row[0]).strip(), str(row[1]).strip()))
    return data

def combine_pdfs_to_one(folder_path, search_data):
    global result_path_folder_var
    if not result_path_folder_var:
        messagebox.showerror("Ошибка", "Папка вывода не выбрана!")
        return

    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]

    with open(os.path.join(result_path_folder_var, "result.txt"), 'w', encoding='utf-8') as f:
        for pdf_file in pdf_files:
            file_path = os.path.join(folder_path, pdf_file)
            try:
                text = extract_text_from_pdf(file_path)
                newtext = ""
                number_of_checked_values = 0
                for lk_value, email in search_data:
                    if text.find(lk_value) != -1:
                        number_of_checked_values += 1
                        newtext += lk_value + "\n"
                        current_pdf_path = file_path
                        print(current_pdf_path, " ", email)
                        message = MIMEMultipart()
                        email_content = email
                        message['Subject'] = email_content
                        message['From'] = 'moderator@rosholod.org'
                        message['To'] = email
                        message_copy_for_moderator = MIMEMultipart()
                        message_copy_for_moderator['Subject'] = email_content
                        message_copy_for_moderator['From'] = 'moderator@rosholod.org'
                        message_copy_for_moderator['To'] = 'moderator@rosholod.org'
                        with open(current_pdf_path, 'rb') as fp:
                            att = MIMEApplication(fp.read(), _subtype="pdf")
                            att.add_header('Content-Disposition', 'attachment', filename=os.path.basename(current_pdf_path))
                            message.attach(att)
                            message_copy_for_moderator.attach(att)
                        try:
                            
                            # Send the email via SMTP
                            with smtplib.SMTP_SSL('smtp.mail.ru', 465) as smtp:
                                smtp.login('moderator@rosholod.org', 'bWT8rBssWzsrtFSYF0nf')
                                smtp.send_message(message)
                                smtp.send_message(message_copy_for_moderator)
                            first_8_lines = text.splitlines()[:10]
                            f.write(
                            "ЕСТЬ ЗНАЧЕНИЯ ИЗ СПИСКА " + file_path + "\n" +
                            "найденное значение: " + lk_value + " документ отправлен: " + email + "\n" +
                            '\n'.join(first_8_lines) + "\n \n \n \n \n"
                            )
                            print("Письмо успешно отправлено!")
                        except Exception as e:
                            f.write(
                            "ЕСТЬ ЗНАЧЕНИЯ ИЗ СПИСКА " + file_path + "\n" +
                            "найденное значение: " + lk_value + " документ не отправлен: " + email +"\n" + str(e) + "\n" +
                            '\n'.join(first_8_lines) + "\n \n \n \n \n"
                            )
                            print(f"Ошибка при отправке письма: {e}")
                if number_of_checked_values == 0:
                    first_8_lines = text.splitlines()[:30]
                    f.write(
                        "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n" +
                        "НЕ НАШЛОСЬ " + file_path + '\n'.join(first_8_lines) +
                        "\n !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! \n \n \n \n \n"
                    )
            except Exception as e:
                print(f"Ошибка при обработке файла {str(pdf_file)}: {e}")


def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        try:
            excel_file_path = result_path_folder_var + "/patterns.xlsx"
            if not excel_file_path:
                messagebox.showerror("Ошибка", "Файл patterns.xlsx не выбран!")
                return
            search_data = load_data_from_excel(result_path_folder_var + "/patterns.xlsx")
            
            combine_pdfs_to_one(folder_path, search_data)
            result = open(os.path.join(result_path_folder_var, "result.txt"), encoding='utf-8')
            copyable_text = scrolledtext.ScrolledText(tk.Tk(),
                                                      width=200,
                                                      height=100,
                                                      font=("Times New Roman", 12))
            result_text = ''.join(result.readlines())
            copyable_text.insert(tk.INSERT, result_text)
            copyable_text.pack(padx=20, pady=20)
            result.close()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

def select_result_folder():
    global result_path_folder_var
    folder_path = filedialog.askdirectory()
    if folder_path:
        result_path_folder_var = folder_path
        result_folder_label.config(text=f"Выбранная папка вывода: {folder_path}/result.txt")
root = tk.Tk()
root.title("PDF Processor")
root.geometry("500x500")

label = tk.Label(root, text="Приложение для автоматической рассылки писем ", font=("Arial", 12))
label.pack(pady=20)
label = tk.Label(root, text="по номерам счетов, как только вы выберете папку с pdf файлами", font=("Arial", 12))
label.pack(pady=20)
label = tk.Label(root, text=",которые вы хотите разослать автоматически начнется рассылка", font=("Arial", 12))
label.pack(pady=20)
label = tk.Label(root, text="ВНИМАНИЕ!!!! Файл patterns.xlsx должен находиться в папке вывода", font=("Arial", 12))
label.pack(pady=20)
label = tk.Label(root, text="Выберите папку с PDF-файлами:", font=("Arial", 12))
label.pack(pady=20)
button = tk.Button(root, text="Выбрать папку", command=select_folder, font=("Arial", 12))
result_button = tk.Button(root, text="Выбрать папку вывода", command=select_result_folder, font=("Arial", 12))
result_button.pack(pady=10)
result_folder_label = tk.Label(root, text="Папка вывода не выбрана", font=("Arial", 10), fg="red")
result_folder_label.pack(pady=5)

button.pack(pady=10)

root.mainloop()