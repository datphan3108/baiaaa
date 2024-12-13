import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import datetime
import os
import pandas as pd

DATA_FILE = "employee_data.csv"

def save_to_file(data):
    file_exists = os.path.isfile(DATA_FILE)
    with open(DATA_FILE, mode='a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if not file_exists:
            writer.writerow(["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", "Số CMND", "Nơi cấp", "Ngày cấp"])
        writer.writerow(data)

def load_file():
    if not os.path.isfile(DATA_FILE):
        return []
    with open(DATA_FILE, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        return list(reader)

def find_birthdays_today():
    employees = load_file()
    today = datetime.datetime.now().strftime('%d/%m/%Y')
    today_birthdays = [emp for emp in employees if emp['Ngày sinh'] == today]
    return today_birthdays

def export_data_to_excel():
    employees = load_file()
    if not employees:
        messagebox.showerror("Lỗi", "Không có dữ liệu để xuất!")
        return
    df = pd.DataFrame(employees)
    df['Ngày sinh'] = pd.to_datetime(df['Ngày sinh'], format='%d/%m/%Y', errors='coerce')
    df = df.sort_values(by='Ngày sinh', ascending=True)
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if output_file:
        df.to_excel(output_file, index=False, encoding='utf-8')
        messagebox.showinfo("Thành công", f"Xuất dữ liệu thành công vào {output_file}")

def save_employee():
    data = [
        emp_id.get(),
        emp_name.get(),
        emp_unit.get(),
        emp_position.get(),
        emp_dob.get(),
        emp_gender.get(),
        emp_id_number.get(),
        emp_issue_place.get(),
        emp_issue_date.get()
    ]
    if "" in data:
        messagebox.showerror("Lỗi", "Vui lòng điền đầy đủ thông tin!")
        return
    save_to_file(data)
    messagebox.showinfo("Thành công", "Lưu thông tin nhân viên thành công!")
    clear_inputs()

def show_today_birthdays():
    birthdays = find_birthdays_today()
    if not birthdays:
        messagebox.showinfo("Thông báo", "Không có nhân viên nào sinh nhật hôm nay.")
        return
    result = "\n".join([f"{emp['Mã']} - {emp['Tên']}" for emp in birthdays])
    messagebox.showinfo("Sinh nhật hôm nay", result)

def clear_inputs():
    emp_id.set("")
    emp_name.set("")
    emp_unit.set("")
    emp_position.set("")
    emp_dob.set("")
    emp_gender.set("")
    emp_id_number.set("")
    emp_issue_place.set("")
    emp_issue_date.set("")

root = tk.Tk()
root.title("Quản lý thông tin nhân viên")
root.geometry("600x400")

emp_id = tk.StringVar()
emp_name = tk.StringVar()
emp_unit = tk.StringVar()
emp_position = tk.StringVar()
emp_dob = tk.StringVar()
emp_gender = tk.StringVar()
emp_id_number = tk.StringVar()
emp_issue_place = tk.StringVar()
emp_issue_date = tk.StringVar()

tk.Label(root, text="Mã:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_id).grid(row=0, column=1, padx=5, pady=5)

tk.Label(root, text="Tên:").grid(row=0, column=2, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_name).grid(row=0, column=3, padx=5, pady=5)

tk.Label(root, text="Đơn vị:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_unit).grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Chức danh:").grid(row=1, column=2, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_position).grid(row=1, column=3, padx=5, pady=5)

tk.Label(root, text="Ngày sinh (dd/mm/yyyy):").grid(row=2, column=0, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_dob).grid(row=2, column=1, padx=5, pady=5)

tk.Label(root, text="Giới tính:").grid(row=2, column=2, padx=5, pady=5, sticky='w')
tk.OptionMenu(root, emp_gender, "Nam", "Nữ").grid(row=2, column=3, padx=5, pady=5)

tk.Label(root, text="Số CMND:").grid(row=3, column=0, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_id_number).grid(row=3, column=1, padx=5, pady=5)

tk.Label(root, text="Nơi cấp:").grid(row=3, column=2, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_issue_place).grid(row=3, column=3, padx=5, pady=5)

tk.Label(root, text="Ngày cấp (dd/mm/yyyy):").grid(row=4, column=0, padx=5, pady=5, sticky='w')
tk.Entry(root, textvariable=emp_issue_date).grid(row=4, column=1, padx=5, pady=5)

tk.Button(root, text="Lưu thông tin", command=save_employee).grid(row=5, column=0, columnspan=2, pady=10)
tk.Button(root, text="Sinh nhật hôm nay", command=show_today_birthdays).grid(row=5, column=2, pady=10)
tk.Button(root, text="Xuất danh sách", command=export_data_to_excel).grid(row=5, column=3, pady=10)

root.mainloop()
