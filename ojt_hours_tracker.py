import tkinter as tk
from tkinter import messagebox,simpledialog
from tkinter import Toplevel,Label
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
import os

FILENAME ="ojt_hours.xlsx"

def setup_excel():
    if not os.path.exists(FILENAME):
        wb = Workbook()
        ws = wb.active
        ws.title = "OJT Logs"
        ws.append(["Date", "Day", "Time IN", "Time Out", "Total Hours"])
        wb.save(FILENAME)

def log_to_excel(date,day,time_in,time_out,total_hours):
    wb = load_workbook(FILENAME)
    ws = wb.active
    ws.append([date, day, time_in, time_out, total_hours])
    wb.save(FILENAME)

def edit_logs():
    wb = load_workbook(FILENAME)
    ws = wb.active 

    edit_date = simpledialog.askstring("Edit Entry", "Enter date to edit (YYYY-MM-DD): ")
    found = False

    for row in ws.iter_rows(min_row=2, values_only=False):
        if row[0].value == edit_date:
            found = True
            new_time_in = simpledialog.askstring("Edit", f"Old Time In: {row[2].value}\nNew Time In (HH:MM): ")
            new_time_out = simpledialog.askstring("Edit", f"Old Time In: {row[3].value}\nNew Time Out (HH:MM): ")
            
            row[2].value = new_time_in
            row[3].value = new_time_out


            try:
                fmt = "%H:%M"
                t_in = datetime.strptime(new_time_in, fmt)
                t_out = datetime.strftime(new_time_out, fmt)
                total_hours = round((t_out - t_in).seconds / 3600, 2)
                row[4].value = total_hours
            except:
                messagebox.showerror("Error", "Invalid time format.")
            break

    if found:
        wb.save(FILENAME)
        messagebox.showinfo("Success", "Log update.")
    else:
        messagebox.showwarning("Not found", "No entry for that date found. ")

def time_in_action():
    global time_in
    time_in = datetime.now()
    time_in_label.config(text=f"Time In: {time_in.strftime('%H:%M')}")

def time_out_action():
    if not time_in:
        messagebox.showwarning("Missing", "You must time in First . ")
        return
    
    time_out = datetime.now()
    time_out_label.config(text=f"Time Out: {time_out.strftime('%H:%M')}")

    total_hours = round((time_out - time_in).seconds / 3600, 2)

    today = datetime.now()
    date = today.strftime("%Y-%m-%d")
    day = today.strftime("%A")

    if day not in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]:
        messagebox.showwarning("Warning", f"It's {day}. Logging only Monday to Friday.")
        return
    
    log_to_excel(date, day, time_in.strftime('%H:%M'), time_out.strftime('%H:%M'), total_hours)
    total_label.config(text=f"Total Hours: {total_hours} hrs")
    messagebox.showinfo("Success", "Log saved to Excel.")

def delete_log_popup():
    def perform_delete():
        date_to_delete = date_entry.get()
        if not date_to_delete:
            messagebox.showwarning("Input Needed", "Please enter a Date")
            return
        wb = load_workbook(FILENAME)
        ws = wb.active
        found = False

        for i in range(2, ws.max_row + 1):
            if ws[f"A{i}"].value == date_to_delete:
                ws.delete_rows(i)
                found = True
                break
        if found:
            wb.save(FILENAME)
            messagebox.showinfo("Sucess", f"Deleted Log for {date_to_delete}")
            date_entry.delete(0, tk.END)
        else:
            messagebox.showwarning("Not found", "No log found for that date.")


    popup = Toplevel(root)
    popup.title("Delete Logs")
    popup.geometry("300x150")

    tk.Label(popup, text="Enter date (YYYY-MM-DD):", font=("Arial", 10)).pack(pady=5)
    date_entry = tk.Entry(popup, width=25)
    date_entry.pack(pady=5)

    tk.Button(popup, text="Delete", command=perform_delete).pack(padx=10)

    

def view_log():
    wb = load_workbook(FILENAME)
    ws = wb.active

    logs=""
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[2] and row[3]:
            logs +=f"{row[0]} ({row[1]}): ({row[2]}) - ({row[3]}) | {row[4]} hrs\n"

    if logs == "":
        logs = "No logs yet."

    log_window = Toplevel(root)
    log_window.title("Log Viewer")
    log_window.geometry("400x300")

    Label(log_window, text="OJT Logs", font=("Arial", 14, "bold")).pack(pady=5)
    Label(log_window, text=logs, justify="left", anchor="nw", font=("Courier", 10),wraplength=380).pack(padx=10, pady=10, fill="both", expand=True)


setup_excel()
time_in = None

root = tk.Tk()
root.title("OJT Hours Tracker")

tk.Label(root, text="OJT Hours Tracker", font=("Arial", 16)).pack(pady=10)

time_in_label = tk.Label(root, text="Time In: ---", font=("Arial", 12))
time_in_label.pack()

time_out_label = tk.Label(root, text="Time Out: ---", font=("Arial", 12))
time_out_label.pack()

total_label = tk.Label(root, text="Total Hours: ---", font=("Arial", 12))
total_label.pack(pady=5)

btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)

tk.Button(btn_frame, text="Time In", width=10, command=time_in_action).grid(row=0, column=0, padx=5)
tk.Button(btn_frame, text="Time Out", width=10, command=time_out_action).grid(row=0, column=1, padx=5)
tk.Button(btn_frame, text="Edit Entry", width=10, command=edit_logs).grid(row=0, column=2, padx=5)
tk.Button(btn_frame, text="Delete Entry", width=10, command=delete_log_popup).grid(row=0, column=3, padx=5)
tk.Button(btn_frame, text="View Logs", width=10, command=view_log).grid(row=0, column=4, padx=5)


root.mainloop()
