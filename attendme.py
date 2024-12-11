import os
import sqlite3
import pandas as pd
from datetime import date
import tkinter as tk
from tkinter import ttk
from ttkbootstrap import Style
from tkinter import messagebox, StringVar

# Constants
DB_NAME = "attendance.db"
ATTENDANCE_OPTIONS = ["Absent", "Here", "Excluded", "Travel"]

# Ensure only Robyn can access
def check_user():
    current_user = os.getlogin()
    authorized_user = "Robyn"  # Replace with Robyn
    if current_user != authorized_user:
        raise PermissionError("Access denied. Only Robyn can access this app.")

# Database setup
def initialize_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_name TEXT NOT NULL,
            date TEXT NOT NULL,
            status TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

# Save attendance to the database
def save_attendance(student_name, attendance_status):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO attendance (student_name, date, status)
        VALUES (?, ?, ?)
    ''', (student_name, date.today().isoformat(), attendance_status))
    conn.commit()
    conn.close()

# Delete an entry from the database
def delete_entry(entry_id):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM attendance WHERE id = ?', (entry_id,))
    conn.commit()
    conn.close()

# Fetch attendance records for a specific day
def fetch_attendance(date_filter=None):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    if date_filter:
        cursor.execute('SELECT id, student_name, date, status FROM attendance WHERE date = ? ORDER BY date DESC', (date_filter,))
    else:
        cursor.execute('SELECT id, student_name, date, status FROM attendance ORDER BY date DESC')
    records = cursor.fetchall()
    conn.close()
    return records

# Export attendance to a CSV file
def export_attendance():
    records = fetch_attendance()
    if not records:
        messagebox.showerror("Error", "No attendance records to export.")
        return
    
    df = pd.DataFrame(records, columns=["ID", "Student Name", "Date", "Status"])
    filename = f"attendance_{date.today().isoformat()}.csv"
    df.to_csv(filename, index=False)
    messagebox.showinfo("Export Successful", f"Attendance exported to {filename}.")

# Tkinter GUI
def create_gui():
    # Apply ttkbootstrap style
    style = Style(theme="flatly")  # Choose a modern, visually appealing theme

    # Main Window
    root = style.master
    root.title("AttendMe - Attendance Tracker")
    root.geometry("900x550")

    # Labels
    label1 = ttk.Label(root, text="Student Name:", font=("Arial", 12))
    label1.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    label2 = ttk.Label(root, text="Attendance Status:", font=("Arial", 12))
    label2.grid(row=1, column=0, padx=10, pady=10, sticky="w")

    label3 = ttk.Label(root, text="Enter Date (YYYY-MM-DD):", font=("Arial", 12))
    label3.grid(row=2, column=0, padx=10, pady=10, sticky="w")

    # Input Fields
    student_name_var = StringVar()
    attendance_status_var = StringVar(value=ATTENDANCE_OPTIONS[0])
    date_var = StringVar(value=date.today().isoformat())

    student_name_entry = ttk.Entry(root, textvariable=student_name_var, width=30)
    student_name_entry.grid(row=0, column=1, padx=10, pady=10)
    
    attendance_status_menu = ttk.OptionMenu(root, attendance_status_var, *ATTENDANCE_OPTIONS)
    attendance_status_menu.grid(row=1, column=1, padx=10, pady=10)

    date_entry = ttk.Entry(root, textvariable=date_var, width=30)
    date_entry.grid(row=2, column=1, padx=10, pady=10)

    # Treeview for Attendance Grid
    columns = ("Date", "Student Name", "Status")
    attendance_tree = ttk.Treeview(root, columns=columns, show="headings", height=15)
    attendance_tree.grid(row=3, column=0, columnspan=4, padx=10, pady=10, sticky="nsew")

    # Define column headings
    for col in columns:
        attendance_tree.heading(col, text=col)
        attendance_tree.column(col, anchor="center", width=200)

    # Buttons
    def add_attendance():
        student_name = student_name_var.get().strip()
        attendance_status = attendance_status_var.get()
        date_input = date_var.get().strip()

        if not student_name or not date_input:
            messagebox.showerror("Error", "Please enter both student name and date.")
            return

        save_attendance(student_name, attendance_status)
        refresh_attendance_list(date_input)
        student_name_var.set("")
        date_var.set(date.today().isoformat())

    def refresh_attendance_list(date_filter=None):
        for row in attendance_tree.get_children():
            attendance_tree.delete(row)
        records = fetch_attendance(date_filter)
        for record in records:
            attendance_tree.insert("", "end", values=record[1:])

    def delete_selected_entry():
        selected_item = attendance_tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select an entry to delete.")
            return
        entry_id = attendance_tree.item(selected_item[0])["values"][0]  # Get ID of the selected entry
        delete_entry(entry_id)
        refresh_attendance_list()

    def export_csv():
        export_attendance()
        refresh_attendance_list()

    add_button = ttk.Button(root, text="Add Attendance", command=add_attendance)
    add_button.grid(row=4, column=0, padx=10, pady=10)
    
    export_button = ttk.Button(root, text="Export to CSV", command=export_csv)
    export_button.grid(row=4, column=1, padx=10, pady=10)

    refresh_button = ttk.Button(root, text="Refresh", command=lambda: refresh_attendance_list(date_var.get()))
    refresh_button.grid(row=4, column=2, padx=10, pady=10)

    delete_button = ttk.Button(root, text="Delete Entry", command=delete_selected_entry)
    delete_button.grid(row=4, column=3, padx=10, pady=10)

    # Configure grid resizing
    root.grid_rowconfigure(3, weight=1)
    root.grid_columnconfigure(0, weight=1)

    # Load initial data
    refresh_attendance_list(date.today().isoformat())
    
    # Run the app
    root.mainloop()

# Main Execution
try:
    check_user()
    initialize_db()
    create_gui()
except PermissionError as e:
    print(e)
except Exception as e:
    messagebox.showerror("Error", f"An unexpected error occurred: {e}")