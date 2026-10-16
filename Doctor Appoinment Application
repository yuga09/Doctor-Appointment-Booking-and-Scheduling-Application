# Doctor-Appointment-Booking-and-Scheduling-Application
import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import os


if not os.path.exists("appointments.xlsx"):
    wb = Workbook()
    sheet = wb.active
    sheet.append(["Patient Name", "Doctor", "Date", "Time"])
    wb.save("appointments.xlsx")

# Book appointment function
def book_appointment():
    name = name_entry.get()
    doctor = doctor_entry.get()
    date = date_entry.get()
    time = time_entry.get()

  if name == "" or doctor == "" or date == "" or time == "":
        messagebox.showerror("Error", "Please fill all details")
    else:
        wb = load_workbook("appointments.xlsx")
        sheet = wb.active

  sheet.append([name, doctor, date, time])
        wb.save("appointments.xlsx")

   messagebox.showinfo("Success", "Appointment booked!")

   name_entry.delete(0, tk.END)
   doctor_entry.delete(0, tk.END)
   date_entry.delete(0, tk.END)
   time_entry.delete(0, tk.END)


root = tk.Tk()
root.title("Doctor Appointment Booking System")
root.geometry("400x350")


heading = tk.Label(
    root,
    text="Doctor Appointment Booking",
    font=("Arial", 16, "bold")
)
heading.pack(pady=15)

# Patient name
tk.Label(root, text="Patient Name").pack()
name_entry = tk.Entry(root)
name_entry.pack()

# Doctor name
tk.Label(root, text="Doctor Name").pack()
doctor_entry = tk.Entry(root)
doctor_entry.pack()


tk.Label(root, text="Appointment Date").pack()
date_entry = tk.Entry(root)
date_entry.pack()

tk.Label(root, text="Appointment Time").pack()
time_entry = tk.Entry(root)
time_entry.pack()


book_button = tk.Button(
    root,
    text="Book Appointment",
    command=book_appointment
)
book_button.pack(pady=20)


root.mainloop()
