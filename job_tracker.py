import tkinter as tk
from tkinter import messagebox
import openpyxl
from openpyxl import Workbook
import os
from datetime import datetime
from PIL import Image, ImageTk
import webbrowser

def get_desktop_path():
    return os.path.join(os.path.expanduser("~"), "Desktop")

def save_to_excel(data):
    file_path = os.path.join(get_desktop_path(), "josh_rules.xlsx")
    if os.path.exists(file_path):
        wb = openpyxl.load_workbook(file_path)
        ws = wb["Form Data"]
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Form Data"
        ws.append(["Status", "Date Applied", "Job Title", "Company", "Link (Board)", 
                   "Link (Company)", "Job Description", "Link (Resume)", "Reach Out Person", 
                   "Reach Out Status", "Interview Questions", "Date Submitted"])
    
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # Convert checkbox states to Yes/No
    status_value = "Yes" if data["status"].get() else "No"
    reach_out_status_value = "Yes" if data["reach_out_status"].get() else "No"
    ws.append([status_value, data["date_applied"], data["job_title"], data["company"], 
               data["link_board"], data["link_company"], data["job_description"], 
               data["link_resume"], data["reach_out_person"], reach_out_status_value, 
               data["interview_questions"], timestamp])
    wb.save(file_path)
    messagebox.showinfo("Success", f"Saved at {timestamp}")

def submit_form():
    form_data = {
        "status": status_var,  # Checkbox variable
        "date_applied": date_applied_entry.get(),
        "job_title": job_title_entry.get(),
        "company": company_entry.get(),
        "link_board": link_board_entry.get(),
        "link_company": link_company_entry.get(),
        "job_description": job_description_text.get("1.0", tk.END).strip(),  # Get all text
        "link_resume": link_resume_entry.get(),
        "reach_out_person": reach_out_person_entry.get(),
        "reach_out_status": reach_out_status_var,  # Checkbox variable
        "interview_questions": interview_questions_text.get("1.0", tk.END).strip()  # Get all text
    }
    # Check all fields except checkboxes (which are optional)
    text_fields = {k: v for k, v in form_data.items() if k not in ["status", "reach_out_status"]}
    if not all(text_fields.values()):
        messagebox.showwarning("Warning", "Fill all text fields!")
        return
    save_to_excel(form_data)
    # Clear fields
    status_var.set(0)  # Uncheck
    date_applied_entry.delete(0, tk.END)
    job_title_entry.delete(0, tk.END)
    company_entry.delete(0, tk.END)
    link_board_entry.delete(0, tk.END)
    link_company_entry.delete(0, tk.END)
    job_description_text.delete("1.0", tk.END)
    link_resume_entry.delete(0, tk.END)
    reach_out_person_entry.delete(0, tk.END)
    reach_out_status_var.set(0)  # Uncheck
    interview_questions_text.delete("1.0", tk.END)

# Function to open webpage on image click
def open_webpage(event):
    webbrowser.open("https://www.skool.com/cyber-range/about?ref=cc61b1b3cb11431b889d57956597cce5")

# Create window
root = tk.Tk()
root.title("Josh's Job Tracker")
root.geometry("800x600")
root.configure(bg="#333333")  # Dark grey background

# Create a frame for the form content
form_frame = tk.Frame(root, bg="#333333")
form_frame.pack(expand=True)

# Form fields
# Status (Checkbox)
status_var = tk.IntVar(value=0)  # Default unchecked (0 = off)
tk.Checkbutton(form_frame, text="Status (Applied)", variable=status_var, bg="#333333", fg="#FFFFFF", 
               selectcolor="#D3D3D3").pack(pady=5)

tk.Label(form_frame, text="Date Applied:", bg="#333333", fg="#FFFFFF").pack(pady=5)
date_applied_entry = tk.Entry(form_frame, width=40, bg="#D3D3D3", fg="#FFFFFF", font=("Arial", 12))
date_applied_entry.pack(pady=5)

tk.Label(form_frame, text="Job Title:", bg="#333333", fg="#FFFFFF").pack(pady=5)
job_title_entry = tk.Entry(form_frame, width=40, bg="#D3D3D3", fg="#FFFFFF", font=("Arial", 12))
job_title_entry.pack(pady=5)

tk.Label(form_frame, text="Company:", bg="#333333", fg="#FFFFFF").pack(pady=5)
company_entry = tk.Entry(form_frame, width=40, bg="#D3D3D3", fg="#FFFFFF", font=("Arial", 12))
company_entry.pack(pady=5)

tk.Label(form_frame, text="Link (Board):", bg="#333333", fg="#FFFFFF").pack(pady=5)
link_board_entry = tk.Entry(form_frame, width=40, bg="#D3D3D3", fg="#FFFFFF", font=("Arial", 12))
link_board_entry.pack(pady=5)

tk.Label(form_frame, text="Link (Company):", bg="#333333", fg="#FFFFFF").pack(pady=5)
link_company_entry = tk.Entry(form_frame, width=40, bg="#D3D3D3", fg="#FFFFFF", font=("Arial", 12))
link_company_entry.pack(pady=5)

# Job Description (Memo Field)
tk.Label(form_frame, text="Job Description:", bg="#333333", fg="#FFFFFF").pack(pady=5)
job_description_text = tk.Text(form_frame, width=40, height=4, bg="#D3D3D3", fg="#FFFFFF", font=("Arial", 12))
job_description_text.pack(pady=5)

tk.Label(form_frame, text="Link (Resume):", bg="#333333", fg="#FFFFFF").pack(pady=5)
link_resume_entry = tk.Entry(form_frame, width=40, bg="#D3D3D3", fg="#FFFFFF", font=("Arial", 12))
link_resume_entry.pack(pady=5)

tk.Label(form_frame, text="Reach Out Person:", bg="#333333", fg="#FFFFFF").pack(pady=5)
reach_out_person_entry = tk.Entry(form_frame, width=40, bg="#D3D3D3", fg="#FFFFFF", font=("Arial", 12))
reach_out_person_entry.pack(pady=5)

# Reach Out Status (Checkbox)
reach_out_status_var = tk.IntVar(value=0)  # Default unchecked
tk.Checkbutton(form_frame, text="Reach Out Status (Contacted)", variable=reach_out_status_var, 
               bg="#333333", fg="#FFFFFF", selectcolor="#D3D3D3").pack(pady=5)

# Interview Questions (Memo Field)
tk.Label(form_frame, text="Interview Questions:", bg="#333333", fg="#FFFFFF").pack(pady=5)
interview_questions_text = tk.Text(form_frame, width=40, height=4, bg="#D3D3D3", fg="#FFFFFF", font=("Arial", 12))
interview_questions_text.pack(pady=5)

tk.Button(form_frame, text="Submit", bg="#0000FF", fg="#000000", command=submit_form).pack(pady=20)

# Load and resize the image
image_path = os.path.join(get_desktop_path(), "graphic.jpg")
if os.path.exists(image_path):
    try:
        image = Image.open(image_path)
        original_width, original_height = image.size
        new_width = int(original_width * 0.25)
        new_height = int(original_height * 0.25)
        resized_image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(resized_image)
        image_label = tk.Label(root, image=photo, bg="#333333", cursor="hand2")
        image_label.place(x=10, y=550, anchor="sw")
        image_label.image = photo
        image_label.bind("<Button-1>", open_webpage)
    except Exception as e:
        print(f"Error loading image: {type(e).__name__}: {str(e)}")
else:
    print("Image file not found at the specified path")

root.mainloop():
