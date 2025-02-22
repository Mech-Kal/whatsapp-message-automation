import tkinter as tk
from tkinter import filedialog, messagebox
from whatsapp_sender import main, validate_date


def select_file(date):
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if file_path:
        main(file_path, date)


# Create the GUI
def create_gui():
    root = tk.Tk()
    root.title("WhatsApp Message Sender")
    root.geometry("400x300")

    # Header label
    label = tk.Label(root, text="Automate WhatsApp Messaging", font=("Helvetica", 16))
    label.pack(pady=10)

    # Instructions
    instructions = tk.Label(
        root,
        text="Select your Excel file to send WhatsApp messages.\nEnsure the file has columns:\n'Student Name', 'Primary Mob', 'Meeting Time', 'Mode of Meeting', 'Meeting Remarks'.",
        justify="center"
    )
    instructions.pack(pady=10)

    # Entry field for Date
    date_label = tk.Label(root, text="Enter Date (dd-mm-yyyy):", font=("Helvetica", 12))
    date_label.pack(pady=5)

    date_entry = tk.Entry(root, font=("Helvetica", 12), width=20)
    date_entry.pack(pady=5)

    # Select button to choose the Excel file
    def start_sending():
        date = date_entry.get()
        if not date:
            messagebox.showerror("Error", "Please enter a date.")
            return

        if not validate_date(date):
            messagebox.showerror("Error", "Invalid date format! Please use dd-mm-yyyy format.")
            return

        select_file(date)

    select_button = tk.Button(
        root,
        text="Select Excel File",
        command=start_sending,
        font=("Helvetica", 12),
        bg="blue",
        fg="white"
    )
    select_button.pack(pady=20)

    # Start the GUI loop
    root.mainloop()


if __name__ == "__main__":
    create_gui()