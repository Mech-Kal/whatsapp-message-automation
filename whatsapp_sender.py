import pandas as pd
import time
import pywhatkit
from tkinter import messagebox


# Load the Excel file
def load_contacts(file_path):
    try:
        data = pd.read_excel(file_path)
        data.columns = data.columns.str.strip()  # Remove leading/trailing spaces
        return data
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load Excel file: {e}")
        return None


# Load the message template
def load_message(file_path):
    try:
        with open(file_path, 'r') as file:
            return file.read()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load message template: {e}")
        return None


# Send WhatsApp messages instantly
def send_whatsapp_message(phone, message):
    try:
        pywhatkit.sendwhatmsg_instantly(phone, message)
        print(f"Message sent to {phone}")
        time.sleep(5)  # Small delay to avoid sending too quickly
    except Exception as e:
        print(f"Failed to send message to {phone}. Error: {str(e)}")


# Validate if the Excel file contains the required columns
def validate_columns(contacts):
    required_columns = ['Student Name', 'Primary Mob', 'Meeting Time', 'Mode of Meeting', 'Meeting Remarks']
    for column in required_columns:
        if column not in contacts.columns:
            messagebox.showerror("Error", f"Missing required column: {column}")
            return False
    return True


# Function to validate date format (dd-mm-yyyy)
def validate_date(date):
    try:
        time.strptime(date, "%d-%m-%Y")
        return True
    except ValueError:
        return False


# Main script to send messages
def main(file_path, date):
    # Paths to your files
    excel_file = file_path
    message_file = "message_template.txt"

    # Load data
    contacts = load_contacts(excel_file)
    if contacts is None:
        return

    # Validate columns
    if not validate_columns(contacts):
        return

    # Filter contacts where 'Meeting Remarks' is empty
    contacts = contacts[contacts['Meeting Remarks'].isna() | contacts['Meeting Remarks'].str.strip().eq("")]

    if contacts.empty:
        messagebox.showinfo("No Messages", "No contacts found with empty 'Meeting Remarks'.")
        return

    message_template = load_message(message_file)
    if message_template is None:
        return

    # Loop through each contact and send the message
    for index, row in contacts.iterrows():
        name = row['Student Name']
        phone = f"+91{row['Primary Mob']}"  # Add country code if not included
        time_info = row['Meeting Time']  # Time-specific information
        place = row['Mode of Meeting']  # Place-specific information

        # Personalize the message with date included
        message = message_template.format(name=name, time=time_info, place=place, date=date)

        # Send the message instantly
        print(f"Sending message to {name} at {phone}...")
        send_whatsapp_message(phone, message)

    messagebox.showinfo("Success", "All messages sent successfully!")