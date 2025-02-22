# WhatsApp Message Automation App

## Overview
This application automates WhatsApp messaging using **Tkinter** for a user-friendly GUI and **PyWhatKit** for sending messages. It reads contact details from an Excel file and personalizes messages before sending.

## Features
- GUI-based interface for selecting an Excel file
- Validates date format and required columns in Excel
- Filters contacts with empty meeting remarks
- Sends personalized WhatsApp messages instantly
- Uses a predefined message template

---

## Tech Stack
- **Python**
- **Tkinter** (GUI for user interaction)
- **Pandas** (Excel data processing)
- **PyWhatKit** (WhatsApp automation)

---

## Installation & Setup

### 1. **Clone the Repository**
```sh
git clone https://github.com/Mech-Kal/whatsapp-message-automation.git
cd whatsapp-message-automation
```
### 2. Install Dependencies
```sh
pip install -r requirements.txt
```
### 3. Run the Application
```sh
python app.py
```
## Usage
1. Enter the date in dd-mm-yyyy format.
2. Select the Excel file containing student contact details.
3. The script will filter students with empty "Meeting Remarks".
4. Messages are personalized and sent via WhatsApp.
## Excel File Format
The Excel file should contain the following columns:

| Column Name       | Description                              |
|-------------------|------------------------------------------|
| **Student Name**  | Name of the student                     |
| **Primary Mob**   | Mobile number (without country code)    |
| **Meeting Time**  | Scheduled meeting time                  |
| **Mode of Meeting** | Meeting mode (e.g., Online, Offline)  |
| **Meeting Remarks** | Remarks (Empty entries will be processed) |


## Folder Structure
```plaintext
.
├── app.py           # Tkinter-based GUI application
├── whatsapp_sender.py  # Message automation script
├── requirements.txt # Dependencies
├── message_template.txt # Template for WhatsApp messages
└── README.md        # Project documentation
```
## Future Enhancements
- Add support for group messages
- Improve UI with real-time status updates
- Add scheduling functionality for delayed messages
## Author
👤 **Kalpak Kamble**

**GitHub:** Mech-Kal

**Email:** kalpak2002@gmail.com
