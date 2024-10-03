import os
import win32com.client
from datetime import datetime
import PyPDF2

# Create a connection to the Outlook application
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# List all available Outlook accounts/mailboxes
print("Available accounts:")
for i, account in enumerate(namespace.Folders):
    print(f"{i}: {account.Name}")

# Choose a specific account by index (you can change this index to select a different account)
account_index = int(input("Enter the index of the account you want to scan: "))
selected_account = namespace.Folders[account_index]

# Select the Inbox folder for the chosen account
sent = selected_account.Folders["Sent Items"]

# Define the folder where the PDFs will be saved temporarily for scanning
download_folder = r'C:\folder\folder'

# Ensure the download folder exists
if not os.path.exists(download_folder):
    os.makedirs(download_folder)

# Define the date range filter (from June 6, 2024, until today)
start_date = datetime(2024, 6, 6).strftime("%m/%d/%Y")
end_date = datetime.now().strftime("%m/%d/%Y")  # Current date

# Function to check if a PDF contains specific words
def pdf_contains_keywords(pdf_path, keywords):
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ""
            for page in reader.pages:
                text += page.extract_text()

            # Check for the keywords in the extracted text
            for keyword in keywords:
                if keyword.lower() in text.lower():
                    return True
        return False
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
        return False

# Function to process messages in a folder
def process_folder(messages):
    # Apply the restriction to get emails from the start_date to the end_date
    filter_condition = f"[ReceivedTime] >= '{start_date}' AND [ReceivedTime] <= '{end_date}'"
    restricted_messages = messages.Restrict(filter_condition)
    
    for message in restricted_messages:
        print(f"Checking email: {message.Subject}")
        # Loop through the attachments
        for attachment in message.Attachments:
            if attachment.FileName.lower().endswith(".pdf"):
                print(f"Found PDF attachment: {attachment.FileName} in email: {message.Subject}")
                
                # Save the PDF attachment to the folder temporarily
                file_path = os.path.join(download_folder, attachment.FileName)
                attachment.SaveAsFile(file_path)
                
                # Check if the PDF contains the words "authorised" or "authorisation"
                if pdf_contains_keywords(file_path, ["invoice"]):
                    print(f"The PDF '{attachment.FileName}' contains the word 'invoice'.")
                else:
                    # If the PDF does not contain the words, delete it
                    os.remove(file_path)
                    print(f"The PDF '{attachment.FileName}' does not contain the keywords. Deleted.")

# Process the selected account's Inbox
process_folder(sent.Items)

print("Done!")
