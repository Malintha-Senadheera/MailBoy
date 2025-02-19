### This Is the program for automating email sending ###

# Importing necessary libraries
import win32com.client
import os
import time  # Import time for making delay
import pyfiglet
import subprocess
import sys
from colorama import init, Fore, Style

# This is for changing the color of program text
init()
print(Style.BRIGHT + Fore.GREEN)

# Pyfiglet CLI
result = pyfiglet.figlet_format("Mail Sender")
print(result)

# Function to read email addresses from a text file
def read_emails_from_file(file_path):
    if not os.path.exists(file_path):
        print(f"⚠️ Warning: {file_path} not found.")
        return []
    with open(file_path, "r", encoding="utf-8") as file:
        return [line.strip() for line in file if line.strip()]

# Function to read subject and body from a single HTML file
def read_subject_and_body(file_path):
    if not os.path.exists(file_path):
        print(f"⚠️ Warning: {file_path} not found.")
        return "No Subject", ""  # Return default values if file is missing

    with open(file_path, "r", encoding="utf-8") as file:
        content = file.read()

    # Split subject and body using a delimiter
    parts = content.split("<!-- SUBJECT_END -->", 1)
    
    if len(parts) == 2:
        subject = parts[0].strip()  # Subject is before the delimiter
        body = parts[1].strip()     # Body is after the delimiter
    else:
        subject, body = "No Subject", content  # Fallback if delimiter is missing

    return subject, body

# Function to send an email with inline images and CC
def send_email_with_content(to, cc, subject, body, image_paths):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.CC = cc  # Add CC recipients
    mail.Subject = subject
    mail.BodyFormat = 2  # Ensure HTML Format
    mail.HTMLBody = body  # Set initial email body

    attachments = mail.Attachments
    image_cids = []

    # Attach images for mail body
    for i, image_path in enumerate(image_paths):
        if os.path.exists(image_path):
            attachment = attachments.Add(image_path)

            # Generate a unique content-ID for the image
            cid = f"image{i+1}"

            # Set Content-ID for embedding
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", cid)

            # **Mark as Inline**
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003", 1)

            image_cids.append(cid)
        else:
            print(f"⚠️ Warning: Image not found - {image_path}")

    # Replace placeholders with actual cid references
    for i, cid in enumerate(image_cids):
        body = body.replace(f"{{image{i+1}}}", f"cid:{cid}")

    mail.HTMLBody = body  # Set the final email body
    mail.Send()
    print(f"✅ Email sent to {to} (CC: {cc}) with embedded images!")

# Paths for mail content
email_file = "emails.txt"
cc_file = "cc_emails.txt"
Alutec_email_content_file = "Alutec_Content.html"  # Single file for subject + body
SPK_email_content_file = "SPK_Content.html"  # Single file for subject + body

# Update image paths (ensure these files exist)
image_paths = [
    r"C:\Users\ALUTECH RATHNAPURA\Documents\Malintha\MailBoy\images\image1.jpg",
    r"C:\Users\ALUTECH RATHNAPURA\Documents\Malintha\MailBoy\images\image2.jpg",
    r"C:\Users\ALUTECH RATHNAPURA\Documents\Malintha\MailBoy\images\image3.png",
    r"C:\Users\ALUTECH RATHNAPURA\Documents\Malintha\MailBoy\images\image4.png"
]

# **Menu System**
while True:
    print("\n📧 EMAIL SENDER MENU:")
    print("1. Show the list of emails")
    print("2. Edit the list of emails")
    print("3. Edit the CC mail")
    print("4. Send the emails")
    print("5. Edit mail Body & Subject")
    print("6. Exit")

    # User choice
    choice = input("Enter Your Choice: ").strip()

    # Choice 1: Show Email List
    if choice == "1":
        email_list = read_emails_from_file(email_file)
        if email_list:
            print("\n📜 Email List:")
            for idx, email in enumerate(email_list, start=1):
                print(f"{idx}. {email}")
        else:
            print("\n⚠️ No emails found in emails.txt.")

    # Choice 2: Edit Email List
    elif choice == "2":
        xfile = r"C:\Users\ALUTECH RATHNAPURA\Documents\Malintha\Else\Hotels in ratnapura distric.xlsx"
        os.startfile(xfile)
        subprocess.run(["notepad.exe", email_file])

    # Choice 3: Edit CC Emails
    elif choice == "3":
        subprocess.run(["notepad.exe", cc_file])

    # Choice 4: Send Emails
    elif choice == "4":
        print("Which Section Do You Want To Send Emails?")
        print("1. SPK Outdoors")
        print("2. Alutech Aluminium")
        choice = input("Enter Your Choice: ").strip()
        
        if choice == "1":
            email_list = read_emails_from_file(email_file)
            cc_list = read_emails_from_file(cc_file)
            subject, body_template = read_subject_and_body(SPK_email_content_file)
            cc_recipients = ", ".join(cc_list) if cc_list else ""
            if not email_list:
                print("\n⚠️ No emails to send.")
            else:
                print("\n📩 Emails will be sent to the following addresses:")
                for email in email_list:
                    print(f"➡️ {email}")
                print(f"📌 CC: {cc_recipients if cc_recipients else 'No CC recipients'}")
                confirm = input("\n❓ Are you sure you want to send the emails? (yes/no): ").strip().lower()
                if confirm != "yes":
                    print("❌ Email sending cancelled.")
                    continue
                if len(email_list) > 3:
                    for email in email_list:
                        send_email_with_content(email, cc_recipients, subject, body_template, image_paths)
                        print("⏳ Waiting 1 minute before sending the next email...")
                        time.sleep(60)  # **Wait for 60 seconds before sending the next email**
                else:
                    for email in email_list:
                        send_email_with_content(email, cc_recipients, subject, body_template, image_paths)
                print("\n✅ All emails sent successfully!")
        elif choice == "2":
            email_list = read_emails_from_file(email_file)
            cc_list = read_emails_from_file(cc_file)
            subject, body_template = read_subject_and_body(Alutec_email_content_file)
            cc_recipients = ", ".join(cc_list) if cc_list else ""
            if not email_list:
                print("\n⚠️ No emails to send.")
            else:
                print("\n📩 Emails will be sent to the following addresses:")
                for email in email_list:
                    print(f"➡️ {email}")
                print(f"📌 CC: {cc_recipients if cc_recipients else 'No CC recipients'}")
                confirm = input("\n❓ Are you sure you want to send the emails? (yes/no): ").strip().lower()
                if confirm != "yes":
                    print("❌ Email sending cancelled.")
                    continue
                if len(email_list) > 3:
                    for email in email_list:
                        send_email_with_content(email, cc_recipients, subject, body_template, image_paths)
                        print("⏳ Waiting 1 minute before sending the next email...")
                        time.sleep(60)  # **Wait for 60 seconds before sending the next email**
                else:
                    for email in email_list:
                        send_email_with_content(email, cc_recipients, subject, body_template, image_paths)
                print("\n✅ All emails sent successfully!")
        else:
            print("\n⚠️ Invalid choice. Please enter a number between 1-2.")


    # Choice 5: Edit Email Content
    elif choice == "5":
        print("Which Section Do You Want To Edit Mail Content?")
        print("1. SPK Outdoors")
        print("2. Alutech Aluminium")
        choice = input("Enter Your Choice: ").strip()
        if choice == "1":
            subprocess.run(["notepad.exe", Alutec_email_content_file])
        elif choice == "2":
            subprocess.run(["notepad.exe", SPK_email_content_file])

    # Choice 6: Exit Program
    elif choice == "6":
        exit_confirm = input("\n Are you sure you want to exit? (yes/no): ").strip().lower()
        if exit_confirm == "yes":
            print("\n👋 Exiting program. Goodbye!")
            time.sleep(2.5) # Small delay befor the exit
            sys.exit()

        else:
            print("\n Returning to the menu...")
            continue

    else:
        print("\n⚠️ Invalid choice. Please enter a number between 1-6.")
