
import json
import win32com.client

#Function to extract SMTP addresses from recipients
def get_recipients(email):
    to_emails = []
    cc_emails = []
    bcc_emails = []

    for recipient in email.Recipients:
        try:
            address_entry = recipient.AddressEntry
            if not address_entry:
                print(f"Recipient {recipient.Name} has no address entry.")
                continue

            email_address = None

            # Process Exchange address format  
            if address_entry.Type == "EX":
                try:
                    exchange_user = address_entry.GetExchangeUser()  # Attempt to resolve
                    email_address = exchange_user.PrimarySmtpAddress if exchange_user else None
                except Exception as ex:
                    print(f"Failed to resolve Exchange user for {recipient.Name}: {ex}")

            # Fallback for non-Exchange (e.g., SMTP)
            if not email_address:
                email_address = getattr(address_entry, "Address", "Unknown")

            # Categorize based on recipient type
            if recipient.Type == 1:
                to_emails.append(email_address)
            elif recipient.Type == 2:
                cc_emails.append(email_address)
            elif recipient.Type == 3:
                bcc_emails.append(email_address)

        except AttributeError as ae:
            print(f"Attribute error for recipient {recipient.Name}: {ae}")
        except Exception as e:
            print(f"Unexpected error for {recipient.Name if hasattr(recipient, 'Name') else 'Unknown'}: {e}")

    return to_emails, cc_emails, bcc_emails



#Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#Access the Inbox (Folder 6 is the default Inbox)
inbox = outlook.GetDefaultFolder(6)

#Get all emails and sort them by received time (newest first)
emails = inbox.Items
emails.Sort("[ReceivedTime]", True)

ds = []
email = emails.GetFirst()

while email:  
    #Check if the item is a valid email
    if not hasattr(email, "Sender"):
        print(f"Skipping item {count + 1}: Not an email message.")
    else:
        try:
            #Try to get proper sender email address
            try:
                exchange_user = email.Sender.GetExchangeUser()
                sender_email = exchange_user.PrimarySmtpAddress if exchange_user else email.SenderEmailAddress
            except Exception as e:
                print(f"Error retrieving sender email: {e}")
                sender_email = email.SenderEmailAddress

            #Get recipients
            to_smtp, cc_smtp, bcc_smtp = get_recipients(email)

            #convert ReceivedTime
            dt_received = str(email.ReceivedTime).split(".")[0]

            #Store email properties in JSON structure
            email_data = {
                "Subject": email.Subject,
                "SenderName": email.SenderName,
                "SenderEmailAddress": sender_email,  
                "To": "; ".join(to_smtp),
                "CC": "; ".join(cc_smtp),
                "BCC": "; ".join(bcc_smtp),
                "Received": dt_received,
                "Body": email.Body
            }
            ds.append(email_data)

        except Exception as e:
            print(f"Error processing email {count + 1}: {e}")

    email = emails.GetNext()


print("Found", len(ds), "emails")

# Save to a JSON file
with open("output/emails.json", "w", encoding="utf-8") as file:
    json.dump(ds, file, indent=4)

print("Saved to emails.json")

