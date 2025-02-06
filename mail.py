import win32com.client
import pythoncom

def get_outlook_folder(folder_name):
    """Retrieve a specific Outlook folder by name."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    # Access the main mailbox folders
    inbox = namespace.GetDefaultFolder(6)  # Inbox folder
    
    # Try to get the custom folder
    try:
        target_folder = inbox.Folders(folder_name)  # Get the specified folder
        return target_folder
    except Exception:
        print(f"Error: Folder '{folder_name}' not found. Using default Inbox.")
        return inbox  # Fall back to Inbox if the folder isn't found

class OutlookAutoForward:
    """Class to handle incoming Outlook emails and forward them based on filters."""

    def __init__(self, folder_name=None, forward_to=None, subject_filter=None, sender_filter=None):
        # Make arguments optional for COM initialization
        self.folder_name = folder_name
        self.forward_to = forward_to
        self.subject_filter = subject_filter
        self.sender_filter = sender_filter

    def configure(self, folder_name, forward_to, subject_filter=None, sender_filter=None):
        # Add a separate method to set the configuration
        self.folder_name = folder_name
        self.forward_to = forward_to
        self.subject_filter = subject_filter
        self.sender_filter = sender_filter

    def OnNewMailEx(self, received_items_ids):
        """Triggered when a new email arrives in Outlook."""
        if not self.forward_to:
            print("No forwarding address set.")
            return
        
        try:
            target_folder = get_outlook_folder(self.folder_name)  # Use custom folder
            messages = target_folder.Items

            for item_id in received_items_ids.split(","):
                mail_item = messages.Item(item_id.strip())
                
                if mail_item and mail_item.Class == 43:  # Ensure it's an email
                    subject = mail_item.Subject
                    sender = mail_item.SenderEmailAddress

                    # Apply filters
                    if self.subject_filter and not any(keyword.lower() in subject.lower() for keyword in self.subject_filter):
                        print(f"Skipping email '{subject}' (does not match subject filter)")
                        continue
                    
                    if self.sender_filter and sender.lower() not in [s.lower() for s in self.sender_filter]:
                        print(f"Skipping email from '{sender}' (does not match sender filter)")
                        continue

                    # Forward the email
                    forward_mail = mail_item.Forward()
                    forward_mail.To = self.forward_to
                    forward_mail.Send()
                    print(f"Forwarded: {mail_item.Subject}")

                    # Mark email as read
                    mail_item.Unread = False
                    mail_item.Save()
        except Exception as e:
            print(f"Error forwarding email: {str(e)}")

def main(forward_to_email, folder_name="Inbox", subject_filter=None, sender_filter=None):
    try:
        # Create the Outlook application with event handling
        OutlookApp = win32com.client.DispatchWithEvents("Outlook.Application", OutlookAutoForward)
        
        # Configure the event handler
        OutlookApp.configure(folder_name, forward_to_email, subject_filter, sender_filter)
        
        print(f"Listening for incoming emails in '{folder_name}'...")
        while True:
            pythoncom.PumpWaitingMessages()  # Keeps the script running
    except Exception as e:
        print(f"Error setting up event listener: {str(e)}")

if __name__ == "__main__":
    forward_to_email = "recipient@example.com"  # Change to your desired email
    custom_folder_name = "CustomFolder"  # Change to your target folder

    # ✅ Set filters (Modify as needed)
    subject_keywords = ["urgent", "invoice"]  # Only forward emails with these words in subject
    allowed_senders = ["boss@example.com", "client@example.com"]  # Only forward emails from these senders

    main(forward_to_email, custom_folder_name, subject_keywords, allowed_senders)
