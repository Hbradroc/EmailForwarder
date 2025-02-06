import win32com.client
import pythoncom

class OutlookAutoForward:
    def __init__(self):
        self.forward_to = None  # Will be set externally

    def OnNewMailEx(self, received_items_ids):
        """Triggered when a new email arrives in Outlook."""
        if not self.forward_to:
            print("No forwarding address set.")
            return
        
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # Inbox folder
            
            # Process each new email
            for item_id in received_items_ids.split(","):
                mail_item = inbox.Items.Item(item_id.strip())
                
                if mail_item and mail_item.Class == 43:  # Ensure it's an email
                    forward_mail = mail_item.Forward()
                    forward_mail.To = self.forward_to
                    forward_mail.Send()
                    print(f"Forwarded: {mail_item.Subject}")

                    # Mark email as read
                    mail_item.Unread = False
                    mail_item.Save()
        except Exception as e:
            print(f"Error forwarding email: {str(e)}")

def main(forward_to_email):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        event_handler = OutlookAutoForward()  # Create an instance
        event_handler.forward_to = forward_to_email  # Set forwarding email
        win32com.client.WithEvents(outlook, OutlookAutoForward)  # Attach event handler
        
        print("Listening for incoming emails...")
        while True:
            pythoncom.PumpWaitingMessages()  # Keeps the script running
    except Exception as e:
        print(f"Error setting up event listener: {str(e)}")

if __name__ == "__main__":
    forward_to_email = "recipient@example.com"  # Change to your desired email
    main(forward_to_email)
