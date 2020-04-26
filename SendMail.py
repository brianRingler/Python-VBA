import win32com.client as win32

def send_email_outlook(subject, email_body,path=""):

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'ringler.brian@gmail.com'
    mail.Subject = subject
    mail.Body = email_body
    #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

    # To attach a file to the email (optional):
    attachment = path
    mail.Attachments.Add(attachment)

    mail.Send()

