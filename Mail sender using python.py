import pandas as pd
from tkinter import *
from tkinter import filedialog
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import time

GR = Tk()
GR.title("Mail Sender")
GR.geometry("275x30")
M = Menu(GR)

L = Label(GR,text="Please select an option from menu",fg = "red",font =("Rockwell",12))
L.grid(row = 1 , column = 0, columnspan =2, sticky = "W")

def One_to_one():
    root = Tk()
    root.title("Email Sender")
    root.config(bg="grey")
    

    bg_color = "#f0f0f0"
    label_text_color = "black"
    button_bg_color = "#4CAF50"
    button_text_color = "white"



    def send_mail():
        sender  = "gaurav.rawat4753@gmail.com"
        #pw = "ougg zwiz httr qhuh"

        print("Intiating command")
        pw = p_entry.get()
        subject = subject_entry.get()
        
        receiver = receiver_mail_entry.get()
        
        text_message = msg_e.get("1.0", END)

        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        msg['Subject'] = subject
        msg.attach(MIMEText(text_message, 'plain'))

        file_path = file_entry.get()
        if file_path:
            with open(file_path, 'rb') as attachment:
                part = MIMEApplication(attachment.read(), Name=file_path)
                part['Content-Disposition'] = f'attachment; filename="{file_path}"'
                msg.attach(part)
                
        server = smtplib.SMTP('smtp.gmail.com',587)
        server.starttls()
        server.login(sender,pw)
        print("Login successfull")
        server.sendmail(sender ,receiver , msg.as_string())
        print("Message sent successfully")
        
    


    
    def attach_file():
        
        file_path = filedialog.askopenfilename()
        file_entry.delete(0, END)
        file_entry.insert(0, file_path)
    

    sender_mail = Label(root, text="Sender mail:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    sender_mail.grid(row=1, column=0, padx=10, pady=5, sticky="E")

    sender1 = Label(root, text="gaurav.rawat4753@gmail.com", font=("Rockwell", 12))
    sender1.grid(row=1, column=1, padx=10, pady=5, sticky="W")

    password = Label(root, text="Password:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    password.grid(row=2, column=0, padx=10, pady=5, sticky="E")

    p_entry = Entry(root, font=("Rockwell", 12), fg="#333333", bd=2, relief="solid")
    p_entry.grid(row=2, column=1, padx=10, pady=5, sticky="W")

    receiver_mail = Label(root, text="Recipient mail:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    receiver_mail.grid(row=3, column=0, padx=10, pady=5, sticky="E")

    receiver_mail_entry = Entry(root, width=30, font=("Rockwell", 12), fg="#333333", bd=2, relief="solid")
    receiver_mail_entry.grid(row=3, column=1, padx=10, pady=5, sticky="W")

    subject_label = Label(root, text="Subject:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    subject_label.grid(row=4, column=0, padx=10, pady=5, sticky="E")

    subject_entry = Entry(root, width=30, font=("Rockwell", 12), fg="#333333", bd=2, relief="solid")
    subject_entry.grid(row=4, column=1, padx=10, pady=5, sticky="W")

    message_label = Label(root, text="Message:", font=("Rockwell", 12), bg=bg_color, fg=label_text_color)
    message_label.grid(row=5, column=0, padx=10, pady=5, sticky="E")

    msg_e = Text(root, font=("Rockwell", 12), width=40, height = 5, fg="#333333", bd=2, relief="solid")
    msg_e.grid(row=5, column=1, padx=10, pady=5, sticky="W")

    file_label = Label(root, text="Attached File:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    file_label.grid(row=6, column=0, padx=10, pady=5, sticky="E")

    file_entry = Entry(root, width=30, font=("Rockwell", 12), fg="#333333", bd=2, relief="solid")
    file_entry.grid(row=6, column=1, padx=10, pady=5, sticky="W")

    file_browse_button = Button(root, text="Attach", font=("Rockwell", 12), bd=10, command=attach_file,bg=button_bg_color, fg=button_text_color)
    file_browse_button.grid(row=6, column=2, pady=5, sticky="W")

    send_button = Button(root, text="Send", font=("Rockwell", 16), bd=10, command=send_mail ,bg=button_bg_color, fg=button_text_color)
    send_button.grid(row=7, column=0, pady=20, columnspan=3)

def using_csv():
    root = Tk()
    root.title("Email Sender")
    root.config(bg="grey")
    

    bg_color = "#f0f0f0"
    label_text_color = "black"
    button_bg_color = "#4CAF50"
    button_text_color = "white"

    def send_mails():

        sender  = "gaurav.rawat4753@gmail.com"
        #pw = "ougg zwiz httr qhuh"
        receiver = receiver_mail_entry.get()
        data = pd.read_csv("Book1.csv")

        file_path = path_entry.get()
            
        df = pd.read_csv(file_path)
        email_list = df['ID'].tolist()
        c = 0
        subject = subject_entry.get()
        pw = p_entry.get()
        text_message = msg_e.get("1.0", END)

        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        msg['Subject'] = subject
        msg.attach(MIMEText(text_message, 'plain'))

        for i in email_list:
            server = smtplib.SMTP('smtp.gmail.com',587)
            
            file_path = file_entry.get()

            if file_path:
                with open(file_path, 'rb') as attachment:
                    part = MIMEApplication(attachment.read(), Name=file_path)
                    part['Content-Disposition'] = f'attachment; filename="{file_path}"'
                    msg.attach(part)

            server.starttls()
            server.login(sender,pw)
             
            server.sendmail(sender ,i , msg.as_string())
            print("Mail sent to",df["Name"][c])
            c+=1
            time.sleep(2)


    def attach_file():
        file_path = filedialog.askopenfilename()
        file_entry.delete(0, END)
        file_entry.insert(0, file_path)
        
    def browse_file():
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        path_entry.delete(0, END)
        path_entry.insert(0, file_path)

        
    sender_mail = Label(root, text="Sender mail:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    sender_mail.grid(row=1, column=0, padx=10, pady=5, sticky="E")

    sender1 = Label(root, text="gaurav.rawat4753@gmail.com", font=("Rockwell", 12))
    sender1.grid(row=1, column=1, padx=10, pady=5, sticky="W")

    password = Label(root, text="Password:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    password.grid(row=2, column=0, padx=10, pady=5, sticky="E")

    p_entry = Entry(root, font=("Rockwell", 12), fg="#333333", bd=2, relief="solid")
    p_entry.grid(row=2, column=1, padx=10, pady=5, sticky="W")

    path_label = Label(root, text="CSV File Path:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    path_label.grid(row=3, column=0, padx=10, pady=5, sticky="E")

    path_entry = Entry(root, width=30, font=("Rockwell", 12), fg="#333333", bd=2, relief="solid")
    path_entry.grid(row=3, column=1, padx=10, pady=5, sticky="W")

    browse_button = Button(root, text="Browse", font=("Rockwell", 12), bd=10, command=browse_file,bg=button_bg_color, fg=button_text_color)
    browse_button.grid(row=3, column=2, pady=5, sticky="W")

    subject_label = Label(root, text="Subject:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    subject_label.grid(row=4, column=0, padx=10, pady=5, sticky="E")

    subject_entry = Entry(root, width=30, font=("Rockwell", 12), fg="#333333", bd=2, relief="solid")
    subject_entry.grid(row=4, column=1, padx=10, pady=5, sticky="W")

    message_label = Label(root, text="Message:", font=("Rockwell", 12), bg=bg_color, fg=label_text_color)
    message_label.grid(row=5, column=0, padx=10, pady=5, sticky="E")

    msg_e = Text(root, font=("Rockwell", 12), width=40, height = 5, fg="#333333", bd=2, relief="solid")
    msg_e.grid(row=5, column=1, padx=10, pady=5, sticky="W")

    file_label = Label(root, text="Attached File:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    file_label.grid(row=6, column=0, padx=10, pady=5, sticky="E")

    file_entry = Entry(root, width=30, font=("Rockwell", 12), fg="#333333", bd=2, relief="solid")
    file_entry.grid(row=6, column=1, padx=10, pady=5, sticky="W")

    file_browse_button = Button(root, text="Attach", font=("Rockwell", 12), bd=10, command=attach_file,bg=button_bg_color, fg=button_text_color)
    file_browse_button.grid(row=6, column=2, pady=5, sticky="W")

    send_button = Button(root, text="Send", font=("Rockwell", 16), bd=10, command=send_mails ,bg=button_bg_color, fg=button_text_color)
    send_button.grid(row=7, column=0, pady=20, columnspan=3)

    
def one_to_multiple():
    root = Tk()
    root.title("Email Sender")
    root.config(bg="grey")

    bg_color = "#f0f0f0"
    label_text_color = "black"
    button_bg_color = "#4CAF50"
    button_text_color = "white"
        
    
    def send_mails():
        
        sender  = "gaurav.rawat4753@gmail.com"
        #pw = "ougg zwiz httr qhuh"
        receiver = mail_entry.get()
        input_text = mail_entry.get()
        list_of_items = input_text.split()
        
        subject = subject_entry.get()
        pw = p_entry.get()
        text_message = msg_e.get("1.0", END)

        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        msg['Subject'] = subject
        msg.attach(MIMEText(text_message, 'plain'))

        for i in list_of_items:
            server = smtplib.SMTP('smtp.gmail.com',587)
            
            file_path = file_entry.get()

            if file_path:
                with open(file_path, 'rb') as attachment:
                    part = MIMEApplication(attachment.read(), Name=file_path)
                    part['Content-Disposition'] = f'attachment; filename="{file_path}"'
                    msg.attach(part)

            server.starttls()
            server.login(sender,pw)
             
            server.sendmail(sender ,i , msg.as_string())
            time.sleep(2)
    def attach_file():
        file_path = filedialog.askopenfilename()
        file_entry.delete(0, END)
        file_entry.insert(0, file_path)

        
    sender_mail = Label(root, text="Sender mail:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    sender_mail.grid(row=1, column=0, padx=10, pady=5, sticky="E")

    sender1 = Label(root, text="gaurav.rawat4753@gmail.com", font=("Rockwell", 12))
    sender1.grid(row=1, column=1, padx=10, pady=5, sticky="W")

    password = Label(root, text="Password:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    password.grid(row=2, column=0, padx=10, pady=5, sticky="E")

    p_entry = Entry(root, font=("Rockwell", 12), fg="#333333", bd=2, relief="solid")
    p_entry.grid(row=2, column=1, padx=10, pady=5, sticky="W")

    receiver_mails = Label(root, text="Receiver mails:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    receiver_mails.grid(row=3, column=0, padx=10, pady=5, sticky="E")

    mail_entry=Entry(root, width=30, font=("Rockwell", 12), fg="#333333", bd=2, relief="solid")
    mail_entry.grid(row = 3, column = 1 , padx = 10, pady = 5 , sticky = "W")

    subject_label = Label(root, text="Subject:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    subject_label.grid(row=4, column=0, padx=10, pady=5, sticky="E")

    subject_entry = Entry(root, width=30, font=("Rockwell", 12), fg="#333333", bd=2, relief="solid")
    subject_entry.grid(row=4, column=1, padx=10, pady=5, sticky="W")

    message_label = Label(root, text="Message:", font=("Rockwell", 12), bg=bg_color, fg=label_text_color)
    message_label.grid(row=5, column=0, padx=10, pady=5, sticky="E")

    msg_e = Text(root, font=("Rockwell", 12), width=40, height = 5, fg="#333333", bd=2, relief="solid")
    msg_e.grid(row=5, column=1, padx=10, pady=5, sticky="W")

    file_label = Label(root, text="Attached File:", font=("Rockwell", 12),bg=bg_color, fg=label_text_color)
    file_label.grid(row=6, column=0, padx=10, pady=5, sticky="E")

    file_entry = Entry(root, width=30, font=("Rockwell", 12), fg="#333333", bd=2, relief="solid")
    file_entry.grid(row=6, column=1, padx=10, pady=5, sticky="W")

    file_browse_button = Button(root, text="Attach", font=("Rockwell", 12), bd=10, command=attach_file,bg=button_bg_color, fg=button_text_color)
    file_browse_button.grid(row=6, column=2, pady=5, sticky="W")

    send_button = Button(root, text="Send", font=("Rockwell", 16), bd=10, command=send_mails ,bg=button_bg_color, fg=button_text_color)
    send_button.grid(row=7, column=0, pady=20, columnspan=3)
    
GR.config(menu=M)
SM1 = Menu(M)
M.add_cascade(label = "To Mail",menu=SM1)
SM1.add_cascade(label="One to one",command=One_to_one)
SM1.add_cascade(label="One to multiple",command=one_to_multiple)
SM1.add_cascade(label="Using CSV file",command = using_csv)


GR.mainloop()
