import cv2
import numpy as np
from PIL import ImageFont, ImageDraw, Image
import os
import pandas as pd
import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import easygui
import logging

coords =[0,0]

def click_event(event, x, y, flags, params):
    
    if event == cv2.EVENT_LBUTTONDOWN:
        print(x, ' ', y)
        coords[0]=x
        coords[1]=y
        cv2.namedWindow('image', cv2.WINDOW_NORMAL)
        cv2.resizeWindow('image', 1020, 768)
        cv2.circle(img,(x,y),4,(255,0,0),2)
        cv2.imshow('image', img)




      
    if event == cv2.EVENT_RBUTTONDOWN:   
        print(x, ' ', y)
        coords[0] = x
        coords[1] = y
        cv2.namedWindow('image',cv2.WINDOW_NORMAL)
        cv2.resizeWindow('image',1020,768)
        cv2.circle(img,(x,y),4,(0,0,255),2)
        cv2.imshow('image', img)



def clean():
    x = easygui.boolbox("Enter 1 to clean 'result' directory: ", choices=["[Y]es","[N]o"], default_choice="[Y]es",cancel_choice="[N]o")
    if x:
        print("Cleaning 'result' folder ........")
        for certificates in os.listdir("result"):
            os.remove("result/{}".format(certificates))
        print("completed.......")
    else:
        msg = "USER: " + os.getlogin() + ", " + "Exited, user selected NO -> cleaning the directory."
        logging.info(msg)
        exit(0)


def verify(emails,names,files):
    if not (len(emails) == len(names) == len(files)):
        return False

    for name,file in zip(names,files):
        if name.strip() != file[:file.index(".")].strip():
            return False
    return True




def certificate_gen(name_ls,location):
        for name in name_ls:
            x = coords[0]
            y = coords[1]
            #1253, 660
            template = cv2.imread(location)
            template_conv = cv2.cvtColor(template, cv2.COLOR_BGR2RGB)
            arr_img = Image.fromarray(template_conv)
            var_draw = ImageDraw.Draw(arr_img)
            pacifico = ImageFont.truetype("fonts/GreatVibes-Regular.ttf", 120)
            var_draw.text((x,y), name, font=pacifico, fill='black')
            final_res = cv2.cvtColor(np.array(arr_img), cv2.COLOR_RGB2BGR)
            cv2.imwrite("result/{}.jpg".format(name), final_res)
            print("{}'s certificate generated".format(name))


def send_Mail(names_ls, emails, filenames, mail_subject, mail_body):
    try:
        os.chdir("result")
    except Exception as e:
        msg = "USER: " + os.getlogin() + ", " + e
        logging.error(msg)
        exit(0)

    sender_email = easygui.enterbox("Enter sender's email address: ")
    password = easygui.passwordbox("Type your password and press enter: ")
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        srno = 0
        for i in range(len(names_ls)):
            name = names_ls[i]
            email = emails[i]
            filename = filenames[i]
            subject = mail_subject
            body = mail_body.format(name)
            receiver_email = email
            message = MIMEMultipart()
            message["From"] = sender_email
            message["To"] = receiver_email
            message["Subject"] = subject
            message.attach(MIMEText(body, "plain"))


        # Open PDF file in binary mode
            with open(filename, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

            encoders.encode_base64(part)

            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {filename}",
            )

            message.attach(part)
            text = message.as_string()
            server.sendmail(sender_email, receiver_email, text)
            print("{} Mail sent to name: {}\t on email: {}\t with file name: {}\t successfully.".format(srno,name,email,filename))
            srno += 1



if __name__ == "__main__":

    logging.basicConfig(level=logging.INFO, filename="genLogs.log", filemode="a+", format="%(asctime)s %(levelname)s %(message)s")
    clean()
    dirlist = os.listdir(".")
    try:
        location = easygui.choicebox("Choose from menu (template image)(.jpg file): ", choices=dirlist)
        hasCoords = easygui.boolbox("Do you have coordinates?",title="Coordinates?",choices=["[Y]es","[N]o"],default_choice="[Y]es",cancel_choice="[N]o")
        if hasCoords:
            coordinates = easygui.multenterbox("Enter Coordinates","Coordinates",["X-Coords: ","Y-Coords: "])
            try:
                coords[0] = int(coordinates[0])
                coords[1] = int(coordinates[1])
            except TypeError as te:
                msg = "USER: " + os.getlogin() + ", " + str(te)
                logging.error(msg)
                exit(0)
        else:
            msg_str = '''
            !!! PLEASE READ !!!
            Please left-click where you'd like to insert your text. 
            A new window with a marked circle will open at
            the selected position. 
            If you're satisfied, close all the windows; 
            otherwise, close the second window 
            and repeat the process by left-clicking on the image. 
            When finished, close all windows.
            '''
            easygui.msgbox(msg_str,"INSTRUCTIONS")
            cv2.namedWindow("template", cv2.WINDOW_NORMAL)
            cv2.resizeWindow("template", 1020, 768)
            img = cv2.imread(location)
            cv2.imshow("template",img)
            cv2.setMouseCallback('template',click_event)
            cv2.waitKey(0)
            cv2.destroyAllWindows()
            easygui.msgbox(f"Your coordinates are:\nX:{coords[0]}\tY:{coords[1]}\n","Coordinates")
        choice = easygui.choicebox("*MAKE SURE THE COLUMNS ARE ALPHABETICALLY SORTD BY NAMES*\nChoose from menu (excel file): ",choices=dirlist)
        df = pd.read_excel(choice)
        cols_list = df.columns.tolist()
        colName = easygui.choicebox("Please select the column name: ",choices=cols_list)
        name_ls = df[colName].tolist()
        certificate_gen(name_ls,location)
        easygui.msgbox(f"{len(name_ls)} Certificates Generated successfully","Generation Completed")
    except Exception as e:
        msg = "USER: " +os.getlogin()+", "+ str(e)
        logging.error(msg)
        print(e)
        exit(0)

    try:
        opt = easygui.boolbox("Do you want to send email? (Please verify certificates before sending email)",choices=["[Y]es","[N]o"],default_choice="[Y]es", cancel_choice="[N]o")
        if opt:
            cName =  easygui.choicebox("Please select the column name: ", choices=cols_list)
            emails = df[cName].tolist()
            filenames = os.listdir("result")
            mail_subject = easygui.textbox("Enter the Subject","Email: Subject",text="Your Event Participation Certificate Is Here!")
            easygui.msgbox("Please put {} in the email body where ever you want to mention NAME of the Participant. Default format is given for you","INSTRUCTION")
            mail_body = easygui.textbox("Enter mail body","Email: Body",text="Dear {},\nWe would like to express our sincere gratitude for your active participation in the event. We are pleased to attach your participation certificate to this email.\nWe are looking forward for your partipation in our upcoming events.\nWarm Regards,\nEVENT TEAM")
            if not verify(emails,name_ls,filenames):
                raise ValueError("Please check the names in excel if they are alphabetically sorted.")
            send_Mail(name_ls,emails,filenames,mail_subject,mail_body)
            easygui.msgbox(f"{len(emails)} mails sent successfully","Mail sent")
        else:
            msg = "USER: " + os.getlogin() + ", " + "Exited, selected NO -> sending email."
            logging.info(msg)
            exit(0)
    except Exception as e:
        easygui.msgbox("Maybe the username or password is wrong or you are using personal mail.\nMake sure you are using institutional/organizational\n mail with \"allow less secure apps\" turned on","Error")
        msg = "USER: " + os.getlogin() + ", " + str(e)
        logging.error(msg)
        print(e)
        exit(0)
