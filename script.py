import pandas as pd
import numpy as np
from PIL import Image, ImageDraw, ImageFont 
from email.mime.text import MIMEText 
from email.mime.image import MIMEImage 
from email.mime.application import MIMEApplication 
from email.mime.multipart import MIMEMultipart 
import smtplib 
import os
import warnings
warnings.filterwarnings("ignore")
smtp = smtplib.SMTP('smtp.gmail.com', 587) 
# smtp = smtplib.SMTP('smtp-mail.outlook.com', 587) 
smtp.ehlo() 
smtp.starttls() 
smtp.login('gdscgcettb@gmail.com', 'bazm dssn vyme jhjf') 
data=pd.read_excel("Copy of Claim Your Swags (Responses).xlsx")
# print(data.keys())
names=list(data.iloc[:,0])
emails=list(data.iloc[:,1])
verified=list(data.iloc[:,2])
status=list(data.iloc[:,3])
# emails=data["Email"]

# print(names)
# print(emails)
print(status)
def message(subject="Python Notification",  
            text="", img=None, 
            attachment=None): 
    
    # build message contents 
    msg = MIMEMultipart() 
      
    # Add Subject 
    msg['Subject'] = subject   
      
    # Add text contents 
    msg.attach(MIMEText(text))   
  
    # Check if we have anything 
    # given in the img parameter 
    if img is not None: 
          
        # Check whether we have the lists of images or not! 
        if type(img) is not list:   
            
              # if it isn't a list, make it one 
            img = [img]  
  
        # Now iterate through our list 
        for one_img in img: 
            
              # read the image binary data 
            img_data = open(one_img, 'rb').read()   
            # Attach the image data to MIMEMultipart 
            # using MIMEImage, we add the given filename use os.basename 
            msg.attach(MIMEImage(img_data, 
                                 name=os.path.basename(one_img))) 
  
    # We do the same for 
    # attachments as we did for images 
    if attachment is not None: 
          
        # Check whether we have the 
        # lists of attachments or not! 
        if type(attachment) is not list: 
            
              # if it isn't a list, make it one 
            attachment = [attachment]   
  
        for one_attachment in attachment: 
  
            with open(one_attachment, 'rb') as f: 
                
                # Read in the attachment 
                # using MIMEApplication 
                file = MIMEApplication( 
                    f.read(), 
                    name=os.path.basename(one_attachment) 
                ) 
            file['Content-Disposition'] = f'attachment;\filename="{os.path.basename(one_attachment)}"' 
              
            # At last, Add the attachment to our message object 
            msg.attach(file) 
    return msg 
  
  

def coupons(names: list, certificate: str, font_path: str): 
   
    for i in range(0,len(names)): 
        if verified[i]=='T' and status[i]!='T':   
            # adjust the position according to  
            # your sample 
            print(f"{i+1}) Sending to {names[i]}....",end="")
            text_y_position = 580 
    
            # opens the image 
            img = Image.open(certificate, mode ='r') 
            
            # gets the image width 
            image_width = img.width 
            
            # gets the image height 
            image_height = img.height  
    
            # creates a drawing canvas overlay  
            # on top of the image 
            draw = ImageDraw.Draw(img) 
    
            # gets the font object from the  
            # font file (TTF) 
            font = ImageFont.truetype( 
                font_path, 
                110 # change this according to your needs 
            ) 
    
            # fetches the text width for  
            # calculations later on 
            text_width, _ = draw.textsize(names[i], font = font) 
    
            draw.text( 
                ( 
                    # this calculation is done  
                    # to centre the image 
                    (image_width - text_width) / 2, 
                    text_y_position
                ), 
                names[i], 
                fill="#000",
                font = font        ) 
    
            # saves the image in png format 
            img.save("./GDSC/{}.png".format(names[i]))  
            msg = message("Congratulations on Completing Gen AI Campaign",
                          f"Dear {names[i]},\n\nCongratulations! You have Successfully completed the Google Generative AI Study Jam by GDSC GCETTB. We are excited to share the certificates with you. Kudos to the hardwork and dedication that you showed throughout the campaign. We hope you learned something new and exciting. \n\nThank you for being a part of GDSC GCETTB. \n\nSwags will be distributed soon. \nStay tuned for further updates.\n\nThank you\nFrom GDSC GCETTB Team.", 
                f"D:/Projects/Code Nexus Certificate Script/GDSC/{names[i]}.png", 
                #   r"C:\Users\Dell\Desktop\slack.py"
                ) 
            smtp.sendmail(from_addr="gdscgcettb@gmail.com", 
                to_addrs=emails[i], msg=msg.as_string()) 
            print("Done")
        else:
           pass
  
        # return img
FONT = "C:/Users/sayan/AppData/Local/Microsoft/Windows/Fonts/AlexBrush-Regular.ttf"
      
    # path to sample certificate 
CERTIFICATE = "D:/Projects/Code Nexus Certificate Script/GenAI.png"
   
coupons(names, CERTIFICATE, FONT) 

smtp.quit()