import base64
import email
import glob
import imaplib
import json
import os
import shutil
import smtplib
import sys
import time
import win32.win32print as win32print
import win32.win32api as win32api
import logging
import random
import re
import autoit
import img2pdf
import asyncio

class PrinterBot:

    ORG_EMAIL   = "@gmail.com"
    FROM_EMAIL  = "" + ORG_EMAIL
    FROM_PWD    = ""
    SMTP_SERVER = "imap.gmail.com"
    SMTP_PORT   = 993

    maindir = os.getcwd()
    attdir = os.path.join(maindir, 'attachments')
    dtdir = os.path.join(maindir, 'data')
    svddir = os.path.join(maindir, 'attachments', 'saved2')
    printerWinName = 'Canon LBP2900'

    def printerTest(self, filename):
        defPrinter = win32print.GetDefaultPrinter()
        print(defPrinter)
        logging.info(defPrinter)
        win32api.ShellExecute(0, "print", filename, defPrinter, ".", 0)
        print("Printing a file " + filename)
        logging.info("Printing a file " + filename)
        while autoit.win_exists(self.printerWinName) !=1:
            print("Window is not here")
            time.sleep(1)
        
        while autoit.win_exists(self.printerWinName) !=0:
            print("Window is still here")
            time.sleep(1)
        
        print("Window is gone")

    def convertToPdf(self):
        for file in glob.glob(self.svddir+'/*.jpg'):
            print('Converting {file} to pdf'.format(file=file))
            logging.info('Converting {file} to pdf'.format(file=file))
            name = file.split('\\')[-1]
            name = re.findall(r'([\w\s]+)\.', file)[0]
            newFileName = os.path.join(self.svddir , name + ".pdf")
            with open(newFileName, "wb") as f:
                f.write(img2pdf.convert(file))
            print('Converted to pdf: {newFileName}'.format(newFileName=newFileName))
            logging.info('Converted to pdf: {newFileName}'.format(newFileName=newFileName))
            os.remove(file)
            print('Removed old file: {file}'.format(file=file))
            logging.info('Removed old file: {file}'.format(file=file))

        for file in glob.glob(self.svddir+'/*.jpeg'):
            print('Converting {file} to pdf'.format(file=file))
            logging.info('Converting {file} to pdf'.format(file=file))
            name = file.split('\\')[-1]
            name = re.findall(r'([\w\s]+)\.', file)[0]
            newFileName = os.path.join(self.svddir , name + ".pdf")
            with open(newFileName, "wb") as f:
                f.write(img2pdf.convert(file))
            print('Converted to pdf: {newFileName}'.format(newFileName=newFileName))
            logging.info('Converted to pdf: {newFileName}'.format(newFileName=newFileName))
            os.remove(file)
            print('Removed old file: {file}'.format(file=file))
            logging.info('Removed old file: {file}'.format(file=file))

    

    def printFile(self):
        self.convertToPdf()

        for file in glob.glob(self.svddir+'/*.pdf'):
            self.printerTest(file)
            for i in range(5,1,-1):
                print("Moving file in... "+str(i))
                logging.info("Moving file in... "+str(i))
                time.sleep(1)
            os.remove(file)
            print("Removed file")
            logging.info("Removed file")

        for file in glob.glob(self.svddir+'/*.gif'):
            self.printerTest(file)
            for i in range(5,1,-1):
                print("Moving file in... "+str(i))
                logging.info("Moving file in... "+str(i))
                time.sleep(1)
            os.remove(file)
            print("Removed file")
            logging.info("Removed file")

    def saveLastAttashment(self):
        logging.info("Saving last attachment")
        mail = imaplib.IMAP4_SSL(self.SMTP_SERVER)
        mail.login('kolotseymax@gmail.com','tylznoooytqtomeq')
        mail.select('inbox')
        print("Connecting to gmail")
        logging.info("Connecting to gmail")

        type, data = mail.search(None, 'FROM', 'sertifikat@komfarm.by')
        mail_ids = data[0]
        print("Messages count: "+ str(len(mail_ids)))
        logging.info("Connecting to gmail")
        id_list = mail_ids.split()
        latest_email_id = int(id_list[-1])  
        print("Saving!")
        logging.info("Saving!")
        with open(os.path.join(self.dtdir, 'lastMail.json'), "w") as f:
            type, data = mail.fetch(str(latest_email_id), '(RFC822)')
            raw = data[0][1]
            raw_email_string = raw.decode('utf-8')
            f.write(json.dumps(raw_email_string))
            print("Saved")
            logging.info("Saved")
            return

    def getExtension(self, fileName):
        try:
            return re.findall(r'=\w{2}([\w]+)\?=',fileName)[0]
        except Exception:
            return  re.findall(r'\.([\w]+)$',fileName)[0]

         

    def saveAttachment(self,email_message, i):
        myFrom = email_message['From']
        if 'sertifikat@komfarm.by' in myFrom:
            for part in email_message.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                fileName = part.get_filename()
                if bool(fileName):
                    rand = random.randrange(0,99999)
                    ext = self.getExtension(fileName)
                    print('Format: ' + ext)
                    if ext.lower() in ['pdf','jpg','jpeg','gif']:
                        filePath = os.path.join(self.svddir, '{i}.{ext}'.format(i=rand, ext = ext))
                        if not os.path.isfile(filePath):
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()
                        print("File saved {i}.{ext}".format(i=rand, ext = ext))
                        logging.info("File saved {i}.{ext}".format(i=rand, ext = ext))

    def read_email_from_gmail(self,lastMail):
        try:
            mails =[]
            mail = imaplib.IMAP4_SSL(self.SMTP_SERVER)
            mail.login('kolotseymax@gmail.com','tylznoooytqtomeq')
            mail.select('inbox')
            print("Connecting to gmail")
            logging.info("Connecting to gmail")

            type, data = mail.search(None, 'FROM', 'sertifikat@komfarm.by')
            mail_ids = data[0]
            print("Messages count: "+ str(len(mail_ids)))
            logging.info("Messages count: "+ str(len(mail_ids)))
            id_list = mail_ids.split()
            first_email_id = int(id_list[0])
            latest_email_id = int(id_list[-1])        

            for i in range(latest_email_id,first_email_id, -1):
                type, data = mail.fetch(str(i), '(RFC822)' )
                raw = data[0][1]
                raw_email_string = raw.decode('utf-8')
                email_message = email.message_from_string(raw_email_string)

                if json.dumps(raw_email_string) == lastMail:
                    print("Coinsidece!")
                    logging.info("Coinsidece!")
                    if i != latest_email_id:
                        print("Saving new Json!")
                        logging.info("Saving new Json!")
                        with open(os.path.join(self.dtdir, 'lastMail.json'), "w") as f:
                            type, data = mail.fetch(str(latest_email_id), '(RFC822)')
                            raw = data[0][1]
                            raw_email_string = raw.decode('utf-8')
                            f.write(json.dumps(raw_email_string))
                            return
                    else:
                        print("Equals to the old Json")
                        logging.info("Equals to the old Json")
                        return

                print("No coinsidence. Saving a file")
                logging.info("No coinsidence. Saving a file")
                self.saveAttachment(email_message, i)


            print(len(mails))
        except Exception as e:
            print(str(e))

    def createDirs(self):
        if os.path.exists(self.attdir) == False:
            os.mkdir(self.attdir)

        if os.path.exists(self.svddir) == False:
            os.mkdir(self.svddir)

        if os.path.exists(self.dtdir) == False:
            os.mkdir(self.dtdir)

    def Start(self):            
        logging.basicConfig(
            filename=self.maindir+"/logs_{dt}.log".format(dt=random.randrange(1,9999)), 
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
        while True:
            print("\nРобот печатальщик\n*****************\n1.Сохранить последнее письмо\n2.Запустить робота\n*****************\n\n")
            ans = input("Выберите пункт:")
            if ans == '1':
                self.saveLastAttashment()
            elif ans =='2':
                lastMail =''
                while True:
                    with open(os.path.join(self.maindir+'/data/', 'lastMail.json'), "r") as f:
                        lastMail = f.read()
                    print("lastMail read. Count of simbols: " + str(len(lastMail)))
                    logging.info("lastMail read. Count of simbols: " + str(len(lastMail)))
                    
                    self.read_email_from_gmail(lastMail)
                    self.printFile()
            else:
                print('Попробуйте еще раз')

async def main():
    try:
        _bot = PrinterBot()
        _bot.createDirs()
        task = asyncio.create_task (_bot.Start())
        await task
        # _bot.convertToPdf()
    except Exception as e:
        logging.error(json.dumps(e))
        print(json.dumps(e))

asyncio.run(main())
