import sys
import imapclient
import pyzmail
import datetime as dt
import json
from os.path import exists
from apscheduler.schedulers.blocking import BlockingScheduler

def downloadAttahcedFileFromEmail():
    # read config file
    config = None
    if exists("config.json")==True:
        with open("config.json", "r") as f:
            config = json.load(f)
    else:
        print("You need a config file!")
        sys.exit()
    detach_dir = config["save_dir"]

    # login
    m = imapclient.IMAPClient("imap.gmail.com", ssl=True)
    m.login(config["id"],config["pw"])

    # finding email
    m.select_folder('INBOX')
    UIDs = m.search(['FROM', config["from_filter"]])

    for uid in UIDs:
        raw_msg = m.fetch(uid, ['BODY[]'])
        msg = pyzmail.PyzMessage.factory(raw_msg[8][b'BODY[]'])
        for mp in msg.mailparts:
            if mp.filename != None and mp.filename.find(config['file_ext_filter']) != -1:
                fullpath = detach_dir + mp.filename
                if exists(fullpath) == True:    # rename
                    fullpath = detach_dir + dt.datetime.today().strftime('%Y-%m-%d_%H-%M-%S') +"-"+mp.filename

                # download a attached file
                file = open(fullpath, "wb")
                file.write(mp.get_payload())
                file.close()

                # backup an email
                m.copy(uid, config['move_mail_after_job'])

                #delete the email
                m.delete_messages(uid)
                m.expunge()

scheduler = BlockingScheduler()
scheduler.add_job(downloadAttahcedFileFromEmail, 'interval',hours=2) # per 2 hours
scheduler.start();