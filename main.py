import imapclient, email, imaplib, time
from datetime import date
from datetime import timedelta
import pyzmail, os, re
import pandas as pd

os.chdir(r"E:")

# Get today's date
today = date.today()
YESTERDAY_DATE = today - timedelta(days=1)

df = pd.DataFrame(
    columns=['name', 'email', 'url', 'sent_time', 'subject', 'reply', 'Leads_Details', 'response', 'Contactno'])
df.to_csv("output" + str(today) + ".csv", index=False)


def create_excel(table_details):
    # Put Ayesha's Excel here
    df = pd.read_csv("output" + str(today) + ".csv", index_col=False, header=None)
    df_write_new = pd.DataFrame(table_details).T

    df_write_new.to_csv("output" + str(today) + ".csv", mode='a', index=False, header=None)
    # print(df)
    # df.to_csv("output"+str(today)+".csv",index=False)
    print(len(table_details))
    pass


def get_into_gmail(name, password):
    imapObj = imapclient.IMAPClient('imap.gmail.com', ssl=True)
    # needs to be set from google account- turn on 2 step verification and then apply new apps pass
    imapObj.login(name, password)
    imapObj.select_folder('INBOX', readonly=True)
    return imapObj


def filter_gmail_messages(imapObj):
    # Get Emails from one day before till today
    UIDs = imapObj.search(
        ['SINCE', "14-Nov-2022", 'FROM', 'ayeshasultana100.ron@gmail.com', 'SUBJECT', 'Fwd: an example? --- 47'])
    rawMessages = imapObj.fetch(UIDs, ['BODY[]', 'FLAGS'])
    return rawMessages, UIDs


def specific_filter(rawMessages, UIDs):
    # Removes TO and SUB here
    def remove_SubandTO(each_mail):
        # global sub,to
        # print(each_mail)

        for i, j in enumerate(each_mail):
            if j.startswith('Subject'):
                del each_mail[i]
        for i, j in enumerate(each_mail):
            if j.startswith('To'):
                del each_mail[i]

        # print(each_mail)
        removed_sub_to = [i.strip('\r') for i in each_mail]
        return removed_sub_to

    # Main code
    for i in UIDs:
        message = pyzmail.PyzMessage.factory(rawMessages[i][b'BODY[]'])
        # print(message.get_subject())
        if message.text_part != None:
            each_mail = message.text_part.get_payload().decode(message.text_part.charset).split("\n")
            removed_sub_to = remove_SubandTO(each_mail)
            removed_blanks = [i for i in removed_sub_to if i]

            # print(removed_blanks)
        # print(removed_blanks)
        table_details = []
        # print(len(removed_blanks))

        # Setting x and contact no some string
        x = " - "
        contactno = 'notfound'
        for ind, content in enumerate(removed_blanks):
            print('---', content)
            print()
            # Checks for mobile since its in the second place

            if ind == 2:
                if content.startswith('+'):
                    content = content[1:]

                x = content.replace("-", "")
                x = x.replace(" ", "")
                print(x)
                if x.isdigit():
                    contactno = x

            name = removed_blanks[0]
            regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
            if re.fullmatch(regex, content):
                email = content
                print("email: ", email)
            elif content.startswith('www') and content.endswith('.com'):
                url = content
                print("url: ", url)

            elif "".join(content.replace("-", "")).isdigit():
                print("afafafsa")
                print("mobile", "".join(content.replace("-", "")))
                # isdigit()
            elif content.startswith('*Sent:'):
                sent = "".join(content.split(" ")[1:-1])
                sent_time = sent + "-" + content.split(" ")[-1]
                print("sent_time: ", sent_time)
            elif content.startswith('*Subject:'):
                subject = "".join(content.split(" ")[1:])
                sub_indx = ind + 1
                reply = removed_blanks[sub_indx]
                print("reply: ", reply)
            elif content.startswith('On'):

                leads_name = " ".join(content.split(" "))
                # print('fasf',leads_name)


            elif content.endswith("wrote:"):
                # or content.startswith('reply')
                # print("reasfnajnjfafafsas")
                # print("content",content)

                leads_email_name = leads_name + " " + "".join(content.split(" ")[-2])
                print("Lead: ", leads_email_name)
                # print("**************")

                response = "".join(removed_blanks[ind + 1:])
                # print("response: ",response)
            # print(ind)

        table_details.append([name, email, url, sent_time, subject, reply, leads_email_name, response, contactno])
        print(table_details[0])

        create_excel(table_details[0])

        print()


name = "tejasbiker23@gmail.com"
password = "pbuwtyryrqqrcazg"

imapObj = get_into_gmail(name, password)
rawMessages, UIDs = filter_gmail_messages(imapObj)
specific_filter(rawMessages, UIDs)

import imapclient, email, imaplib, time
from datetime import date
from datetime import timedelta
import pyzmail, os, re
import pandas as pd
import

os.chdir(r"E:")

# Get today's date
today = date.today()
YESTERDAY_DATE = today - timedelta(days=1)

df = pd.DataFrame(
    columns=['name', 'email', 'url', 'sent_time', 'subject', 'reply', 'Leads_Details', 'response', 'Contactno'])
df.to_csv("output" + str(today) + ".csv", index=False)


def create_excel(table_details):
    # Put Ayesha's Excel here
    df = pd.read_csv("output" + str(today) + ".csv", index_col=False, header=None)
    df_write_new = pd.DataFrame(table_details).T

    df_write_new.to_csv("output" + str(today) + ".csv", mode='a', index=False, header=None)
    # print(df)
    # df.to_csv("output"+str(today)+".csv",index=False)
    print(len(table_details))
    pass


def get_into_gmail(name, password):
    imapObj = imapclient.IMAPClient('imap.gmail.com', ssl=True)
    # needs to be set from google account- turn on 2 step verification and then apply new apps pass
    imapObj.login(name, password)
    imapObj.select_folder('INBOX', readonly=True)
    return imapObj


def filter_gmail_messages(imapObj):
    # Get Emails from one day before till today
    UIDs = imapObj.search(
        ['SINCE', "14-Nov-2022", 'FROM', 'ayeshasultana100.ron@gmail.com', 'SUBJECT', 'Fwd: an example? --- 47'])
    rawMessages = imapObj.fetch(UIDs, ['BODY[]', 'FLAGS'])
    return rawMessages, UIDs


def specific_filter(rawMessages, UIDs):
    # Removes TO and SUB here
    def remove_SubandTO(each_mail):
        # global sub,to
        # print(each_mail)

        for i, j in enumerate(each_mail):
            if j.startswith('Subject'):
                del each_mail[i]
        for i, j in enumerate(each_mail):
            if j.startswith('To'):
                del each_mail[i]

        # print(each_mail)
        removed_sub_to = [i.strip('\r') for i in each_mail]
        return removed_sub_to

    # Main code
    for i in UIDs:
        message = pyzmail.PyzMessage.factory(rawMessages[i][b'BODY[]'])
        # print(message.get_subject())
        if message.text_part != None:
            each_mail = message.text_part.get_payload().decode(message.text_part.charset).split("\n")
            removed_sub_to = remove_SubandTO(each_mail)
            removed_blanks = [i for i in removed_sub_to if i]

            # print(removed_blanks)
        # print(removed_blanks)
        table_details = []
        # print(len(removed_blanks))

        # Setting x and contact no some string
        x = " - "
        contactno = 'notfound'
        for ind, content in enumerate(removed_blanks):
            print('---', content)
            print()
            # Checks for mobile since its in the second place

            if ind == 2:
                if content.startswith('+'):
                    content = content[1:]

                x = content.replace("-", "")
                x = x.replace(" ", "")
                print(x)
                if x.isdigit():
                    contactno = x

            name = removed_blanks[0]
            regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
            if re.fullmatch(regex, content):
                email = content
                print("email: ", email)
            elif content.startswith('www') and content.endswith('.com'):
                url = content
                print("url: ", url)

            elif "".join(content.replace("-", "")).isdigit():
                print("afafafsa")
                print("mobile", "".join(content.replace("-", "")))
                # isdigit()
            elif content.startswith('*Sent:'):
                sent = "".join(content.split(" ")[1:-1])
                sent_time = sent + "-" + content.split(" ")[-1]
                print("sent_time: ", sent_time)
            elif content.startswith('*Subject:'):
                subject = "".join(content.split(" ")[1:])
                sub_indx = ind + 1
                reply = removed_blanks[sub_indx]
                print("reply: ", reply)
            elif content.startswith('On'):

                leads_name = " ".join(content.split(" "))
                # print('fasf',leads_name)


            elif content.endswith("wrote:"):
                # or content.startswith('reply')
                # print("reasfnajnjfafafsas")
                # print("content",content)

                leads_email_name = leads_name + " " + "".join(content.split(" ")[-2])
                print("Lead: ", leads_email_name)
                # print("**************")

                response = "".join(removed_blanks[ind + 1:])
                # print("response: ",response)
            # print(ind)

        table_details.append([name, email, url, sent_time, subject, reply, leads_email_name, response, contactno])
        print(table_details[0])

        create_excel(table_details[0])

        print()


name = "your email id "
password = "your set password"

imapObj = get_into_gmail(name, password)
rawMessages, UIDs = filter_gmail_messages(imapObj)
specific_filter(rawMessages, UIDs)


