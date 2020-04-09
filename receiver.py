# coding:utf-8
import re
import imaplib, email, os
from email.header import decode_header
from sender import *
from datetime import datetime
from pytz import timezone


target1 = "From:(.*)@draft.builders"
target2 = "From:(.*)@specif.io"

mail_type = 'imap.gmail.com'
mail_ssl = 993
mail_username = 'partslists@specif.io'
mail_password = '96hue?IprUB!NEp'


def renew_input_dir(filename):
    if os.path.exists(filename):
        shutil.rmtree(filename)
    os.mkdir(filename)


def get_time():
    curr_timezone = timezone('America/Los_Angeles')
    utc = timezone('UTC')
    now = datetime.utcnow().replace(tzinfo=utc).astimezone(curr_timezone)

    time = str(now.year) + '-' + str(now.month) + '-' + str(now.day) + ' ' \
           + str(now.hour) + ':' + str(now.minute) + ':' + str(now.second)

    return time


def write_log(log_path, message):
    with open(log_path, "a") as outputFile:
        outputFile.write(message + " [at time: {}]".format(get_time()) + '\n')
    outputFile.close()


def mail_login(mail_type, mail_ssl, mail_username, mail_password):
    """ Connection and Log-in """
    # connect server
    try:
        conn = imaplib.IMAP4_SSL(mail_type, mail_ssl)
        print('\n' + 'Connection Success...')
    except imaplib.IMAP4_SSL.error:
        message = '!!!!!Error : fail to connect the mail-server!!!!!!'
        print(message + '\n')
        write_log("logs/error.log", message)

        return None

    # login
    try:
        conn.login(mail_username, mail_password)
        print('Log-in Success...')
    except imaplib.IMAP4_SSL.error:
        message = '!!!!!Error : fail to Log-in the account!!!!!'
        print(message + '\n')
        write_log("logs/error.log", message)

        return None

    return conn


def search_email(conn):
    """ Searching unread email from dedicated address """
    conn.select("INBOX")  # default folder is INBOX
    type, mails = conn.search(None, 'UNSEEN')  # SEEN--已读邮件,UNSEEN--未读邮件,ALL--全部邮件

    if mails[0]:
        mails_number = len(mails[0].split())
        print('\n' + "Found {} unread mail".format(mails_number) + '\n')

        irrelevant_mail = 0
        no_att_number = 0
        wrong_format_num = 0
        send_error_num = 0
        convert_error_num = 0

        attachment_num = 0
        processed_num = 0
        # Traverse each email
        for mail_id in range(mails_number):

            resp, data = conn.fetch(mails[0].split()[mail_id], '(RFC822)')
            emailbody = data[0][1]
            mail = email.message_from_bytes(emailbody)

            # email content in string format
            email_content = data[0][1].decode(encoding='UTF-8', errors='strict')

            # match the dedicated email
            match1 = re.search(target1, email_content)
            match2 = re.search(target2, email_content)
            if match1 or match2:
                if match1:
                    sender_address = match1.group(0).split("<")[1]
                else:
                    sender_address = match2.group(0).split("<")[1]

                input_fileName = get_attachment(mail)

                if input_fileName == 'None':    # no attachment
                    no_att_number += 1
                elif input_fileName == 'Wrong':     # wrong attachment
                    wrong_format_num += 1
                else:
                    attachment_num += 1

                    # get text_body
                    text_body = get_text(mail)

                    # send result by email
                    status = send_output(sender_address, text_body, input_fileName)

                    # re-download and re-convert if 'convert_error'
                    restart_count = 0
                    while status == 'convert_error' and restart_count < 3:
                        input_fileName = get_attachment(mail)
                        text_body = get_text(mail)
                        status = send_output(sender_address, text_body, input_fileName)
                        if status == 'convert_error':
                            restart_count += 1
                            print("Error: conversion result is invalid, conversion restart..." + '\n')

                    # process error log
                    if status == 'convert_error':   # unable to convert
                        convert_error_num += 1
                        message = "-----Error: conversion time expired-----"
                        print(message + '\n')
                        write_log("logs/error.log", message)

                    elif status == 'send_error':    # unable to send email
                        send_error_num += 1
                        message = "-----Error: unable to send email-----"
                        print(message + '\n')
                        write_log("logs/error.log", message)

                    else:                           # process successful!
                        processed_num += 1

            else:
                irrelevant_mail += 1


        # success log
        print('\n' + "Finish download {} attachments.".format(attachment_num))
        print("Finish send {} emails.".format(processed_num) + '\n')
        if attachment_num != 0 or processed_num != 0:
            message = "Finish download [{}] attachments and send [{}] emails.".format(attachment_num, processed_num)
            write_log("logs/success.log", message)


        # warning log
        if irrelevant_mail != 0:
            message = "<Warning: Found {} irrelevant mails.>".format(irrelevant_mail)
            print(message)
            write_log("logs/error.log", message)

        if no_att_number != 0:
            message = "<Warning: Found {} mails without attachment.>".format(no_att_number)
            print(message)
            write_log("logs/error.log", message)

        if wrong_format_num != 0:
            message = "<Warning: Found {} mails with incorrect attachment.>".format(wrong_format_num)
            print(message)
            write_log("logs/error.log", message)

        if convert_error_num != 0:
            message = "<Error: Found {} document conversion failed.>".format(convert_error_num)
            print(message)
            write_log("logs/error.log", message)

        if send_error_num != 0:
            message = "<Error: Found {} mails sending failed.>".format(send_error_num)
            print(message)
            write_log("logs/error.log", message)

        return True
    else:
        print('\n' + "No unread mail found" + '\n')
        return False


def get_text(mail):
    for part in mail.walk():
        if part.get_content_maintype() == 'multipart':
            continue

        payload = part.get_payload(decode=True)  # returns a bytes object
        strtext = payload.decode()  # utf-8 is default

        return strtext


def get_attachment(mail):
    renew_input_dir('input_temp')

    name = 'None'
    # 获取邮件附件名称
    for part in mail.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()

        # 如果文件名为纯数字、字母时不需要解码，否则需要解码
        try:
            fileName = decode_header(fileName)[0][0].decode(decode_header(fileName)[0][1])
        except:
            pass
        # 如果获取到了文件，则将文件保存在指定的目录下
        if fileName != 'None':
            name = fileName.split('.')[0]
            format = fileName.split('.')[1]

            # 如果文件的格式错误
            if format != 'docx':
                return 'Wrong'

            # filePath = os.path.join("input_history",
            #                         name + '_' + get_time_file() + '.docx')
            filePath = 'input_temp/input.docx'

            if not os.path.isfile(filePath):
                fp = open(filePath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()

        # # copy the file from <input_history> to <input_temp>
        # copyfile(filePath, 'input_temp/input.docx')

    return name



def process_file():

    conn = mail_login(mail_type, mail_ssl, mail_username, mail_password)

    if conn is not None:
        process_success = search_email(conn)

        if process_success:
            return 'success'

    return 'none'
