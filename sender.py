#!/usr/bin/python
# -*- coding: UTF-8 -*-
import time
import shutil
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import subprocess
import os
from shutil import copyfile
from datetime import datetime
from pytz import timezone

sender = 'partslists@specif.io'                 # 固定发件人邮箱
my_pass = '96hue?IprUB!NEp'                     # 发件人邮箱密码


input_docx_path = 'input_temp/input.docx'
output_path = 'output_temp/output.docx'
docplan_name = 'N/A'


def renew_output_dir(filename):
    if os.path.exists(filename):
        shutil.rmtree(filename)
    os.mkdir(filename)


def get_time_file():
    curr_timezone = timezone('America/Los_Angeles')
    utc = timezone('UTC')
    now = datetime.utcnow().replace(tzinfo=utc).astimezone(curr_timezone)

    time = str(now.year) + '_' + str(now.month) + '_' + str(now.day) + '__' \
           + str(now.hour) + '_' + str(now.minute) + '_' + str(now.second)

    return time

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


def word2PartsList(input_fileName):
    cmd = ["java", "-jar", "Word2PartsList.jar", input_docx_path, output_path, input_fileName, docplan_name]
    subprocess.Popen(cmd)

    start = time.time()
    print('Document converting...')
    while not os.path.exists(output_path):
        time.sleep(1)
        end = time.time()
        # print("processing docx... ,spent %.1fs" % (end - start))
        if(end - start) > 600:
            return 'error', 0

    fsize = os.path.getsize(output_path)

    if fsize < 2000:
        if not os.path.exists("error_file_history"):
            os.mkdir("error_file_history")
        copyfile(output_path,
                 'error_file_history/err_output' + get_time_file() + '.docx')
        copyfile(input_fileName,
                 'error_file_history/{}'.format(input_fileName) + get_time_file() + '.docx')

        return 'error', fsize

    # # copy the file from <output_temp> to <output_history>
    # copyfile(output_path, 'output_history/output_' + get_time_file() + '.docx')
    return 'success', fsize


def sending_email(file_path, receivers, text_body, input_fileName):
    # sender ：发件人邮箱账号
    # my_pass ： 发件人邮箱密码
    # receivers ：收件人邮箱账号

    # Instantiate a email massage with attachment
    message = MIMEMultipart()
    message['From'] = Header(sender, 'utf-8')
    message['To'] = Header(receivers, 'utf-8')
    subject = 'Draft Builders :: Parts List Generated :: {}'.format(get_time())
    message['Subject'] = Header(subject, 'utf-8')

    # massage content
    # message.attach(MIMEText(text_body, 'plain', 'utf-8'))
    message.attach(MIMEText(text_body, 'html', 'us-ascii'))

    # Construct attachment1，upload file
    att1 = MIMEText(open(file_path, 'rb').read(), 'base64', 'utf-8')
    att1["Content-Type"] = 'application/octet-stream'
    new_fileName = "Parts List ({}).docx".format(input_fileName)
    att1["Content-Disposition"] = 'attachment; filename="{}"'.format(new_fileName)
    message.attach(att1)

    try:
        server = smtplib.SMTP_SSL("smtp.gmail.com", 465)  # SMTP server of sender，port-num: 25
        server.login(sender, my_pass)  # sender's email address and password
        server.sendmail(sender, [receivers, ], message.as_string())  # sender's email, receiver's email and massage
        server.quit()  # close the connection
        return 'success'

    except smtplib.SMTPException:
        return 'failed'


def send_output(receiver_address, text_body, input_fileName):
    renew_output_dir('output_temp')
    if not os.path.exists(input_docx_path):
        message = "-----Error: unable find input.docx------"
        print(message + '\n')
        write_log("logs/error.log", message)

        return 'convert_error'

    # converting document
    status, fsize = word2PartsList(input_fileName)

    if status == 'success':
        print('\n' + "Finish convert document, document size : {}kb".format(round(fsize/1024, 2)))

        status = sending_email(output_path, receiver_address, text_body, input_fileName)

        if status == 'success':
            print("Sending email success" + '\n')
            return 'success'
        else:
            return 'send_error'
    else:
        return 'convert_error'

# send_output('This the partsList output file.', 'test')    # use for debug and test

