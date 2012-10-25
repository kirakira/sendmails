#!/usr/bin/python
#coding:utf-8

import smtplib
import configparser
import sys
import mimetypes
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def readlist():
  with open('subject', 'r') as fp:
    subject = fp.read()
    fp.close()

  with open('content', 'r') as fp:
    content = fp.read()
    fp.close()

  with open('list', 'r') as fp:
    lines = fp.readlines()
    
    res = []
    declared = False
    for line in lines:
      line = line.strip()
      if len(line) == 0 or line.startswith('#'):
        continue
      if not declared:
        var = [x.strip() for x in line.split(';')]
        declared = True
      else:
        fields = [x.strip() for x in line.split(',')]
        defs = [x.strip() for x in fields[3].split(';')]
        current_subject = subject
        current_content = content
        for i in range(0, len(var)):
          current_subject = current_subject.replace(var[i], defs[i])
          current_content = current_content.replace(var[i], defs[i])
        res.append({'name': fields[0], 'email': fields[1], 'subject': current_subject, 'content': current_content, 'attachments': [x.strip() for x in fields[2].split(';')]})
    return res


def format(name, email):
  return '%s <%s>' % (name, email)


def init(username, password):
  sys.stdout.write('Logging in as %s ..' % username)
  sys.stdout.flush()
  server = smtplib.SMTP('smtp.gmail.com:587')
  server.starttls()
  server.login(username, password)
  print(' finished')
  return server

 
def construct(sender, receiver, subject, message, attachments):
  outer = MIMEMultipart()
  outer['Subject'] = subject
  outer['From'] = sender
  outer['To'] = receiver

  msg = MIMEText(message.encode('utf-8'), 'plain', 'utf-8')
  outer.attach(msg)

  for filename in attachments:
    ctype, encoding = mimetypes.guess_type(filename)
    if ctype is None or encoding is not None:
      ctype = 'application/octet-stream'

    maintype, subtype = ctype.split('/', 1)
    if maintype == 'text':
      fp = open(filename)
      msg = MIMEText(fp.read(), _subtype=subtype)
      fp.close()
    elif maintype == 'image':
      fp = open(filename, 'rb')
      msg = MIMEImage(fp.read(), _subtype=subtype)
      fp.close()
    elif maintype == 'audio':
      fp = open(filename, 'rb')
      msg = MIMEAudio(fp.read(), _subtype=subtype)
      fp.close()
    else:
      fp = open(filename, 'rb')
      msg = MIMEBase(maintype, subtype)
      msg.set_payload(fp.read())
      fp.close()
      encoders.encode_base64(msg)

    if filename.rfind('/') == -1:
      attachname = filename
    else:
      attachname = filename[filename.rfind('/') + 1:]
    msg.add_header('Content-Disposition', 'attachment', filename=attachname)
    outer.attach(msg)

  return outer


def send(server, sender, receiver, mail):
  sys.stdout.write('Sending mail to %s ..' % receiver)
  sys.stdout.flush()
  server.sendmail(sender, [receiver], mail.as_string())
  print(' finished.')


def main():
  config = configparser.ConfigParser()
  config.read('config')
  username = config['account']['mailaddress']
  if len(username) == 0:
    sys.stdout.write('Input your gmail address (sender address): ')
    sys.stdout.flush()
    username = sys.stdin.readline().splitlines()[0]
  password = config['account']['password']
  if len(password) == 0:
    import getpass
    password=getpass.getpass('Input your password for %s: ' % username)
  name = config['account']['name']
  if len(name) == 0:
    sys.stdout.write('Input your name: ')
    sys.stdout.flush()
    name = sys.stdin.readline().splitlines()[0]

  mail_list = readlist()
  print(len(mail_list), 'mails to send')
  fail = 0

  server = init(username, password)
  sender = format(name, username)

  for i in range(0, len(mail_list)):
    receiver = format(mail_list[i]['name'], mail_list[i]['email'])
    try:
      mail = construct(sender, receiver, mail_list[i]['subject'], mail_list[i]['content'], mail_list[i]['attachments'])
      sys.stdout.write('[%d/%d] ' % (i + 1, len(mail_list)))
      send(server, sender, receiver, mail)
    except:
      print('Failed to send mail to %s' % receiver)
      fail = fail + 1

  if fail == 0:
    print('Successfully sent all mails')
  else:
    print('Failed to send %d mails' % fail)

  server.quit()

if __name__ == '__main__':
  main()
