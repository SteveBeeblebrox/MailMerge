from argparse import ArgumentParser
import win32com.client as win32
from textwrap import wrap
from csv import DictReader
import os
from os import path
from sys import exit, stderr
import re as regex

def info(msg: str):
    print('\x1b[34m[Info]\x1b[39m: %s' % msg, file=stderr)

def error(msg: str):
    print('\x1b[31m[Error]\x1b[39m: %s' % msg, file=stderr)
    exit(1)

def confirm(msg: str, default: bool = False) -> bool:
    t = input('\x1b[32m[Please Confirm]\x1b[39m: %s (%s) ' % (msg, 'Y/n' if default else 'y/N')).lower()
    if t in ['y', 'yes']:
        return True
    elif t in ['n', 'no']:
        return False
    else:
        return default
    
def template(text: str, args: dict) -> str:
    def f(match):
        key = match.group(1).upper()
        if key in args and args[key] is not None:
            return args[key]
        return match.group(0)
    return regex.sub('\\[\\[(.+?)\\]\\]', f, text, 0, regex.DOTALL)
    
    return text

def isEmailAddress(text: str) -> bool:
    return regex.compile('^\\S+@\\S+\\.\\S+$').match(text)

def isEmailAddressList(text: str) -> bool:
    return all(isEmailAddress(t.strip()) for t in text.replace(',', ';').split(';'))

print('Trin\'s Super Mail Merge Script', file=stderr)
print(file=stderr)

parser = ArgumentParser(description='''
This is an improved mail merge utility for Outlook. Note that the recent web-app version won't work.
You'll need the classic version from https://support.microsoft.com/en-us/office/install-or-reinstall-classic-outlook-on-a-windows-pc-5c94902b-31a5-4274-abb0-b07f4661edf5.

To start, make a '.txt' or basic '.html' file that you want to be the contents of the messages. Replace dynamic values with '[[KEY]]' where 'KEY' is some identifier you pick like 'Last Name' or 'Company'.
These keys are case insensitive and should consist of only alphanumeric characters plus underscores and spaces.

Next, make a '*.csv' file (you can export an Excel sheet as this) where each of your keys is a column (without any square brackets or quotes) and each row represents a new message.
                        
The '[[KEY]]' in the template will be replaced by the row and 'KEY' pair in the table. Additionally, keys are substituted in the subject line field.
                        
The keys 'ADDRESS', 'CC', 'BCC', and 'ATTACHMENTS' are special. They indicate the target recipients and any files to attach. Each accept a list of entries separated by ';' and addresses can use contact shorthands. 'ADDRESS' is always required.

Nothing is sent until you confirm. In most cases, user errors are caught before anything is sent; however, Outlook errors or network issues may interrupt partway through sending the list of messages.
''')
parser.add_argument('TEMPLATE', type=str, help='The template to use for message (Accepts TXT or HTML file path)')
parser.add_argument('VALUES', type=str, help='The values to substitute into the template (Accepts CSV file path)')
parser.add_argument('-S', '--subject', dest='SUBJECT', default='', type=str, help='The subject line (Accepts text, Default is empty)')
parser.add_argument('-F', '--from', dest='ADDRESS', type=str, help='The address to send messages from (Accepts email address, Default depends on mail config)')
parser.add_argument('-A', '--folder', dest='FOLDER', type=str, help='Where to look for attachments (Accepts folder path, Default is current folder)', default=os.getcwd())
args = parser.parse_args()

# Validate from address
if args.ADDRESS is not None:
    if isEmailAddress(args.ADDRESS):
        info('Sending as \'%s\'' % args.ADDRESS)
    else:
        error('Provided from address \'%s\' is invalid!' % args.ADDRESS)
else:
    info('Sending as default address')

# Fix up subject line
args.SUBJECT = args.SUBJECT or ''
info('Template for message subject is \'%s\'' % args.SUBJECT)

# Fix up folder path
args.FOLDER = path.abspath(args.FOLDER)
info('Attachments will be loaded from \'%s\'' % args.FOLDER)

# Validate template
if args.TEMPLATE is not None:
    args.TEMPLATE = path.abspath(args.TEMPLATE)
    if args.TEMPLATE.lower().endswith('.txt'):
        args.html = False
        info('Using template \'%s\' as plain text' % args.TEMPLATE)
    elif args.TEMPLATE.lower().endswith('.html'):
        args.html = True
        info('Using template \'%s\' as HTML' % args.TEMPLATE)
    else:
        error('Provided template is not a TXT or HTML file!')
    try:
        with open(args.TEMPLATE) as f:
            args.TEMPLATE = f.read()
    except:
        error('Unable to open \'%s\'', args.TEMPLATE)
else:
    error('Provided template does not exist!')

# Validate values
if args.VALUES is not None and args.VALUES.endswith('.csv'):
    args.VALUES = path.abspath(args.VALUES)
    info('Using values \'%s\'' % args.VALUES)
    try:
        with open(args.VALUES) as f:
            args.VALUES = [{k.upper(): v.strip() if v else v for k, v in row.items()} for row in DictReader(f, skipinitialspace=True)]
    except:
        error('Unable to parse \'%s\'', args.VALUES)
    # info('Found %d recipient(s)' % len(args.VALUES))
else:
    error('Provided values are not a CSV file!')

# Connect to Outlook
try:
    outlook = win32.Dispatch('Outlook.Application')
except Exception as e:
    error('Unable to connect to Outlook (%s)' % e)

# Build Messages
messages = []
for n,row in enumerate(args.VALUES):
    n=n+1
    try:
        msg = outlook.CreateItem(0)
        msg.Sender = args.ADDRESS
        msg.Subject = template(args.SUBJECT, row)

        if 'ADDRESS' in row and row['ADDRESS'] is not None:
            if isEmailAddressList(row['ADDRESS']):
                msg.To = row['ADDRESS'].replace(',', ';')
            else:
                error('Invalid \'ADDRESS\' value \'%s\' in row %d' % (row['ADDRESS'], n))
        else:
            error('Row %d is missing \'ADDRESS\' value' % n)

        if 'CC' in row and row['CC'] is not None:
            if isEmailAddressList(row['CC']):
                msg.CC = row['CC'].replace(',', ';')
            else:
                error('Invalid \'CC\' value \'%s\' in row %d' % (row['CC'], n))

        if 'BCC' in row and row['BCC'] is not None:
            if isEmailAddressList(row['BCC']):
                msg.BCC = row['BCC'].replace(',', ';')
            else:
                error('Invalid \'BCC\' value \'%s\' in row %d' % (row['BCC'], n))

        if args.html:
            msg.HTMLBody = template(args.TEMPLATE, row)
        else:
            msg.Body = template(args.TEMPLATE, row)

        if 'ATTACHMENTS' in row and row['ATTACHMENTS'] is not None:
            for attachment in row['ATTACHMENTS'].replace(',', ';').split(';'):
                attachment = attachment.strip()
                if attachment:
                    msg.Attachments.Add(path.join(args.FOLDER, attachment))

        messages.append(msg)
    except Exception as e:
        error('Unable to create message from row %d (%s)' % (n, e))

info('Created %d message(s)' % len(messages))

if len(messages):
    msg = messages[0]
    info("This is the first message that will be sent:")
    print(file=stderr)
    print('\tFrom:', msg.Sender or 'Default Account', file=stderr)
    print('\tTo:', msg.To, file=stderr)
    print('\tCC:', msg.CC, file=stderr)
    print('\tBCC:', msg.BCC, file=stderr)
    print('\tSubject:', msg.Subject, file=stderr)
    print('\tAttachments (%d):' % len(msg.Attachments), ', '.join(str(s) for s in msg.Attachments), file=stderr)
    print(file=stderr)
    print('\t'+'\n\t'.join('\n\t'.join(wrap(line)) for line in msg.Body.split('\n')))
    print(file=stderr)

# # Final Confirmation
if not confirm('Is this correct?') or not confirm('Send?'):
    error('User aborted! No messages sent!')

for msg in messages:
    address = msg.To
    msg.Send()
    info('Sent message to \'%s\'' % address)

info('Sent %d messages' % len(messages))