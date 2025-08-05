# import click
from os import path
from sys import stderr, exit, version_info as py_version
import re as regex
import signal

VERSION='1.0.0'

def error(msg: str):
    print('\x1b[31m[Error]\x1b[39m: %s' % msg, file=stderr)
    exit(1)

if not (py_version.major == 3 and py_version.minor >= 12):
    error(f'''Running on Python {py_version.major}.{py_version.minor}, please update to at least 3.12!''')

signal.signal(signal.SIGINT, lambda signum, frame: print(file=stderr) or error('User aborted! No messages sent!'))

def info(msg: str):
    print('\x1b[34m[Info]\x1b[39m: %s' % msg, file=stderr)

def warn(msg: str):
    print('\x1b[33m[Warning]\x1b[39m: %s' % msg, file=stderr)

def confirm(msg: str, default: bool | None = None) -> bool:
    try:
        match input('\x1b[32m[Please Confirm]\x1b[39m: %s (%s) ' % (msg, 'Y/n' if default else 'y/N')).lower():
            case 'y' | 'yes': return True
            case 'n' | 'no': return False
            case _ if type(default) is bool: return default
            case t: 
                warn(f'''Unexpected response '{t}', try again...''')
                return confirm(msg, default)
    except KeyboardInterrupt:
        print(file=stderr)
        return False

def hasimport(module: str) -> bool:
    import importlib.util
    return importlib.util.find_spec(module) is not None

def sanitize(key: str) -> str:
    return regex.compile(r'\s+|\W|^(?=\d)').sub('_', key).upper()

def execexpr(script: str, globals=None, locals=None):
    import ast
    stmts = list(ast.iter_child_nodes(ast.parse(script)))
    if not stmts:
        return None
    if isinstance(stmts[-1], ast.Expr):
        if len(stmts) > 1:
            exec(compile(ast.Module(body=stmts[:-1],type_ignores=[]), filename="<ast>", mode="exec"), globals, locals)
        return eval(compile(ast.Expression(body=stmts[-1].value), filename="<ast>", mode="eval"), globals, locals)
    else:
        return exec(script, globals, locals)

def template(text: str, scope: dict = dict()) -> str:
    return regex.compile(r'\{\{(.+?)(?::(.*?))?\}\}', regex.DOTALL).sub(lambda m: f'''{(lambda arg: '' if arg is None else arg)(execexpr(m.group(1), dict(), scope)):{m.group(2) or ''}}''', text)

@lambda _: _()
def args():
    from argparse import ArgumentParser, RawTextHelpFormatter
    from textwrap import dedent
    from os import getcwd

    version = f'''Trin's Super Mail Script (v{VERSION})'''

    parser = ArgumentParser(formatter_class=RawTextHelpFormatter, add_help=False, description=dedent(f'''
        {version}
        {'-'*len(version)}
        This is a python utility for sending mail via Outlook. Note that the recent web-app version won't work.
        You'll need the classic version from https://support.microsoft.com/en-us/office/install-or-reinstall-classic-outlook-on-a-windows-pc-5c94902b-31a5-4274-abb0-b07f4661edf5.
    
        To start, run the script with a nonexistent filepath ending in '.txt' or '.html'. If the 'markdown' package is installed, '.md' also works. Next, open and edit the newly created file.
        
        The header (everything up until a line of three or more dashes) is YAML and controls things like the subject, recipient, CC, and BCC parts of the email. The header may contain these fields:
          FROM                  Sets which account to send the mail from (Accepts email address, defaults to Outlook's default)
          TO,CC,BCC             Sets the various recipients of the message (Accepts a comma or semicolon separated list of addresses and contact names)
          ATTACHMENTS           Specifies files relative to the '--dir' parameter below to attach to the message (Accepts a comma or semicolon separated list of paths)
          SUBJECT               Sets the subject line of the message, can contain template formatting just like the body (Default empty)
          RUN                   Provides a python initialization script to run before templating, can span multiple lines if starting with '|' (Default empty)

        The body of the message is interpreted as whatever format the file is (either plain text, HTML, or Markdown). Note that markdown will be converted to HTML internally. Additionally, values inside of double curly brackets are evaluated like Python f-string expressions (e.g., {'\'{{2+2:.1f}}\' -> \'5.0\''}). Unlike Python, statements are allowed so long as they are followed by a semicolon and an expression at the very end. The initialization script and all template expressions share a scope.
    
        When a mail merge data source is provided, instead of sending a single message, the message file is used as a template for each row of the data. Each row may specify the header parameters TO, CC, BCC, ATTACHMENTS, and SUBJECT (not RUN or FROM). Providing SUBJECT here will overwrite any global one while the other parameters concat their values to existing ones. Other values in the row are exposed to the templating expressions. Note that case is ignored and non [A-Z0-9_] characters are replaced with underscores to make valid identifiers ('1-More Thing' -> '_1_MORE_THING').

        Finally, nothing is sent until you preview the results and confirm. In most cases, user errors are caught before anything is sent; however, Outlook errors or network issues may interrupt partway through sending the list of messages.
    '''))
    parser.add_argument('--help', '-h', action='help', help='Show this message and exit')
    parser.add_argument('--version', '-v', action='version', version=version, help='''Show program's version number and exit''')
    parser.add_argument('FILE', type=str, help='''The message to send (Accepts '.txt' and '.html' files plus '.md' if the 'markdown' package is installed)''')
    parser.add_argument('-M', '--merge', dest='MERGE_FILE', type=str, help='''If set, treats the message as a mail merge template and fills it in with values from this file (Accepts '.csv')''')
    parser.add_argument('-D', '--dir', dest='DIRECTORY', type=str, default=getcwd(), help='''Where to look for attachments (Accepts a folder path, default is current folder)''')
    args = parser.parse_args()
    args.DIRECTORY = path.abspath(args.DIRECTORY)
    args.FILE = path.abspath(args.FILE)
    if args.MERGE_FILE:
        args.MERGE_FILE = path.abspath(args.MERGE_FILE)
    return args

if not path.exists(args.FILE):
    from textwrap import dedent
    try:
        info(f'''Creating new empty message file '{args.FILE}\'''')
        with open(args.FILE, 'x') as file: file.write(dedent('''
            # Save and close to finish
            FROM: 
            TO: 
            CC: 
            BCC: 
            SUBJECT: 
            ---
        ''').lstrip())
        
        if hasimport('click'):
            from click import edit
            edit(filename=args.FILE)
        else:
            warn(f'''Manually edit '{args.FILE}' and try again!''', file=stderr)
            exit(1)

    except Exception as e:
        error(f'''Unable to create new message file! ({e})''')

@lambda _: _()
def message():
    @lambda _: _()
    def message():
        from dataclasses import dataclass, field
        from typing import Literal

        @dataclass
        class Header:
            FROM: str | None = None
            TO: list[str] = field(default_factory=list)
            CC: list[str] = field(default_factory=list)
            BCC: list[str] = field(default_factory=list)
            ATTACHMENTS: list[str] = field(default_factory=list)
            SUBJECT: str = ''
            RUN: str | None = None

        @dataclass
        class Message:
            HEADER: Header = field(default_factory=Header)
            BODY: str = ''
            TYPE: Literal['hmtl', 'text', 'markdown'] = 'text'
        return Message()

    try:
        with open(args.FILE) as f:
            info(f'''Reading message from '{args.FILE}\'''')
            header, body = (lambda a: [''] * max(0, 2 - len(a)) + a)(regex.compile(r'^-{3,}$', regex.MULTILINE).split(f.read(), 1))
            message.BODY = body.strip()
    except Exception as e:
        error(f'''Unable to open file '{file}'! ({e})''')

    try:
        import yaml
        for key, value in yaml.safe_load(header).items():
            match (key.upper(), value):
                case (k, None): pass
                case (k, str()) if hasattr(message.HEADER, k) and type(getattr(message.HEADER, k)) == list: getattr(message.HEADER, k).extend([tt for t in value.replace(',', ';').split(';') if (tt := t.strip())])
                case (k, list()) if hasattr(message.HEADER, k) and type(getattr(message.HEADER, k)) == list: getattr(message.HEADER, k).extend(value)
                case (k, str()) if hasattr(message.HEADER, k): setattr(message.HEADER, k, value.strip())
                case (k, _) if hasattr(message.HEADER, k): setattr(message.HEADER, k, value)
                case _: error(f'''Unknown header '{key}'!''')
    except Exception as e:
        error(f'''Unable to parse message header! ({e})''')
     
    match path.splitext(args.FILE)[1][1:].lower():
        case 'txt': 
            info('Message(s) will be treated as plain text')
            message.TYPE = 'text'
        case 'html':
            info('Message(s) will be treated as HTML')
            message.TYPE = 'html'
        case 'md' if hasimport('markdown'):
            info('Message(s) will be converted from Markdown to HTML')
            message.TYPE = 'markdown'
        case ext: error(f'''Unsupported file type '.{ext}'!''')

    return message

@lambda _: _()
def data():
    if args.MERGE_FILE:
        if not args.MERGE_FILE.endswith('.csv'):
            error(f'''Mail merge data must be a '.csv' file!''')

        info(f'''Message '{args.FILE}' will be used as the template with values from '{args.MERGE_FILE}' to create a mail merge''')

        try:
            with open(args.MERGE_FILE) as f:
                try:
                    from csv import DictReader
                    return [{k.upper(): v.strip() if v else v for k, v in row.items()} for row in DictReader(f, skipinitialspace=True)]
                except Exception as e:
                    error(f'''Unable to parse mail merge data! ({e})''')
        except Exception as e:
            error(f'''Unable to open file '{args.MERGE_FILE}'! ({e})''')
    return [{}]


@lambda _: _()
def outlook():
    try:
        info(f'''Trying to connect to Outlook... (If this takes more than a couple minutes, it's not working!)''')
        from win32com.client import Dispatch as win32_dispatch
        return win32_dispatch('Outlook.Application')
    except Exception as e:
        error(f'''Unable to connect to Outlook! ({e})''')
info('Successfully connected to Outlook')

@lambda _: _()
def messages():
    messages = []
    class Scope(dict):
        def __getitem__(self, key):
            return super().__getitem__(sanitize(key))

        def __setitem__(self, key, value):
            super().__setitem__(sanitize(key), value)

        def __contains__(self, key):
            return super().__contains__(sanitize(key))

    for n, row in enumerate(data):
        n=n+1
        try:
            msg = outlook.CreateItem(0)
            scope = Scope(**dict([[k, [*v]] for k, v in message.HEADER.__dict__.items() if type(v) == list]) , SUBJECT=message.HEADER.SUBJECT, N=n)

            for k,v in row.items():
                k = sanitize(k)
                if hasattr(message.HEADER, k) and type(getattr(message.HEADER, k)) == list:
                    scope[k] += v if type(v) == list else [vv for v in v.replace(',', ';').split(';') if (vv := v.strip())]
                elif not k == 'SUBJECT' or v:
                    scope[k] = v

            if message.HEADER.RUN:
                try:
                    execexpr(message.HEADER.RUN, dict(), scope)
                except Exception as e:
                    error(f'''Unable to run initialization script for message {n}! ({e})''')

            msg.Sender = message.HEADER.FROM

            try:
                msg.Subject = template(scope['SUBJECT'], scope)
                if not msg.Subject:
                    warn(f'''Subject line of message {n} is empty!''')
            except Exception as e:
                error(f'''Unable to template subject line for message {n}! ({e})''')
            
            try:
                match message.TYPE:
                    case 'text': msg.Body = template(message.BODY, scope)
                    case 'html': msg.HTMLBody = template(message.BODY, scope)
                    case 'markdown':
                        if hasimport('markdown'):
                            from markdown import markdown
                            msg.HTMLBody = markdown(template(message.BODY, scope))
                        else:
                            error('Got markdown content but no markdown support available')
            except Exception as e:
                error(f'''Unable to template body for message {n}! ({e})''')

            if len(scope['TO']) + len(scope['CC']) + len(scope['BCC']) == 0:
                error(f'''No recipients in TO, CC, or BCC fields of message {n}!''')

            if any(not regex.compile('^\\S+@\\S+\\.\\S+$').match(e) for e in [*scope['TO'], *scope['CC'], *scope['BCC']]):
                warn(f'''Message {n} uses contact names instead of full emails!''')

            msg.To = ';'.join(scope['TO'])
            msg.CC = ';'.join(scope['CC'])
            msg.BCC = ';'.join(scope['BCC'])
            
            for attachment in scope['ATTACHMENTS']:
                msg.Attachments.Add(path.join(args.DIRECTORY, attachment))

            messages.append(msg)
        except Exception  as e:
            error(f'''Unable to create message {n}! ({e})''')
    return messages

info(f'''Created {len(messages)} message(s)''')

if path.isdir(args.DIRECTORY):
    info(f'''Attachments will be loaded relative to '{args.DIRECTORY}\'''')
else:
    error('''Directory argument must point to a valid folder!''')

if not message.HEADER.FROM:
    warn('''Messages will be sent as default user!''')

if not message.BODY:
    warn('''Message body is empty!''')

if len(messages):
    from textwrap import wrap
    msg = messages[0]
    info('This is the first message that will be sent:')
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

# Final Confirmation
if not confirm('Is this correct?') or not confirm(f'''Send {len(messages)} message(s)?'''):
    error('User aborted! No messages sent!')

for msg in messages:
    address = msg.To
    msg.Send()
    info(f'''Sent message to '{address}\'''')

info(f'''Sent {len(messages)} message(s) successfully''')