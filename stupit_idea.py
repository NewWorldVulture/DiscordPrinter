# Brought to life by Ada
#
# 2020/10/31

import win32print
from subprocess import Popen, PIPE
import string
import os
import discord

CLIENT_TOKEN = ""

# Templates for use later
PRINT_PREFIX = string.Template("^XA^GB1980,0,0^LL$LabelLength ^FO10,10^ADN,24,18^FD $author ^FS")
PRINT_CONTENT = string.Template("^FO 30, $yCoord ^ADN,30,24^FD $text ^FS")
PRINT_POSTFIX = "^XZ"

# takes content of message and formats it into ZPL
def createMessage(author, text):
    text_len = len(text)

    # Split message every 44 characters.
    text_parts = [text[i:i+44] for i in range(0, text_len, 44)]
    label_len = len(text_parts) * 32 + 40

    message = PRINT_PREFIX.substitute(LabelLength=label_len, author=author)

    for c, line in enumerate(text_parts):
        yCoord = 40 + c*32
        message += PRINT_CONTENT.substitute(yCoord=yCoord, text=line)

    message += PRINT_POSTFIX
    return message

client = discord.Client()

def get_online_printer():
    # Uses Powershell to grab all online printers from 
    process = Popen(['Powershell', '. "./online_printers";', '&getOnlinePrinters($_)'],
        stdout=PIPE, stderr=PIPE)
    stdout, stderr = process.communicate()
    stdout = stdout.strip()
    return stdout #stdout returns as bytecode. Decode it to utf8

@client.event
async def on_ready():
    print(f'{client.user} has connected to Discord.')
    print('Connecting to printer...\n')
    PRINTER_NAME = get_online_printer().decode("utf-8")
    print(f'Found printer: {PRINTER_NAME}')

@client.event
async def on_message(message):
    # Document message being printed
    #print(message.author + " : " + message.content)

    # Create print message
    message_thing = PRINT_TEMPLATE.substitute(author=str(message.author)[:-5], content=createMessage(message.author, message.content))
    message_thing = bytes(message_thing, "utf-8")

    try:
        print('Connecting to printer...\n')
        PRINTER_NAME = get_online_printer().decode("utf-8")
        print(f'Found printer: {PRINTER_NAME}')
        print()

        p = win32print.OpenPrinter(PRINTER_NAME)
        print()
        try:
            #open printer, print, close out printer 
            win32print.StartDocPrinter(p, 1,("Printer Test", None, "RAW"))
            win32print.StartPagePrinter(p)
            win32print.WritePrinter(p, message_thing)
            win32print.EndPagePrinter(p)
        except Exception as e:
            print("Print Error: ", e)
        win32print.ClosePrinter(p)
    except:
        print("Error finding printer")

client.run(CLIENT_TOKEN)
