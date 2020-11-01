# Brought to life by Ada
#
# 2020/10/31

import win32print
from subprocess import Popen, PIPE
import string
import os
import discord

PRINT_TEMPLATE = string.Template("^XA^^LL120^FO40,20^ADN,20,16^FD $author ^FS^FD20,20^ADN,20,16^FD $content^FS^XZ")
CLIENT_TOKEN = ""

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
    message_thing = PRINT_TEMPLATE.substitute(author=message.author, content=message.content)
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