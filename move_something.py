#!/bin/python

from getpass import getpass
from datetime import datetime
from tempfile import NamedTemporaryFile

from openpyxl import Workbook
from synology_drive_api.drive import SynologyDrive

from syno_tools import txt2dict, get_link, copy_to, move_to

def get_ctr(folder):
    if folder[:7] == 'Mailbox':
        return folder[8:]
    elif folder[:10] == '8.Mailbox-':
        return folder[10:]


config = txt2dict("config.txt")
year = datetime.today().strftime('%Y')
date = datetime.today().strftime('%d/%m/%Y')
PASS = getpass()

with SynologyDrive(config['user'],PASS,"nas.prome.sg",dsm_version='7') as synd:
    syn_notes = synd.list_folder(f"/mydrive/Temp")['data']
    notes = syn_notes['items']

    for note in notes:
        print(move_to(synd,note,'/mydrive','link_team',convert=True))
