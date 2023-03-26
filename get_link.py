#!/bin/python

from getpass import getpass
from datetime import datetime
from tempfile import NamedTemporaryFile

from openpyxl import Workbook
from synology_drive_api.drive import SynologyDrive

def txt2dict(file):
    DICT = {}
    
    with open(file,mode='r') as inp:
        lines = inp.read().splitlines()

        DICT = {ln.split(':')[0]:ln.split(':')[1] for ln in lines if ":" in ln}
    
    return DICT

config = txt2dict("config.txt")

PASS = getpass()

with SynologyDrive(config['user'],PASS,"nas.prome.sg",dsm_version='7') as synd:
    link = synd.create_link('/myDrive/Admin/Aes Archive/test.xlsx')
    print(link)
