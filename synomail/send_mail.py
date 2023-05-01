#!/bin/python

from getpass import getpass
from datetime import datetime
from tempfile import NamedTemporaryFile
import time
import re
from pathlib import Path

import logging

import pickle

from openpyxl import Workbook
from synology_drive_api.drive import SynologyDrive

from synomail import CONFIG, GROUPS, EXT


def copy_to(synd,file_path,dest):
    name = Path(file_path).name
    try:
        ext = Path(file_path).suffix[1:]

        if ext in EXT.values():
            rst = synd.copy(file_path,f"{dest}/{name}")
            logging.debug(f"Copied file {name} to {dest}")
        else:
            tmp_file = synd.download_file(file_path)
            rst = synd.upload_file(tmp_file,dest_folder_path=dest)
    except:
        logging.error(f"Cannot copy the file {name} to {dest}")

def move_to(synd,file_path,dest):
    logging.debug(f"Moving {file_path}...")        
    try:
        rst = synd.move_path(file_path,dest)
    except:
        logging.error(f"Cannot move {file_path}")


def change_names(PASS):
    with SynologyDrive(CONFIG['user'],PASS,"nas.prome.sg",dsm_version='7') as synd:
        for group,ctrs in GROUPS.items():
            try:
                notes = synd.list_folder(f"/mydrive/ToSend/{group}")['data']['items']
            except Exception as err:
                logging.warning(err)
                continue

            for note in notes:
                try:
                    if note['name'][0].isdigit():
                        synd.rename_path(f"cr{note['name']}",f"{note['display_path']}")
                except Exception as err:
                    logging.error(err)
                    logging.error(f"Cannot rename note['name']")



def send_to_all(PASS):
    with SynologyDrive(CONFIG['user'],PASS,"nas.prome.sg",dsm_version='7') as synd:
        for group,ctrs in GROUPS.items():
            try:
                notes = synd.list_folder(f"/mydrive/ToSend/{group}")['data']['items']
            except Exception as err:
                logging.error(err)
                continue

            for note in notes:
                for ctr in ctrs.split(","):
                    if ctr == ctrs.split(",")[-1]:
                        move_to(synd,note['display_path'],f"/team-folders/Mailbox {ctr}/cr to {ctr}")
                    else:
                        copy_to(synd,note['display_path'],f"/team-folders/Mailbox {ctr}/cr to {ctr}")
                    

def init_send_mail(PASS):
    logging.info('Starting to send mail to ctr')
    change_names(PASS)
    send_to_all(PASS)
    logging.info('Finish to send mail to ctr')

def main():
    PASS = getpass()
    init_send_mail(PASS)     
    input("Done")

if __name__ == '__main__':
    main()
