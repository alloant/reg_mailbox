#!/bin/python

from getpass import getpass
from datetime import datetime
from tempfile import NamedTemporaryFile
import time
import re
from pathlib import Path

import pickle

from openpyxl import Workbook
from synology_drive_api.drive import SynologyDrive

EXT = {'xls':'osheet','xlsx':'osheet','docx':'odoc'}


def txt2dict(file):
    DICT = {}
    
    with open(file,mode='r') as inp:
        lines = inp.read().splitlines()

        DICT = {ln.split(':',1)[0]:ln.split(':',1)[1] for ln in lines if ":" in ln}
    
    return DICT


def copy_to(synd,file_path,dest):
    name = Path(file_path).name
    try:
        ext = Path(file_path).suffix[1:]

        if ext in EXT.values():
            rst = synd.copy(file_path,f"{dest}/{name}")
            print(f"Copied file {name} to {dest}")
        else:
            tmp_file = synd.download_file(file_path)
            rst = synd.upload_file(tmp_file,dest_folder_path=dest)
    except:
        print("ERROR: Cannot copy the file {name} to {dest}")

def move_to(synd,file_path,dest):
    print(f"Moving {file_path}...")        
    try:
        rst = synd.move_path(file_path,dest)
    except:
        print("Cannot move")


def change_names(PASS):
    config = txt2dict("config.txt")
    groups = txt2dict("groups.txt")

    with SynologyDrive(config['user'],PASS,"nas.prome.sg",dsm_version='7') as synd:
        for group,ctrs in groups.items():
            try:
                notes = synd.list_folder(f"/mydrive/ToSend/{group}")['data']['items']
            except:
                continue

            for note in notes:
                try:
                    if note['name'][0].isdigit():
                        synd.rename_path(f"cr{note['name']}",f"{note['display_path']}")
                except:
                    raise
                    print("Cannot rename")



def send_to_all(PASS):
    config = txt2dict("config.txt")
    groups = txt2dict("groups.txt")

    with SynologyDrive(config['user'],PASS,"nas.prome.sg",dsm_version='7') as synd:
        for group,ctrs in groups.items():
            try:
                notes = synd.list_folder(f"/mydrive/ToSend/{group}")['data']['items']
            except:
                continue

            for note in notes:
                for ctr in ctrs.split(","):
                    if ctr == ctrs.split(",")[-1]:
                        move_to(synd,note['display_path'],f"/team-folders/Mailbox {ctr}/cr to {ctr}")
                    else:
                        copy_to(synd,note['display_path'],f"/team-folders/Mailbox {ctr}/cr to {ctr}")
                    



def main():
    PASS = getpass()

    change_names(PASS)
   
    send_to_all(PASS)

    input("Done")

if __name__ == '__main__':
    main()
