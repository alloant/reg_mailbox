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

from syno_tools import txt2dict, copy_to, convert_to


def get_ctr(folder):
    if folder[:7] == 'Mailbox':
        return folder[8:]
    elif folder[:10] == '8.Mailbox-':
        return folder[10:]


ext = {'xls':'osheet','xlsx':'osheet','docx':'odoc'}

def tests(PASS):
    config = txt2dict("config.txt")
    year = datetime.today().strftime('%Y')

    with SynologyDrive(config['user'],PASS,"nas.prome.sg",dsm_version='7') as synd:
        parent_id = "738616337546392650"
        file_id = "744010245052872375"
        file_id = "744271352449511822"
        #rst = synd.move_path("/mydrive/Admin/Despacho/to_archive/gul4.xlsx","/mydrive/Admin/Despacho")
        #print(rst)
        #rst = synd.get_task_status(rst['data']['async_task_id'])
        #file_id = rst['data']['result'][0]['data']['result']['targets'][0]['file_id']
        #file_id = "744377953154021634"
        #print(file_id)
        #synd.copy(int(file_id),"/mydrive/Admin/Despacho/to_archive/gul4.xlsx")

        #notes = synd.list_folder('mydrive/Admin/tests')['data']['items']
        #for note in notes:
        #    print(get_link(synd,note))
        #print(synd.get_file_or_folder_info(parent_id))
        #print(synd.get_file_or_folder_info(file_id))
        #print(synd.create_link(parent_id))
        #print(synd.create_link(file_id))
        #file_path = "/mydrive/Admin/Despacho/to_archive/223bPart2records.osheet"
        #print(synd.get_file_or_folder_info(file_path))
        
        #pid = "744419054692908035"
        #print(synd.get_file_or_folder_info(pid))
        #rst = synd.copy(pid,"/mydrive/Admin/Despacho/223.osheet")
        #file_id = "744419054692908035"
        

        #print(synd.get_file_or_folder_info("/mydrive"))

        #path = "mydrive/Admin/Despacho/to_archive/gul3.osheet"
        #dest = "mydrive/Admin/Despacho"

        #folder_id = "742786542633793039"
        #print(synd.get_file_or_folder_info("/mydrive/Admin/Despacho"))
        #print(synd.copy3(path,dest))
        #rst = synd.copy2(f"id:{file_id}","/mydrive/Admin/Despacho/2.23r.osheet",f"id:{folder_id}")
        #print(synd.get_file_or_folder_info(file))
        #copy_to(synd,f"id:{file_id}",f"/mydrive/Admin/Despacho")

        #synd.rename_path('gul003.osheet', '/mydrive/Admin/Despacho/to_archive/003.osheet')
        #synd.copy("/mydrive/Admin/Despacho/to_archive/003.osheet","/mydrive/gul3.osheet")
        #synd.copy("/mydrive/Admin/Despacho/to_archive/003.osheet","/mydrive/Admin/Despacho/003.osheet")

        convert_to(synd,"/mydrive/Admin/Despacho/to_archive/gul_test.rtf")


def main():
    PASS = getpass()
   
    tests(PASS)


if __name__ == '__main__':
    main()
