#!/bin/python

from getpass import getpass
from datetime import datetime
from tempfile import NamedTemporaryFile
from pathlib import Path
import time

from openpyxl import Workbook
from synology_drive_api.drive import SynologyDrive


EXT = {'xls':'osheet','xlsx':'osheet','docx':'odoc'}

def txt2dict(file):
    DICT = {}
    
    with open(file,mode='r') as inp:
        lines = inp.read().splitlines()

        DICT = {ln.split(':',1)[0]:ln.split(':',1)[1] for ln in lines if ":" in ln}
    
    return DICT

def build_link(p_link,name_link,is_folder = False):
    if is_folder:
        chain = 'd/f'
    else:
        chain = 'oo/r'

    return f'=HYPERLINK("#dlink=/{chain}/{p_link}", "{name_link}")' if p_link != '' else name_link


def move_to(synd,file_path,dest):
    print(f"Moving {file_path}...")        
    try:
        rst = synd.move_path(file_path,dest)
        task_id = rst['data']['async_task_id']
    
        rst = synd.get_task_status(task_id)
        
        while(rst['data']['result'][0]['data']['progress'] < 100 or rst['data']['has_fail']):
            rst = synd.get_task_status(task_id)
    
        print('Done')

        file_id = rst['data']['result'][0]['data']['result']['targets'][0]['file_id']
        permanent_link = rst['data']['result'][0]['data']['result']['targets'][0]['permanent_link']

        return file_id,permanent_link
    except:
        print('Cannot move')
        return '',''

def convert_to(synd,file_path,delete = False):
    print(f"Converting {file_path}...")        
    try:
        rst = synd.convert_to_online_office(file_path,delete_original_file=delete)
        task_id = rst['data']['async_task_id']
    
        rst = synd.get_task_status(task_id)
    except:
        raise
        return '',''
   
    while(not rst['data']['has_fail'] and rst['data']['result'][0]['data']['status'] == 'in_progress'):
        #print(".",end='')
        rst = synd.get_task_status(task_id)
    
    #time.sleep(10) 
    print('Done')
    
    ext = Path(file_path).suffix[1:]
    name = file_path.replace(ext,EXT[ext])
    
    new_file = synd.get_file_or_folder_info(name)
    file_id = new_file['data']['file_id']
    permanent_link = new_file['data']['permanent_link']
    file_path = new_file['data']['display_path'] 

    return file_path,file_id,permanent_link



def copy_to(synd,file_path,dest):
    error = False
    print(f"Copying file {file_path}...")
    try:
        name = Path(file_path).name
        ext = Path(file_path).suffix[1:]
        
        if ext in EXT.values():
            rst = synd.copy(file_path,f"{dest}/{name}")
        else:
            tmp_file = synd.download_file(file_path)
            rst = synd.upload_file(tmp_file,dest_folder_path=dest)
    except:
        raise
        print("ERROR: Cannot copy the file")
        error = True

    if error:
        return '',''
    else:
        print("Done")
        if 'permanent_link' in rst['data']:
            permanent_link = rst['data']['permanent_link']
        elif 'link_id' in rst['data']:
            permanent_link = rst['data']['link_id']
        else:
            permanent_link = ''


        file_id = rst['data']['file_id']

        return file_id,permanent_link
