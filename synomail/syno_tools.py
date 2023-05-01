#!/bin/python

from getpass import getpass
from datetime import datetime
from tempfile import NamedTemporaryFile
from pathlib import Path
import time

import logging

from openpyxl import Workbook
from synology_drive_api.drive import SynologyDrive

from synomail import EXT

def build_link(p_link,name_link,is_folder = False):
    if is_folder:
        chain = 'd/f'
    else:
        chain = 'oo/r'

    return f'=HYPERLINK("#dlink=/{chain}/{p_link}", "{name_link}")' if p_link != '' else name_link


def move_to(synd,file_path,dest):
    logging.debug(f"Moving {file_path}...")        
    try:
        rst = synd.move_path(file_path,dest)
        task_id = rst['data']['async_task_id']
    
        rst = synd.get_task_status(task_id)
        
        while(rst['data']['result'][0]['data']['progress'] < 100 or rst['data']['has_fail']):
            time.sleep(0.2)
            rst = synd.get_task_status(task_id)

        rst_data = rst['data']['result'][0]['data']['result']
        
        if not 'targets' in rst_data:
            logging.error('Synology cannot move the file {file_path}')

    except Exception as err:
        logging.error(err)
        logging.warning('Cannot move the file {file_path}')


def convert_to(synd,file_path,delete = False):
    logging.info(f"Converting {file_path}...")        
    try:
        rst = synd.convert_to_online_office(file_path,delete_original_file=delete)
        task_id = rst['data']['async_task_id']
    
        rst = synd.get_task_status(task_id)
        while(not rst['data']['has_fail'] and rst['data']['result'][0]['data']['status'] == 'in_progress'):
            rst = synd.get_task_status(task_id)
        
    except Exception as err:
        logging.error(err)
        logging.warning(f'Cannot convert {file_path}')
        return '','',''
   
        
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
