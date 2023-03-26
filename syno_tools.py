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


def get_link(file_data,name=''):
    if file_data['type'] == 'file':
        chain = 'oo/r'
    else:
        chain = 'd/f'
    if name == '': name = file_data["name"].split(".")[0]

    return f'=HYPERLINK("#dlink=/{chain}/{file_data["permanent_link"]}", "{name}")'
    


def convert_to(synd,file,delete=False):
    try:
        print("Converting...",file)
        r_con = synd.convert_to_online_office(file,
            delete_original_file=delete,
            conflict_action='autorename')
        
        print(file[-4:],file[:-5],r_con) 
        #if file[-4:] == 'xlsx':
        #    print(f"{file[:-5]}.osheet")
        #    new_file = synd.get_file_or_folder_info(f"{file[-5:]}.osheet")

        #return True,get_link(new_file['data'])
    except:
        raise
        return False,[]

def copy_to(synd,file,dest,convert = False):
    try:
        synd.copy(file['display_path'],f"{dest}/{file['name']}")
    except:
        try:
            tmp_file = synd.download_file(file['display_path'])
            r_up = synd.upload_file(tmp_file,dest_folder_path=dest)
            try:
                if convert:
                    r_con = synd.convert_to_online_office(r_up['data']['display_path'],
                    delete_original_file=False,
                    conflict_action='autorename')
            except:
                print("Cannot convert")
        except:
            print("Cannot copy the file")
            return False
    
    return True
                
def move_to(synd,file,dest,team_data,convert=False):
    links = []
    try:
        synd.move_path(file['display_path'],dest)
        if file['name'].split('.')[-1] in ['odoc','osheet']:
            links.append(get_link(file))
        else:
            links.append(get_link(team_data,file['name'].split('.')[0]))
        
        if file['name'].split('.')[-1] in ['xlsx','docx'] and convert:
            result,lns = convert_to(synd,f"{dest}/{file['name']}")
            if result: links += lns

        return True,links
    except:
        raise
        return False,[]
