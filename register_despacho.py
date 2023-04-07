#!/bin/python

from getpass import getpass
from datetime import datetime
from tempfile import NamedTemporaryFile
import time
import re
from pathlib import Path
import ast


from openpyxl import Workbook, load_workbook
from synology_drive_api.drive import SynologyDrive

from syno_tools import txt2dict, copy_to, move_to, convert_to, EXT

TO_ARCHIVE = "Outbox Despacho/to_archive"

TITLES = ['type','source','No','Year','Ref','Date','Content','Dept','Name','Original'] 

def read_register(PASS):
    config = txt2dict("config.txt")
    
    with SynologyDrive(config['user'],PASS,"nas.prome.sg",dsm_version='7') as synd:
        register = synd.list_folder(f"{config['despacho']}/Outbox Despacho")['data']['items']
        pattern = re.compile("\d\d-\d\d-\d\d\d\d-ctr-incoming.osheet")
        notes = {}
        for reg in register:
            if re.match(pattern,reg["name"]):
                try:
                    reg_file = synd.download_file(reg['display_path'])
                    wb = load_workbook(reg_file)
                    for ws_name in wb.sheetnames:
                        ws = wb[ws_name]
                        
                        for row in ws.iter_rows(values_only=True):
                            if row[2] != '' and row[1] != '' and row[0] != 'type':
                                notes[f"{row[1]}{row[2]}"] = {'notes':[]}
                       
                        for row in ws.iter_rows(values_only=True):
                            if row[2] != '' and row[1] != '' and row[0] != 'type':
                                temp = {}
                                for i,title in enumerate(TITLES):
                                    temp[title] = row[i]

                                notes[f"{row[1]}{row[2]}"]['notes'].append(temp.copy())
                except:
                    raise
                    print(f"Cannot download {note['name']}")
                    continue

        return notes

def archive_notes(PASS,reg_notes):
    config = txt2dict("config.txt")
    path = f"{config['despacho']}/{TO_ARCHIVE}"
    
    with SynologyDrive(config['user'],PASS,"nas.prome.sg",dsm_version='7') as synd:
        for num,data in reg_notes.items():
            dest = f"{config['archive']}/{data['type']} in {data['Year']}"
            
            if data['create_folder']:
                try:
                    new_folder = synd.get_file_or_folder_info(f"{dest}/{num}")
                    exist = True
                except:
                    exist = False

                try:
                    if not exist:
                        rst = synd.create_folder(num, dest)
                        p_link = rst['data']['permanent_link']
                    else:
                        p_link = new_folder['data']['permanent_link']

                    dest = f"{dest}/{num}"
                except:
                    print(f"Cannot create folder {num}")
                    continue

            if data['Dept'] != '':
                for note in data['notes']:
                    name = note['Name'][:-2].split('","')[1]
                    #print(f"{path}/{name}",dest)
                    #f_id,p_link = move_to(synd,f"{path}/{name}",dest)
                    if note['Original'] != '':
                        move_to(synd,f"{path}/{note['Original']}",dest)

                src = re.findall('\d+',num)
                only_num = src[0] if src else num
                if data['create_folder']:
                    data['link'] = f'=HYPERLINK("#dlink=/d/f/{p_link}", "{only_num}")'
                else:
                    p_link = 'patata'
                    data['link'] = f'=HYPERLINK("#dlink=/oo/r/{p_link}", "{only_num}")'

def dept_notes(PASS,reg_notes):
    config = txt2dict("config.txt")
    path = f"{config['despacho']}/Inbox Despacho"
    
    with SynologyDrive(config['user'],PASS,"nas.prome.sg",dsm_version='7') as synd:
        for num,data in reg_notes.items():
            if data['Dept'] != '':
                dest = []
                for dep in data['Dept'].split(","):
                    dest.append(f"/team-folders/Mail {dep}/Inbox {dep}")

                for note in data['notes']:
                    name = note['Name'][:-2].split('","')[1]
                    print(dest)
                    for des in dest:
                        if des == dest[-1]:
                            #print('COPY-Move',f"{path}/{name}",des)
                            move_to(synd,f"{path}/{name}",des)
                        else:
                            #print('COPY',f"{path}/{name}",des)
                            copy_to(synd,f"{path}/{name}",des)

def upload_register(PASS,wb):
    config = txt2dict("config.txt")
    date = datetime.today().strftime('%d-%m-%Y')
    
    with SynologyDrive(config['user'],PASS,"nas.prome.sg",dsm_version='7') as synd:
        file = NamedTemporaryFile()
        wb.save(file)
        file.seek(0)
        file.name = f"{date}-register.xlsx"
    
        print("Creating register file")
        uploaded = True
        
        try:
            ret_upload = synd.upload_file(file, dest_folder_path=f"{config['despacho']}/Outbox Despacho")
        except:
            print("Cannot upload register")
            wb.save(f"{date}-register.xlsx")
            uploaded = False

        if uploaded:
            try:
                ret_convert = synd.convert_to_online_office(ret_upload['data']['display_path'],
                    delete_original_file=True,
                    conflict_action='autorename')
            except:
                print("Cannot convert file to Synology Office")

        return True
    
    print("Cannot upload register")
    wb.save(f"{date}-register.xlsx")


def create_register(ws,reg_notes):
    reg_titles = ['source','link','Year','Ref','Date','Content','Dept']
    ws.append(reg_titles)

    for num,data in reg_notes.items():
        row = []
        for title in reg_titles:
            #if data['No'] != '' and row['Dept'] != '':
            row.append(data[title])

        ws.append(row)


def fill_data(reg_notes):
    not_to_reg = ['Name','Original']
    for num,data in reg_notes.items():
        create_folder = False
        cont = 0
        for note in data['notes']:
            cont += 1
            for title in TITLES:
                if not title in not_to_reg and note[title] != '':
                    data[title] = note[title]

        if cont > 1:
            create_folder = True
        else:
            nm = note['Name'][:-2].split('","')[1]
            ext = Path(nm).suffix[1:]
            if not ext in EXT.values():
                create_folder = True

        data['create_folder'] = create_folder
                
            


def main():
    PASS = getpass()
    
    reg_notes = read_register(PASS)
    input("go")

    if reg_notes != {}:
        fill_data(reg_notes)
        #print(reg_notes)

        archive_notes(PASS,reg_notes)
        dept_notes(PASS,reg_notes)
        
        wb = Workbook()
        ws = wb.active
        create_register(ws,reg_notes)
        upload_register(PASS,wb)

    
    input("Pulse Enter to continue")



if __name__ == '__main__':
    main()
