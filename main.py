#!/bin/python

from getpass import getpass
from datetime import datetime
from tempfile import NamedTemporaryFile

#USER = getpass.getuser()
PASS = getpass()

from openpyxl import Workbook
from synology_drive_api.drive import SynologyDrive

# default http port is 5000, https is 5001. 

def get_ctr(folder):
    if folder[:7] == 'Mailbox':
        return t[8:]
    elif folder[:10] == '8.Mailbox-':
        return t[11:]

with SynologyDrive("vInd1",PASS,"nas.prome.sg",dsm_version='7') as synd:
    tf = synd.get_teamfolder_info()
    
    year = datetime.today().strftime('%Y')
    date = datetime.today().strftime('%d/%m/%Y')

    despacho = "/team-folders/Despacho/Inbox Despacho"
    archive = f"/team-folders/Aes Archive/ctr in {year}"

    wb = Workbook()
    ws = wb.active

    ws.append(['No','Year','Ref','Date','Content','Dept','#of docs'])
    
    
    for t in tf:
        #if t[:7] == 'Mailbox':
        if 'Mailbox' in t:
            ctr = get_ctr(t)

            mbs = synd.list_folder(f"/team-folders/{t}")
            for mb in mbs['data']['items']:
                if mb['name'] in [f"{ctr} to cr",f"{ctr}-to-cr"]:
                    print(f"Checking {t}")
                    
                    mail = synd.list_folder(mb['display_path'])
                    for m in mail['data']['items']:
                        note = f"{mb['display_path']}/{m['name']}"
                        
                        print("Copying note to despacho")
                        synd.copy(note,f"{despacho}/{m['name']}")
                        
                        print("Moving note to archive")
                        synd.move_path(note,archive)
                        
                        print("Saving link in register")
                        p_link = synd.get_file_or_folder_info(f'{archive}/{m["name"]}')['data']['permanent_link']
                        #link = f"https://nas.prome.sg:5001/oo/r/{p_link}"
                        h_link = f'=HYPERLINK("#dlink=/oo/r/{p_link}", "{m["name"]}")'


                        ws.append(['',year,h_link,date,'','',''])

    
    file = NamedTemporaryFile()
    wb.save(file)
    file.seek(0)
    file.name = f"{date.replace('/','-')}-reg.xlsx"
    
    print("Creating register file")
    ret_upload = synd.upload_file(file, dest_folder_path='/mydrive')
    ret_convert = synd.convert_to_online_office(ret_upload['data']['display_path'],
                    delete_original_file=True,
                    conflict_action='autorename')

