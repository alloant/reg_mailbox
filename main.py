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
        return t[10:]


with SynologyDrive("df-scr",PASS,"nas.prome.sg",dsm_version='7') as synd:
    tf = synd.get_teamfolder_info()
    
    year = datetime.today().strftime('%Y')
    date = datetime.today().strftime('%d/%m/%Y')

    despacho = "/team-folders/Despacho/Inbox Despacho"
    archive = f"/team-folders/Aes Archive/ctr in {year}"

    #despacho = "/mydrive/despacho"
    #archive = "/mydrive/archive"

    wb = Workbook()
    ws = wb.active

    ws.append(['ctr','No','Year','Ref','Date','Content','Dept','#of docs'])
   
    cont = 0
    
    for t in tf:
        if 'Mailbox gul' in t:
            ctr = get_ctr(t)
            
            try:
                mbs = synd.list_folder(f"/team-folders/{t}")
            except:
                print(f'Cannot get folders from {t}')
                break

            for mb in mbs['data']['items']:
                if mb['name'] in [f"{ctr} to cr",f"{ctr}-to-cr"]:
                    print(f"Checking {t}")
                    p_path = synd.get_file_or_folder_info(f"{mb['display_path']}")['data']['permanent_link']
                    h_path = f'=HYPERLINK("#dlink=/oo/r/{p_path}", "{mb["name"]}")'
 
                    try:
                        mail = synd.list_folder(mb['display_path'])
                    except:
                        print(f"Cannot get file list from {mb['name']}")
                    
                    for m in mail['data']['items']:
                        cont += 1
                        note = f"{mb['display_path']}/{m['name']}"
                        
                        print(f"    Copying {m['name']} to despacho")
                        

                        try:
                            synd.copy(note,f"{despacho}/{m['name']}")
                        except:
                            print("Cannot copy files")
                            ws.append([ctr,h_path,year,'',date,f"ERROR in {mb['display_path']}",'',''])
                            break
                        
                        print("    Moving {m['name']} to archive")
                        
                        
                        try:
                            synd.move_path(note,archive)
                        except:
                            print(f"Cannot move {m['name']}")
                            ws.append([ctr,h_path,year,'',date,f"ERROR in {mb['display_path']}",'',''])
                            break

                        print("    Saving link in register")


                        try:
                            p_link = synd.get_file_or_folder_info(f'{archive}/{m["name"]}')['data']['permanent_link']
                            h_link = f'=HYPERLINK("#dlink=/oo/r/{p_link}", "{m["name"]}")'
                        except:
                            print("Cannot get link")
                            h_link = h_path
                        

                        ws.append([ctr,h_link,year,'',date,'','',''])

    
    if cont > 0:
        file = NamedTemporaryFile()
        wb.save(file)
        file.seek(0)
        file.name = f"{date.replace('/','-')}-reg.xlsx"
    
        print("Creating register file")
        uploaded = True
        try:
            ret_upload = synd.upload_file(file, dest_folder_path='/mydrive')
        except:
            print("Cannot upload register")
            wb.save(f"{date.replace('/','-')}-reg.xlsx")
            uploaded = False

        if uploaded:
            try:
                ret_convert = synd.convert_to_online_office(ret_upload['data']['display_path'],
                    delete_original_file=True,
                    conflict_action='autorename')
            except:
                print("Cannot convert file to Synology Office")
