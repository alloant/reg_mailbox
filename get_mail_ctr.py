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

def get_ctr(folder):
    if folder[:7] == 'Mailbox':
        return t[8:]
    elif folder[:10] == '8.Mailbox-':
        return t[10:]


config = txt2dict("config.txt")


PASS = getpass()

with SynologyDrive(config['user'],PASS,"nas.prome.sg",dsm_version='7') as synd:
    tf = synd.get_teamfolder_info()
    
    year = datetime.today().strftime('%Y')
    date = datetime.today().strftime('%d/%m/%Y')

    wb = Workbook()
    ws = wb.active

    ws.append(['ctr','No','Year','Ref','Date','Content','Dept','#of docs'])
   
    cont = 0
    
    for t in tf:
        if 'Mailbox' in t:
            ctr = get_ctr(t)
            
            try:
                mbs = synd.list_folder(f"/team-folders/{t}")
            except:
                print(f'Cannot get folders from {t}')
                break
            
            # Cheking all folders
            for mb in mbs['data']['items']:
                if mb['name'] in [f"{ctr} to cr",f"{ctr}-to-cr"]:
                    print(f"Checking {t}")
                    p_path = synd.get_file_or_folder_info(f"{mb['display_path']}")['data']['permanent_link']
                    h_path = f'=HYPERLINK("#dlink=/d/f/{p_path}", "{mb["name"]}")'
 
                    try:
                        mail = synd.list_folder(mb['display_path'])
                    except:
                        print(f"Cannot get file list from {mb['name']}")
                    

                    # Checking files inside folder
                    error = False
                    for m in mail['data']['items']:
                        cont += 1
                        note = f"{mb['display_path']}/{m['name']}"
                        
                        print(f"    Copying {m['name']} to despacho")
                        copy_way = 'normal'
                        try:
                            synd.copy(note,f"{config['despacho']}/{m['name']}")
                        except:
                            try:
                                tmp_file = synd.download_file(note)
                                ret_upload = synd.upload_file(tmp_file, dest_folder_path=f"{config['despacho']}")
                                
                                """
                                try:
                                    if 'docx' in m['name'] or 'xlsx' in m['name']:
                                        ret_convert = synd.convert_to_online_office(ret_upload['data']['display_path'],
                                            delete_original_file=False,
                                            conflict_action='autorename')
                                """
                                copy_way = 'download'
                            except:
                                print("Cannot copy files")
                                if not error:
                                    ws.append([ctr,h_path,year,'',date,f"ERROR in {mb['display_path']}",'',''])
                                    error = True
                                continue
                            
                        
                        print(f"    Moving {m['name']} to archive")
                        try:
                            synd.move_path(note,f"{config['archive']}/ctr in {year}")
                        except:
                            print(f"Cannot move {m['name']}")
                            if not error:
                                ws.append([ctr,h_path,year,'',date,f"ERROR in {mb['display_path']}",'',''])
                                error = True
                            continue
                            

                        print("    Saving link in register")

                        try:
                            tmp_data = synd.get_file_or_folder_info(f'{config["archive"]}/ctr in {year}/{m["name"]}')['data']
                            p_link = tmp_data['permanent_link']
                            
                            if copy_way == 'normal':
                                h_link = f'=HYPERLINK("#dlink=/oo/r/{p_link}", "{m["name"]}")'
                            else:
                                h_link = f'=HYPERLINK("#dlink=/d/f/{p_link}", "{m["name"]}")'

                        except:
                            print("Cannot get link")
                            if not error:
                                ws.append([ctr,h_path,year,'',date,f"ERROR in {mb['display_path']}",'',''])
                                error = True
                            continue
                        
                        ws.append([ctr,h_link,year,'',date,'','',''])

    
    if cont > 0:
        file = NamedTemporaryFile()
        wb.save(file)
        file.seek(0)
        file.name = f"{date.replace('/','-')}-ctr-incoming.xlsx"
    
        print("Creating register file")
        uploaded = True
        try:
            ret_upload = synd.upload_file(file, dest_folder_path='/mydrive')
        except:
            print("Cannot upload register")
            wb.save(f"{date.replace('/','-')}-ctr-incoming.xlsx")
            uploaded = False

        if uploaded:
            try:
                ret_convert = synd.convert_to_online_office(ret_upload['data']['display_path'],
                    delete_original_file=True,
                    conflict_action='autorename')
            except:
                print("Cannot convert file to Synology Office")

    input("Pulse Enter to continue")
    #os.system("pause")

