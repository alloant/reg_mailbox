#!/bin/python

from getpass import getpass
from datetime import datetime
from tempfile import NamedTemporaryFile

from openpyxl import Workbook
from synology_drive_api.drive import SynologyDrive

from syno_tools import txt2dict, get_link, copy_to, move_to

def get_ctr(folder):
    if folder[:7] == 'Mailbox':
        return folder[8:]
    elif folder[:10] == '8.Mailbox-':
        return folder[10:]


config = txt2dict("config.txt")
year = datetime.today().strftime('%Y')
date_normal = datetime.today().strftime('%d-%m-%Y')
date = datetime.today().toordinal() - 693594
PASS = getpass()

with SynologyDrive(config['user'],PASS,"nas.prome.sg",dsm_version='7') as synd:
    wb = Workbook()
    ws = wb.active
    ws.append(['ctr','No','Year','Ref','Date','Content','Dept','#of docs'])
   
    cont = 0
    
    team_folders = synd.get_teamfolder_info()
    for team in team_folders:
        if 'Mailbox' in team:
            print(f"Checking {team}")
            
            ctr = get_ctr(team)

            # Getting notes from ctr to cr in t  ##############################
            try:
                if team[0] != '8':
                    team_data = synd.get_file_or_folder_info(f"/team-folders/{team}/{ctr} to cr")['data']
                    syn_notes = synd.list_folder(f"/team-folders/{team}/{ctr} to cr")['data']
                else:
                    team_data = synd.get_file_or_folder_info(f"/team-folders/{team}/{ctr}-to-cr")['data']
                    syn_notes = synd.list_folder(f"/team-folders/{team}/{ctr}-to-cr")['data']
                
                notes = syn_notes['items']
                team_link = get_link(team_data)
            except:
                print(f'Cannot get notes from {team}')
                break

            # Cheking all notes ########################################
            error = False
            for note in notes:
                cont += 1

                # First copy note to Despacho
                if not copy_to(synd,note,config['despacho'],convert=False):
                    if not error:
                        ws.append([ctr,team_link,year,'',date,f"ERROR in {team_data['display_path']}",'',''])
                        error = True
                    continue

                print(f"Moving {note['name']} to archive")
                rst,links = move_to(synd,note,f"{config['archive']}/ctr in {year}",team_data,convert=False)

                if not rst:
                    print(f"Cannot move {note['name']}")
                    if not error:
                        ws.append([ctr,team_link,year,'',date,f"ERROR in {team_data['display_path']}",'',''])
                        error = True
                    continue
                
                for ln in links:
                    ws.append([ctr,ln,year,'',date,'','',''])
    
    if cont > 0:
        file = NamedTemporaryFile()
        wb.save(file)
        file.seek(0)
        file.name = f"{date_normal}-ctr-incoming.xlsx"
    
        print("Creating register file")
        uploaded = True
        try:
            ret_upload = synd.upload_file(file, dest_folder_path='/mydrive')
        except:
            print("Cannot upload register")
            wb.save(f"{date_normal}-ctr-incoming.xlsx")
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
