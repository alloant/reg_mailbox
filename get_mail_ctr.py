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
    ws.append(['source','No','Year','Ref','Date','Content','Dept','#of docs'])
   
    cont = 0
    
    team_folders = synd.get_teamfolder_info()
    team_selected = []

    for team in team_folders:
        if 'Mailbox gul' in team:
            ctr = get_ctr(team)
            if team[0] != '8':
                source = f"/team-folders/{team}/{ctr} to cr"
            else:
                source = f"/team-folders/{team}/{ctr}-to-cr"
            dest = f"{config['archive']}/ctr in 2023"
        
            team_selected.append({'name':team,'source':source,'dest':dest,'source_name':ctr})

    team_selected.append({'name':'Mail asr','source':'/team-folders/Mail asr/Mail from asr','dest':f'{config["archive"]}/sf in {year}','source_name':'asr'})
    team_selected.append({'name':'Mail cg','source':'/team-folders/Mail cg and r/Mail from cg','dest':f'{config["archive"]}/cg in {year}','source_name':'cg'})
    team_selected.append({'name':'Mail r','source':'/team-folders/Mail cg and r/Mail from r','dest':f'{config["archive"]}/r in {year}','source_name':''})


    for team in team_selected:
        print(f"Checking {team['name']}")
            
        # Getting notes from ctr to cr in t  ##############################
        try:
            team_data = synd.get_file_or_folder_info(team['source'])['data']
            syn_notes = synd.list_folder(team['source'])['data']
                
            notes = syn_notes['items']
            team_link = get_link(team_data)
        except:
            print(f'Cannot get notes from {team["name"]}')
            continue

        # Cheking all notes ########################################
        error = False
        for note in notes:
            cont += 1

            ctr = team['source_name']
            if ctr == '':
                ctr = note['name'].split('.')[0]

            # First copy note to Despacho
            if not copy_to(synd,note,config['despacho'],convert=False):
                if not error:
                    ws.append([ctr,team_link,year,'',date,f"ERROR in {team_data['display_path']}",'',''])
                    error = True
                continue

            print(f"Moving {note['name']} to archive")
            rst,links = move_to(synd,note,f"{team['dest']}",team_data,convert=False)

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
