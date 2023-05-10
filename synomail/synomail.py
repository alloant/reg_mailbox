#!/bin/python
# -*- coding: utf-8 -*-

import sys
import os
import io

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QToolBar,
    QPushButton, QWidget,
    QMessageBox, QFileDialog,
    QPlainTextEdit, QLineEdit, QCheckBox,
    QHBoxLayout, QVBoxLayout
    )

from PySide6.QtCore import Qt, QSize, QDir, QSettings
from PySide6.QtGui import QKeySequence, QIcon, QAction

import logging

from synology_drive_api.drive import SynologyDrive

from synomail import _ROOT, CONFIG
from synomail.get_mail import init_get_mail
from synomail.get_notes_from_d import init_get_notes_from_d
from synomail.register_despacho import init_register_despacho
from synomail.send_mail import init_send_mail
from synomail.syno_tools import upload_file, download_file
# Uncomment below for terminal log messages
# logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(name)s - %(levelname)s - %(message)s')

class QTextEditLogger(logging.Handler):
    def __init__(self, parent):
        super().__init__()
        self.widget = QPlainTextEdit(parent)
        self.widget.setReadOnly(True)

    def emit(self, record):
        msg = self.format(record)
        self.widget.appendPlainText(msg)



class mainWindow(QMainWindow, QPlainTextEdit):
    def __init__(self):
        super(mainWindow,self).__init__()
        self.initUI()
    
    def new_action(self,icon_name,icon_path,name,status="",enable = True):
        icon = QIcon.fromTheme(icon_name, QIcon(os.path.join(_ROOT,icon_path)))
        act = QAction(icon, name.upper(), self)
        act.setObjectName(name)
        act.setShortcuts(QKeySequence.Open)
        act.setStatusTip(status)
        act.setEnabled(enable)
        act.triggered.connect(self.toolBarActions)

        return act
        
    def toolBar(self):
        self.toolBar = QToolBar(self.tr('File toolbar'), self)
        self.addToolBar(Qt.TopToolBarArea, self.toolBar)
        self.toolBar.setIconSize(QSize(22,22))
        
        self.le_pass = QLineEdit()
        self.le_pass.setMaximumWidth(200)
        self.le_pass.setEchoMode(QLineEdit.Password)
        self.le_pass.setObjectName('pass_return')
        self.le_pass.returnPressed.connect(self.toolBarActions)
        self.toolBar.addWidget(self.le_pass)
        
        self.toolBar.addAction(self.new_action('key','icons/key.svg','pass','Password'))

        self.toolBar.addSeparator()

        buttons = [] 
        buttons.append(['vcs-pull','icons/email-download.svg','get_mail','Get mail from cg, asr, r y ctr'])
        buttons.append(['vcs-pull','icons/file.svg','register','Register mail despacho and send message to d'])
        buttons.append(['vcs-pull','icons/send.svg','send','Send mail to ctr'])
        buttons.append(['vcs-pull','icons/block-up-bracket.svg','upload','Upload files from cg and asr'])
        buttons.append(['vcs-pull','icons/block-down-bracket.svg','download','Download files for cg and asr'])

        for but in buttons:
            self.toolBar.addAction(self.new_action(but[0],but[1],but[2],but[3],False))
        
        self.ck_debug = QCheckBox("DEBUG")
        self.ck_debug.setObjectName('debug')
        self.ck_debug.stateChanged.connect(self.toolBarActions)
        self.toolBar.addWidget(self.ck_debug)
        

    def toolBarActions(self,rst = None):
        sender = self.sender().objectName()
        if sender in ['pass','pass_return']:
            self.PASS = self.le_pass.text()
            self.le_pass.clear()
            for act in self.toolBar.actions():
                act.setEnabled(True)
        elif sender == 'get_mail':
            init_get_mail(self.PASS)
            init_get_notes_from_d(self.PASS)
        elif sender == 'register':
            init_register_despacho(self.PASS)
        elif sender == 'send':
            init_send_mail(self.PASS)
        elif sender == 'debug':
            if rst == 2:
                logging.getLogger().setLevel(logging.DEBUG)
            else:
                logging.getLogger().setLevel(logging.INFO)
        elif sender == 'download':
            logging.info('Downloading')
            folders = ['cg','asr','r']
            with SynologyDrive(CONFIG['user'],self.PASS,"nas.prome.sg",dsm_version='7') as synd: 
                for fd in folders:
                    notes = synd.list_folder(f"/team-folders/Mail {fd}/Mail to {fd}")['data']['items']
                    for note in notes:
                        logging.info(f"Downloading {note['display_path']}")
                        download_file(synd,note['display_path'],f"{CONFIG['local_folder']}/Mail {fd}/Mail to {fd}")
            
            logging.info('Downloading is over')

        elif sender == 'upload':
            logging.info('Uploading')
            folders = ['cg','asr','r']

            for fd in folders:
                mypath = f"{CONFIG['local_folder']}/Mail {fd}/Mail from {fd}"
                notes = [f for f in os.listdir(mypath) if os.path.isfile(os.path.join(mypath, f))]
                for note in notes:
                    with open(f"{mypath}/{note}",mode='rb') as nt:
                        file = io.BytesIO(nt.read())
                        file.name = note
                        logging.info(f"Uploading {note}")
                        upload_file(self.PASS,file,f"/team-folders/Mail {fd}/Mail from {fd}")
            
            logging.info('Uploading is over')

        
    def initUI(self):
        self.toolBar()

        logTextBox = QTextEditLogger(self)
        # You can format what is printed to text box
        logTextBox.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(logTextBox)
        # You can control the logging level
        logging.getLogger().setLevel(logging.INFO)

        self.centralWidget = QWidget()
        layout = QVBoxLayout(self.centralWidget)

        layout.addWidget(logTextBox.widget)
        
        self.setCentralWidget(self.centralWidget)
        self.statusBar().showMessage("Register Kamet")

def main_ui():
    app = QApplication(sys.argv)
    #settings = QSettings("alloant","quotes")

    #app.setStyleSheet(qdarktheme.load_stylesheet())

    ex = mainWindow()
    ex.setGeometry(100, 100, ex.width()+600, 600)
    ex.show()
    
    sys.exit(app.exec())

