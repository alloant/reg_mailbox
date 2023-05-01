#!/bin/python
# -*- coding: utf-8 -*-

import sys
import os

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QToolBar,
    QPushButton, QWidget,
    QMessageBox, QFileDialog,
    QPlainTextEdit, QLineEdit,
    QHBoxLayout, QVBoxLayout
    )

from PySide6.QtCore import Qt, QSize, QDir, QSettings
from PySide6.QtGui import QKeySequence, QIcon, QAction

import logging

from synomail import _ROOT
from synomail.get_mail import init_get_mail
from synomail.get_notes_from_d import init_get_notes_from_d
from synomail.register_despacho import init_register_despacho
from synomail.send_mail import init_send_mail
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

        for but in buttons:
            self.toolBar.addAction(self.new_action(but[0],but[1],but[2],but[3],False))
        

    def toolBarActions(self):
        sender = self.sender().objectName()
        if sender == 'test':
            logging.debug('damn, a bug')
            logging.info('something to remember')
            logging.warning('that\'s not right')
            logging.error('foobar')
        elif sender in ['pass','pass_return']:
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


    def test(self):
        logging.debug('damn, a bug')
        logging.info('something to remember')
        logging.warning('that\'s not right')
        logging.error('foobar')

def main_ui():
    app = QApplication(sys.argv)
    #settings = QSettings("alloant","quotes")

    #app.setStyleSheet(qdarktheme.load_stylesheet())

    ex = mainWindow()
    ex.setGeometry(100, 100, ex.width()+600, 600)
    ex.show()
    
    sys.exit(app.exec())

