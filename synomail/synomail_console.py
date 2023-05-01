#!/bin/python
# -*- coding: utf-8 -*-

import sys
import os

from getpass import getpass

import logging

from synomail import _ROOT
from synomail.get_mail import init_get_mail
from synomail.get_notes_from_d import init_get_notes_from_d
from synomail.register_despacho import init_register_despacho
from synomail.send_mail import init_send_mail

def main_console(argv):
    PASS = getpass()
    if argv == 'get':
        init_get_mail(PASS)
        init_get_notes_from_d(PASS)
    elif argv == 'despacho':
        init_register_despacho(PASS)
    elif argv == 'send':
        init_send_mail(PASS)

