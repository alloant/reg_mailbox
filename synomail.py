#!/bin/python
# -*- coding: utf-8 -*-

import sys


if len(sys.argv) <= 1:
    from synomail.synomail import main_ui
    main_ui()
else:
    from synomail.synomail_console import main_console
    main_console(sys.argv[1])
