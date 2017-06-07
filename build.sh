#!/usr/bin/env bash
pip install -r requirement.txt
pyinstaller -F -w -icon=logo.ico main.py