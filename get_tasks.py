import pandas as pd
import os
from pathlib import Path
from win32com.client import Dispatch



tasks = []
with open('Default.ics', 'r') as file:
    content = ''.join(file.readlines())
    
entries = content.split('BEGIN:VTODO\n')
entries = [entry.replace('\nEND:VTODO\n', '') for entry in entries if '\nEND:VTODO\n' in entry]
jobs = []
for entry in entries:
    if 'URL' in entry:
        job_dict = {}
        properties_list = entry.split('\n')
        properties_split = [prop.split(':', 1) for prop in properties_list]
        for prop in properties_split:
            if prop[0] in ['URL', 'SUMMARY'] and len(prop[1]) > 3:
                job_dict[prop[0]] = prop[1]
        if 'URL' in job_dict.keys():
            jobs.append(job_dict)

for job in jobs:
    path = Path(os.getcwd(), '{}.lnk'.format(job['SUMMARY'].replace(':', ''))).as_posix()
    target = Path(job['URL']).as_posix()

    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.save()