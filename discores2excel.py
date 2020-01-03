#!/usr/bin/env python2.7
# -*- coding: utf-8 -*-
"""
Created on Thu Jan 02 14:31:47 2019

Currently running under python2.7 with numpy V1.13 
because that was easiest to compile for windows and it works.

Cross-compiling with the following as basis:
https://www.andreafortuna.org/2017/12/27/how-to-cross-compile-a-python-script-into-a-windows-executable-on-linux/

@author: Sebastian Frisch
"""

import json
import os
from pathlib import Path
import xlsxwriter
import numpy as np
import zipfile
from Tkinter import Tk
from tkFileDialog import askopenfilename
from tkFileDialog import askdirectory
import sys

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
zipfilepath = askopenfilename(title='Choose discores.zip') # show an "Open" dialog box and return the path to the selected file
save_path = askdirectory(title='Choose directory for excel files')

if len(zipfilepath) == 0:
    sys.exit()
elif len(save_path) == 0:
    sys.exit()

#open zipfile with data in it
zfile = zipfile.ZipFile(zipfilepath)
for finfo in zfile.infolist():
    if finfo.filename == 'games.json':
        ifile = zfile.open(finfo)
        games_data = json.load(ifile)
    elif finfo.filename == 'players.json':
        ifile = zfile.open(finfo)
        players_data = json.load(ifile)
    elif finfo.filename == 'courses.json':
        ifile = zfile.open(finfo)
        courses_data = json.load(ifile)

collectedData = {}
    
for PLAYER_ID in range(len(players_data['players'])):
    
    print('Getting data for: '+str(players_data['players'][PLAYER_ID]['name']))
    target_uuid = players_data['players'][PLAYER_ID]['uuid']
        
    
    currentCourse = None
    currentHole = None
    currentPar = None
    currentScore = None
    
    collectedData[str(players_data['players'][PLAYER_ID]['name'])] = {}
    
    # go through all the scores
    for score in games_data['scores']:
        
        # check who the score belongs to
        for gamePlayer in games_data['gamePlayers']:
            if gamePlayer['uuid'] == score['gamePlayerUuid']:
                if gamePlayer['playerUuid'] == target_uuid:
                    # this score belongs to target
                    currentScore = score['score']
                    
                    # check which course this score belongs to
                    for game in games_data['games']:
                        if game['uuid'] == score['gameUuid']:
                           #game['courseUuid'] is current course
                           for course in courses_data['courses']:
                               if course['uuid'] == game['courseUuid']:
                                   currentCourse = course['name']
                           
                    for gameHole in games_data['gameHoles']:
                        if gameHole['uuid'] == score['gameHoleUuid']:
                            currentHole = gameHole['hole']
                            currentPar = gameHole['par']
                            
                    # add collected data to own dictionary        
                    if str(currentCourse) not in collectedData[str(players_data['players'][PLAYER_ID]['name'])]:
                        collectedData[str(players_data['players'][PLAYER_ID]['name'])][str(currentCourse)] = {'pars': [], 'scores': []}
                    if currentHole > len(collectedData[str(players_data['players'][PLAYER_ID]['name'])][str(currentCourse)]['scores']):
                        collectedData[str(players_data['players'][PLAYER_ID]['name'])][str(currentCourse)]['scores'].append([currentScore])
                        collectedData[str(players_data['players'][PLAYER_ID]['name'])][str(currentCourse)]['pars'].append(currentPar)
                    else:
                        collectedData[str(players_data['players'][PLAYER_ID]['name'])][str(currentCourse)]['scores'][currentHole-1].append(currentScore)
                        collectedData[str(players_data['players'][PLAYER_ID]['name'])][str(currentCourse)]['pars'][currentHole-1] = currentPar
                    

# save collected data to txt files
if not os.path.exists(str(Path(save_path+os.path.sep+'discgolf_data'))):
        os.mkdir(str(Path(save_path+os.path.sep+'discgolf_data')))

for name in collectedData:
    workbook = xlsxwriter.Workbook(str(Path(save_path+os.path.sep+'discgolf_data'+os.path.sep+name+'.xlsx')))
    #bold = workbook.add_format({'bold': True})
    right_border = workbook.add_format({'right': 2})
    top_border = workbook.add_format({'top': 2})
    
    for course in collectedData[name]:
        worksheet = workbook.add_worksheet(course)
        worksheet.set_column('A:D', 10) #adjust column width
        worksheet.set_column('E:E', 0.15) #adjust column width
        worksheet.set_column('F:ZZ', 3) #adjust column width
        worksheet.write('A1', course)
        worksheet.write('A2', 'Hole')
        worksheet.write('B2', 'Par')
        worksheet.write('C2', 'Average')
        worksheet.write('D2', 'Best')
        worksheet.write('E2', '', right_border)
        worksheet.write('F2', 'Scores')
        
        index = 0
        sumPars = 0
        sumAvg = 0
        sumBest = 0
        sumHole = []
        for hole in collectedData[name][course]['scores']:
            index += 1
            
            if index == 1:
                sumHole = [0]*len(hole)
            
            #prep data: subtract par from each score
            hole = [tmp-collectedData[name][course]['pars'][index-1] for tmp in hole]
            sumPars += collectedData[name][course]['pars'][index-1]
            sumAvg += np.mean(np.array(hole))
            sumBest += np.min(np.array(hole))
            sumHole = list(map(sum, zip(sumHole, hole)))
            
            worksheet.write('A'+str(index+2), index)
            worksheet.write('B'+str(index+2), collectedData[name][course]['pars'][index-1])
            worksheet.write('C'+str(index+2), round(float(np.mean(np.array(hole))), 2))
            worksheet.write('D'+str(index+2), round(float(np.min(np.array(hole))), 2))
            worksheet.write('E'+str(index+2), '', right_border)
            worksheet.write_row('F'+str(index+2), hole)
            
            if index == len(collectedData[name][course]['pars']):
                worksheet.write('B'+str(index+3), sumPars, top_border)
                worksheet.write('C'+str(index+3), round(sumAvg, 2), top_border)
                worksheet.write('D'+str(index+3), round(sumBest, 2), top_border)
                worksheet.write('E'+str(index+3), '', right_border)
                worksheet.write_row('F'+str(index+3), sumHole, top_border)
            
            #insert chart for each hole
            #chart = workbook.add_chart({'type': 'line'})
            #chart.add_series({'values': '='+course+'!$D$'+str(index+2)+':$'+chr(ord('D')+len(hole))+'$'+str(index+2)})
            #worksheet.insert_chart('O'+str(index+2), chart)
    
    workbook.close()


