# -*- coding: utf-8 -*-
"""
Created on Mon Dec  3 14:23:31 2018

@author: Paulo
"""

import time
start_time = time.time()
import numpy as np
import matplotlib.pylab as plt
import pandas as pd
import os




#path = '/home/paulo/Dropbox/Doutorado/Resultados/PL/Dados_elipso/'
path = "C:/Users/Paulo/Documents/MEGA/program/create_xls/Par/"


pathTab = path + '/data_xls/'

if not os.path.exists(pathTab):
    os.makedirs(pathTab)

files = []
for i in os.listdir(path):
    if i.endswith(".txt"):
         files.append(path + str(i))
files.sort()
cols = ['Cell / Measure','Voc / V', 'Jsc / mAcm-2', 'FF / %', 'PCE / %']
data = {}

for i in range(0, len(cols)):
    data[cols[i]]=list()
    
devices = []

devices_list = []

for i in range (0, len(files)):
    filename = files[i]

    
    with open(filename) as infile:
        devices_list.append(infile.readline()[:-1])
        for line in infile:
            if (line.find("Célula")==0):
                data['Cell / Measure'].append(line)
                devices.append(open(filename).readline()[:-1])
            elif (line.find("Voc=")==0):
                data['Voc / V'].append(float(line.split(' ')[1]))
            elif (line.find("Jsc=")==0):
                data['Jsc / mAcm-2'].append(float(line.split(' ')[1]))
            elif (line.find("FF=")==0):
                data['FF / %'].append(float(line.split(' ')[1]))
            elif (line.find("n=")==0):
                data['PCE / %'].append(float(line.split(' ')[1]))



df = pd.DataFrame(data, devices, columns = cols)

writer = pd.ExcelWriter(pathTab + 'all_cells.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='Sheet1')

writer.save()

dict_devices={}
for i in range(0,len(devices_list)):
    dict_devices[devices_list[i]]=list()
    
for i in dict_devices:
    
    dict_devices[i]=df.loc[i]
    
for i in dict_devices:
    writer2=pd.ExcelWriter(pathTab + i + '.xlsx', engine='xlsxwriter')
    dict_devices[i].to_excel(writer2, sheet_name=i)
    writer2.save()




print('\n  --- %s seconds --- \n'% (time.time() - start_time))