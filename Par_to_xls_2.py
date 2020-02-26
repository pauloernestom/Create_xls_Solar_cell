import time
start_time = time.time()
import numpy as np
import matplotlib.pylab as plt
import pandas as pd
import os



def createDir(path, name):
    Dir = path + name
    if not os.path.exists(Dir):
        os.makedirs(Dir)
    return Dir

def findFiles(path, extension):
    files = []
    for i in os.listdir(path):
        if i.endswith(extension):
            files.append(path + str(i))
    files.sort()
    return files


def createEmptyDict(list_names):
    """
    Cria um dicionario com com as chaves
    cada chave corresponde a um array
    """
    dicionario = {}

    for i in range(0, len(list_names)):
        dicionario[list_names[i]] = []
    return dicionario

def organizeData():
    data = {}
    for i in range(1, len(cols)):
        data[cols[i]] = list()

    devices = []

    devices_list = []

    for i in range(0, len(files)):
        filename = files[i]

        with open(filename) as infile:
            devices_list.append(infile.readline()[:-1])

            for line in infile:
                if (line.find("Cell") == 0):
                    data['Cell / Measure'].append(str(line))
                    devices.append(open(filename).readline()[:-1])
                elif (line.find("Célula") == 0):
                    data['Cell / Measure'].append(str(line))
                    devices.append(open(filename).readline()[:-1])
                elif (line.find("Voc=") == 0):
                    data['Voc / V'].append(float(line.split(' ')[1]))
                elif (line.find("Jsc=") == 0):
                    data['Jsc / mAcm-2'].append(float(line.split(' ')[1]))
                elif (line.find("FF=") == 0):
                    data['FF / %'].append(float(line.split(' ')[1]))
                elif (line.find("n =") == 0):
                    data['PCE / %'].append(float(line.split(' ')[2]))
                elif (line.find("n=") == 0):
                    data['PCE / %'].append(float(line.split(' ')[1]))
    data
    #data['device']=devices
    data2 = data
    data2['device'] = devices

    df = pd.DataFrame(data, devices, columns=cols[1:])

    df2 = pd.DataFrame(data2, columns=cols)

    dict_devices = createEmptyDict(devices_list)
    for i in dict_devices:
        dict_devices[i] = df.loc[i]

    return (devices, devices_list, df, df2, dict_devices)


def createTab_med():
    dict_eff = createEmptyDict(sentido)

    for a in range(0, len(df2['Cell / Measure'])):
        if df2['Cell / Measure'][a].find('Forward') != -1 and df2['PCE / %'][a] > 0:
            dict_eff['Forward'].append(df2.loc[[a]])
        elif df2['Cell / Measure'][a].find('Reverse') != -1 and df2['PCE / %'][a] > 0:
            dict_eff['Reverse'].append(df2.loc[[a]])

    for i in range(0, len(sentido)):
        dict_eff[sentido[i]] = pd.concat(dict_eff[sentido[i]])

    med_eff = {}
    for i in range(0, len(sentido)):
        med_eff[sentido[i]] = {}
        for a in range(0, len(devices_list)):
            med_eff[sentido[i]][devices_list[a].split('_')[0]] = []

    for i in range(0, len(sentido)):
        for a in med_eff[sentido[i]]:
            for w, z in zip(dict_eff[sentido[i]]['device'], dict_eff[sentido[i]].index):
                if w.find(a) != -1:
                    med_eff[sentido[i]][a].append(dict_eff[sentido[i]].loc[[z]])

    for i in range(0, len(sentido)):
        for a in med_eff[sentido[i]]:
            med_eff[sentido[i]][a] = pd.concat(med_eff[sentido[i]][a])

    tab_med = {}
    for a in sentido:
        for i in med_eff[a]:
            tab_med[a] = {}
            tab_med[a + '-std_dev'] = {}

    for a in sentido:
        for i in med_eff[a]:
            tab_med[a][i] = med_eff[a][i].mean()
            tab_med[a + '-std_dev'][i] = med_eff[a][i].std()

    for i in tab_med:
        tab_med[i] = np.transpose(pd.DataFrame.from_dict(tab_med[i]))

    return tab_med

def save_xls(pathTab, dataFrame, fileName):
    writer = pd.ExcelWriter(pathTab + '/' + fileName, engine='xlsxwriter')
    dataFrame.to_excel(writer, sheet_name='Sheet1')
    writer.save()


path = ""


pathTab = createDir(path, 'data_xls')


files = findFiles(path, ".txt")


cols = ['device','Cell / Measure','Voc / V', 'Jsc / mAcm-2', 'FF / %', 'PCE / %']
sentido = ['Forward', 'Reverse']


devices, devices_list, df, df2, dict_devices = organizeData()

save_xls(pathTab, df, 'all_cells.xlsx')

tab_med = createTab_med()

for i in dict_devices:
    save_xls(pathTab, dict_devices[i], i + '.xlsx')


for i in tab_med:
    for a in tab_med[i]:
        save_xls(pathTab, tab_med[i], 'Med_' + i + '.xlsx')
