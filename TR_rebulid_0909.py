import os
import zipfile
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
import winreg
from dbfread import DBF
import numpy as np

#openfilename = ""
def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key, "Desktop")[0]

def get_LIST(list):
    a = []
    b = []
    for x in list:
        if x[-4:] == '.zip':
            a.append(x)
            b.append(x)
        if x[-4:] == '.csv':
            for y in range(8,len(x)):
                if x[y] == 'T':
                    if x[:y-1] not in a:
                        a.append(x[:y-1])
                    if x[y+2:] not in b:
                        b.append(x[y+2:])
    return a,b

def get_T(file):
    for x in range(10,len(file)):
        if file[x] == 'T':
            break
    return file[x:x+2] 

def open_file():
    folder_selected = fd.askdirectory()
    allFileList = os.listdir(folder_selected)
    for file in allFileList:
        if file[0:2] != "TR":
            z = zipfile.ZipFile(folder_selected + '/' + file, "r")
            z.extractall(folder_selected)
    if os.path.exists(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案') == False:
        os.mkdir(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案')
        fp = get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案'
    elif os.path.exists(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-1') == False:
        os.mkdir(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-1')
        fp = get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-1'
    elif os.path.exists(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-2') == False:
        os.mkdir(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-2')
        fp = get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-2'
    elif os.path.exists(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-3') == False:
        os.mkdir(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-3')
        fp = get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-3'
    elif os.path.exists(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-4') == False:
        os.mkdir(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-4')
        fp = get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-4'
    elif os.path.exists(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-5') == False:
        os.mkdir(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-5')
        fp = get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-5'
    elif os.path.exists(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-6') == False:
        os.mkdir(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-6')
        fp = get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-6'
    elif os.path.exists(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-7') == False:
        os.mkdir(get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-7')
        fp = get_desktop() + '/' + 'ERIS可以丟到JMP裡面的檔案-7'
    
    allFileList = os.listdir(folder_selected)
    newFileList = get_LIST(allFileList)[0]
    #print(get_LIST(allFileList)[1])
    for file in newFileList:
        #print(file[:-4])
        if file[-4:] != ".zip":
            print(file)
            finaldata = pd.DataFrame()
            dd = {'-T1':h01.get(),
                    '-T2':h02.get(),
                    '-T3':h03.get(),
                    '-T4':h04.get(),
                    '-T5':h05.get(),
                    '-T6':h06.get()}
            ddd = []
            for x in dd:
                if dd[x] == True:
                    if os.path.isfile(folder_selected + '/' + file + x + '.eris.csv'):
                        ddd.append(file + x + '.eris.csv')
                    else:
                        ddd.append(file + x + '.csv')
                    
            #print(ddd)
            for file2 in ddd:
                if os.path.isfile(folder_selected + '/' + file2):
                    data = pd.read_csv(folder_selected + '/' + file2,delimiter="\t",header = None)
                    data = data[0].str.split(',',expand=True)
                    while data[0][0] != "No":
                        data = data.drop(0)
                        data = data.reset_index(drop=True)
                    for x in range(0,data.shape[1]):
                        data.rename(columns={x:data[x][0]},inplace=True)
                    data = data.drop(0)
                    data = data.reset_index(drop=True)
                    if get_T(file2) == "T1" and h01.get() == True:
                        #GPP TVS單向 TVS雙向 SKY
                        if (data.columns == ['No', 'Bin', 'TRR N(nS)', 'VF N(mV)', 'VB N(V)', 'IR N(uA)','DVR1 N(V)', 'IR% N(uA)', 'TRR R(nS)', 'VF R(mV)', 'VB R(V)','IR R(uA)', 'DVR1 R(V)', 'IR% R(uA)', '']).all() == True:
                            if h2.get() == True:#VF
                                data.rename(columns={'VF N(mV)':'T1 VF'}, inplace=True)
                                data['T1 VF'] = pd.to_numeric(data['T1 VF'], errors='ignore')
                                if np.mean(data['T1 VF']) != 0:
                                    finaldata = pd.concat([finaldata,data['T1 VF']],axis = 1)
                            if h3.get() == True:#VB
                                data.rename(columns={'VB N(V)':'T1 VB'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T1 VB']],axis = 1)
                            if h4.get() == True:#IR
                                data.rename(columns={'IR N(uA)':'T1 IR'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T1 IR']],axis = 1)
                            if h5.get() == True:#DVR
                                data.rename(columns={'DVR1 N(V)':'T1 DVR'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T1 DVR']],axis = 1)

                    if get_T(file2) == "T2" and h02.get() == True:
                        #GPP
                        if (data.columns == ['No', 'Bin', 'TRR-1 (nS)', 'TRR-3 (nS)', '', '', '', '', 'TRR-2 (nS)','TRR-4 (nS)', '', '', '', '', '']).all() == True:
                            if h1.get() == True:#TRR
                                data.rename(columns={'TRR-1 (nS)':'T2 TRR'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T2 TRR']],axis = 1)
                        #TVS單向
                        if (data.columns == ['No', 'Bin', 'VC N(V)', 'IP N(A)', 'Vcap N(V)', '', '', '', 'VC R(V)','IP R(A)', 'Vcap R(V)', '', '', '', '']).all() == True:
                            if h6.get() == True:#VC
                                data.rename(columns={'VC N(V)':'T2 VC'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T2 VC']],axis = 1)
                        #TVS雙向 SKY
                        if (data.columns == ['No', 'Bin', 'TRR N(nS)', 'VF N(mV)', 'VB N(V)', 'IR N(uA)','DVR1 N(V)', 'IR% N(uA)', 'TRR R(nS)', 'VF R(mV)', 'VB R(V)','IR R(uA)', 'DVR1 R(V)', 'IR% R(uA)', '']).all() == True:
                            if h2.get() == True:#VF
                                data.rename(columns={'VF N(mV)':'T2 VF'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T2 VF']],axis = 1)
                            if h3.get() == True:#VB
                                if np.mean(data['VB N(V)']) != 0:
                                    data.rename(columns={'VB N(V)':'T2 VB'}, inplace=True)
                                    finaldata = pd.concat([finaldata,data['T2 VB']],axis = 1)
                                if np.mean(data['VB R(V)']) != 0:
                                    data.rename(columns={'VB R(V)':'T2 VB'}, inplace=True)
                                    finaldata = pd.concat([finaldata,data['T2 VB']],axis = 1)
                            if h4.get() == True:#IR
                                if np.mean(data['IR N(uA)']) != 0:
                                    data.rename(columns={'IR R(uA)':'T2 IR'}, inplace=True)
                                    finaldata = pd.concat([finaldata,data['T2 IR']],axis = 1)
                                if np.mean(data['IR R(uA)']) != 0:
                                    data.rename(columns={'IR R(uA)':'T2 IR'}, inplace=True)
                                    finaldata = pd.concat([finaldata,data['T2 IR']],axis = 1)
                            if h5.get() == True:#DVR
                                data.rename(columns={'DVR1 R(V)':'T2 DVR'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T2 DVR']],axis = 1) 
                    if get_T(file2) == "T3" and h03.get() == True:
                        #GPP TVS單向 SKY
                        if (data.columns == ['No', 'Bin', 'VF N(mV)', 'VF1 N(mV)', 'VF2 N(mV)', 'DVF N(mV)','VFT N(mV)', '', 'VF R(mV)', 'VF1 R(mV)', 'VF2 R(mV)', 'DVF R(mV)','VFT R(mV)', '', '']).all() == True:
                            if h2.get() == True:#VF
                                data.rename(columns={'VF1 N(mV)':'T3 VF1'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T3 VF1']],axis = 1)

                                data.rename(columns={'VF2 N(mV)':'T3 VF2'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T3 VF2']],axis = 1)

                            if h7.get() == True:#DVF
                                data.rename(columns={'DVF N(mV)':'T3 DVF'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T3 DVF']],axis = 1)
                        #TVS雙向
                        if (data.columns == ['No', 'Bin', 'VC N(V)', 'IP N(A)', 'Vcap N(V)', '', '', '', 'VC R(V)','IP R(A)', 'Vcap R(V)', '', '', '', '']).all() == True:
                            if h6.get() == True:#VC
                                data.rename(columns={'VC N(V)':'T3 VC'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T3 VC']],axis = 1)
                    if get_T(file2) == "T4" and h04.get() == True:
                        #雙向
                        if (data.columns == ['No', 'Bin', 'VC N(V)', 'IP N(A)', 'Vcap N(V)', '', '', '', 'VC R(V)','IP R(A)', 'Vcap R(V)', '', '', '', '']).all() == True:
                            if h6.get() == True:#VC
                                data.rename(columns={'VC R(V)':'T4 VC'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T4 VC']],axis = 1)
                        #SKY
                        if (data.columns == ['No', 'Bin', 'VR N(V)', 'IRt N(A)', 'Vcap N(V)', '', '', '', 'VR R(V)','IRt R(A)', 'Vcap R(V)', '', '', '', '']).all() == True:
                            if h3.get() == True:#VB VR
                                data.rename(columns={'VR N(V)':'T4 VR'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T4 VR']],axis = 1)
                                
                    if get_T(file2) == "T5" and h05.get() == True:
                        #GPP TVS單向 TVS雙向 SKY
                        if (data.columns == ['No', 'Bin', 'TRR N(nS)', 'VF N(mV)', 'VB N(V)', 'IR N(uA)','DVR1 N(V)', 'IR% N(uA)', 'TRR R(nS)', 'VF R(mV)', 'VB R(V)','IR R(uA)', 'DVR1 R(V)', 'IR% R(uA)', '']).all() == True:
                            if h2.get() == True:#VF
                                data.rename(columns={'VF N(mV)':'T5 VF'}, inplace=True)
                                data['T5 VF'] = pd.to_numeric(data['T5 VF'], errors='ignore')
                                if np.mean(data['T5 VF']) != 0:
                                    finaldata = pd.concat([finaldata,data['T5 VF']],axis = 1)
                            if h3.get() == True:#VB
                                data.rename(columns={'VB N(V)':'T5 VB1'}, inplace=True)
                                data['T5 VB1'] = pd.to_numeric(data['T5 VB1'], errors='ignore')
                                if np.mean(data['T5 VB1']) != 0:
                                    finaldata = pd.concat([finaldata,data['T5 VB1']],axis = 1)

                                data.rename(columns={'VB R(V)':'T5 VB2'}, inplace=True)
                                data['T5 VB2'] = pd.to_numeric(data['T5 VB2'], errors='ignore')
                                if np.mean(data['T5 VB2']) != 0:
                                    finaldata = pd.concat([finaldata,data['T5 VB2']],axis = 1)
                    if get_T(file2) == "T6" and h06.get() == True:
                        #GPP TVS單向 TVS雙向 SKY
                        if (data.columns == ['No', 'Bin', 'TRR N(nS)', 'VF N(mV)', 'VB N(V)', 'IR N(uA)','DVR1 N(V)', 'IR% N(uA)', 'TRR R(nS)', 'VF R(mV)', 'VB R(V)','IR R(uA)', 'DVR1 R(V)', 'IR% R(uA)', '']).all() == True:
                            if h2.get() == True:#VF
                                data.rename(columns={'VF N(mV)':'T6 VF'}, inplace=True)
                                data['T6 VF'] = pd.to_numeric(data['T6 VF'], errors='ignore')
                                if np.mean(data['T6 VF']) != 0:
                                    finaldata = pd.concat([finaldata,data['T6 VF']],axis = 1)
                            if h4.get() == True:#IR
                                data.rename(columns={'IR N(uA)':'T6 IR1'}, inplace=True)
                                finaldata = pd.concat([finaldata,data['T6 IR1']],axis = 1)

                                data.rename(columns={'IR R(uA)':'T6 IR2'}, inplace=True)
                                data['T6 IR2'] = pd.to_numeric(data['T6 IR2'], errors='ignore')
                                if np.mean(data['T6 IR2']) != 0:
                                    finaldata = pd.concat([finaldata,data['T6 IR2']],axis = 1)
            #print(finaldata)

            finaldata['TR'] = file
            mm = finaldata.columns
            mm = pd.DataFrame(mm).T
            mm.columns = finaldata.columns
            finaldata = pd.concat([mm,finaldata], axis=0 ,ignore_index=True)
            finaldata.to_csv(fp + '/' + file + '.csv',index=0,header=0)


        if file[-4:] == ".zip" and file[0:2] == "TR":
            print(file[:-4])
            z = zipfile.ZipFile(folder_selected + '/' + file, "r")
            #cfg.dbf.prd.pat
            z.extract(z.namelist()[2])
            data_prd = open(z.namelist()[2], mode='r')
            data_prd = pd.DataFrame(data_prd)
            os.remove(z.namelist()[2])
            a = data_prd[0][3][16:30]
            a = a.rstrip()
            finaldata = pd.DataFrame()
            if a[-2:] == "CA":
                z.extract(z.namelist()[1])
                dbf = DBF(z.namelist()[1])
                frame = pd.DataFrame(iter(dbf))[['TESTERNO','RXDATA1','RXDATA2','RXDATA3','RXDATA4','RXDATA5','RXDATA6','RXDATA7','RXDATA8','RXDATA9','RXDATA10','RXDATA11']]
                os.remove(z.namelist()[1])
                if h3.get() == True:#VB
                    if h01.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 1].copy()
                        frame_tmp.rename(columns={'RXDATA3':'T1 VB'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T1 VB']],axis = 1)
                    if h02.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 2].copy()
                        frame_tmp.rename(columns={'RXDATA9':'T2 VB'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T2 VB']],axis = 1)
                    if h05.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 5].copy()
                        frame_tmp.rename(columns={'RXDATA3':'T5 VB1'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T5 VB1']],axis = 1)

                        frame_tmp = frame.loc[frame.TESTERNO == 5].copy()
                        frame_tmp.rename(columns={'RXDATA9':'T5 VB2'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T5 VB2']],axis = 1)
                if h4.get() == True:#IR
                    if h01.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 1].copy()
                        frame_tmp.rename(columns={'RXDATA4':'T1 IR'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T1 IR']],axis = 1)
                    if h02.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 2].copy()
                        frame_tmp.rename(columns={'RXDATA10':'T2 IR'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T2 IR']],axis = 1)
                    if h06.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 6].copy()
                        frame_tmp.rename(columns={'RXDATA4':'T6 IR1'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T6 IR1']],axis = 1)

                        frame_tmp = frame.loc[frame.TESTERNO == 6].copy()
                        frame_tmp.rename(columns={'RXDATA10':'T6 IR2'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T6 IR2']],axis = 1)
                if h5.get() == True:#DVR
                    if h01.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 1].copy()
                        frame_tmp.rename(columns={'RXDATA5':'T1 DVR'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T1 DVR']],axis = 1)
                if h6.get() == True:#VC
                    if h03.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 3].copy()
                        frame_tmp.rename(columns={'RXDATA1':'T3 VC'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T3 VC']],axis = 1)
                    if h04.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 4].copy()
                        frame_tmp.rename(columns={'RXDATA7':'T4 VC'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T4 VC']],axis = 1)
                finaldata['TR'] = file[:-4]
                mm = finaldata.columns
                mm = pd.DataFrame(mm).T
                mm.columns = finaldata.columns
                finaldata = pd.concat([mm,finaldata], axis=0 ,ignore_index=True)
                finaldata.to_csv(fp + '/' + file[:-4] + '.csv',index=0,header=0)


            elif a[-1:] == "A":
                z.extract(z.namelist()[1])
                dbf = DBF(z.namelist()[1])
                frame = pd.DataFrame(iter(dbf))[['TESTERNO','RXDATA1','RXDATA2','RXDATA3','RXDATA4','RXDATA5']]
                #print(frame)
                os.remove(z.namelist()[1])
                if h3.get() == True:#VB
                    if h01.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 1].copy()
                        frame_tmp.rename(columns={'RXDATA3':'T1 VB'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T1 VB']],axis = 1)
                    if h05.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 5].copy()
                        frame_tmp.rename(columns={'RXDATA3':'T5 VB'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T5 VB']],axis = 1)
                if h4.get() == True:#IR
                    if h01.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 1].copy()
                        frame_tmp.rename(columns={'RXDATA4':'T1 IR'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T1 IR']],axis = 1)
                    if h06.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 6].copy()
                        frame_tmp.rename(columns={'RXDATA4':'T6 IR'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T6 IR']],axis = 1)
                if h5.get() == True:#DVR
                    if h01.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 1].copy()
                        frame_tmp.rename(columns={'RXDATA5':'T1 DVR'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T1 DVR']],axis = 1)
                if h6.get() == True:#VC
                    if h02.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 2].copy()
                        frame_tmp.rename(columns={'RXDATA1':'T2 VC'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T2 VC']],axis = 1)
                if h7.get() == True:#DVF
                    if h03.get() == True:
                        frame_tmp = frame.loc[frame.TESTERNO == 3].copy()
                        frame_tmp.rename(columns={'RXDATA4':'T3 DVF'}, inplace=True)
                        frame_tmp.index = range(len(frame_tmp))
                        finaldata = pd.concat([finaldata,frame_tmp['T3 DVF']],axis = 1)
                finaldata['TR'] = file[:-4]
                mm = finaldata.columns
                mm = pd.DataFrame(mm).T
                mm.columns = finaldata.columns
                finaldata = pd.concat([mm,finaldata], axis=0 ,ignore_index=True)
                finaldata.to_csv(fp + '/' + file[:-4] + '.csv',index=0,header=0)
            
            else:
                z.extract(z.namelist()[1])
                dbf = DBF(z.namelist()[1])
                frame = pd.DataFrame(iter(dbf))[['TESTERNO','RXDATA1','RXDATA2','RXDATA3','RXDATA4','RXDATA5']]
                os.remove(z.namelist()[1])
                #分GPP SKY
                if np.mean(frame['RXDATA5']) != 0:
                    if h1.get() == True:#TRR
                        if h02.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 2].copy()
                            frame_tmp.rename(columns={'RXDATA1':'T2 TRR'}, inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T2 TRR']],axis = 1)
                    if h2.get() == True:#VF
                        if h01.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 1].copy()
                            frame_tmp.rename(columns={'RXDATA2':'T1 VF'}, inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T1 VF']],axis = 1)
                        if h05.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 5].copy()
                            frame_tmp.rename(columns={'RXDATA2':'T5 VF'}, inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T5 VF']],axis = 1)
                        if h06.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 6].copy()
                            frame_tmp.rename(columns={'RXDATA2':'T6 VF'}, inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T6 VF']],axis = 1)

                    if h3.get() == True:#VB
                        if h01.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 1].copy()
                            frame_tmp.rename(columns={'RXDATA3':'T1 VB'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T1 VB']],axis = 1)
                    if h4.get() == True:#IR
                        if h01.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 1].copy()
                            frame_tmp.rename(columns={'RXDATA4':'T1 IR'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T1 IR']],axis = 1)
                        if h06.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 6].copy()
                            frame_tmp.rename(columns={'RXDATA4':'T6 IR'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T6 IR']],axis = 1)
                    if h5.get() == True:#DVR
                        if h01.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 1].copy()
                            frame_tmp.rename(columns={'RXDATA5':'T1 DVR'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T1 DVR']],axis = 1)
                    if h7.get() == True:#DVF
                        if h03.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 3].copy()
                            frame_tmp.rename(columns={'RXDATA4':'T3 DVF'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T3 DVF']],axis = 1)
                else:
                    if h2.get() == True:#VF
                        if h01.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 1].copy()
                            frame_tmp.rename(columns={'RXDATA2':'T1 VF'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T1 VF']],axis = 1)

                        if h02.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 2].copy()
                            frame_tmp.rename(columns={'RXDATA2':'T2 VF'},inplace=True)
                            if np.mean(frame_tmp['T2 VF']) != 0:
                                frame_tmp.index = range(len(frame_tmp))
                                finaldata = pd.concat([finaldata,frame_tmp['T2 VF']],axis = 1)
                        
                        if h03.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 3].copy()
                            frame_tmp.rename(columns={'RXDATA2':'T3 VF1'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T3 VF1']],axis = 1)

                            frame_tmp = frame.loc[frame.TESTERNO == 3].copy()
                            frame_tmp.rename(columns={'RXDATA3':'T3 VF2'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T3 VF2']],axis = 1)
                        
                        if h05.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 5].copy()
                            frame_tmp.rename(columns={'RXDATA2':'T5 VF'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T5 VF']],axis = 1)

                        if h06.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 6].copy()
                            frame_tmp.rename(columns={'RXDATA2':'T6 VF'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T6 VF']],axis = 1)

                    if h3.get() == True:#VB
                        if h01.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 1].copy()
                            frame_tmp.rename(columns={'RXDATA3':'T1 VB'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T1 VB']],axis = 1)

                        if h02.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 2].copy()
                            frame_tmp.rename(columns={'RXDATA3':'T2 VB'},inplace=True)
                            if np.mean(frame_tmp['T2 VB']) != 0:
                                frame_tmp.index = range(len(frame_tmp))
                                finaldata = pd.concat([finaldata,frame_tmp['T2 VB']],axis = 1)

                        if h04.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 4].copy()
                            frame_tmp.rename(columns={'RXDATA1':'T4 VB'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T4 VB']],axis = 1)
                        
                        if h05.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 5].copy()
                            frame_tmp.rename(columns={'RXDATA3':'T5 VB'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T5 VB']],axis = 1)

                    if h4.get() == True:#IR
                        if h01.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 1].copy()
                            frame_tmp.rename(columns={'RXDATA4':'T1 IR'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T1 IR']],axis = 1)

                        if h02.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 2].copy()
                            frame_tmp.rename(columns={'RXDATA4':'T2 IR'},inplace=True)
                            if np.mean(frame_tmp['T2 IR']) != 0:
                                frame_tmp.index = range(len(frame_tmp))
                                finaldata = pd.concat([finaldata,frame_tmp['T2 IR']],axis = 1)

                        if h06.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 6].copy()
                            frame_tmp.rename(columns={'RXDATA4':'T6 IR'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T6 IR']],axis = 1)

                    if h7.get() == True:#DVF
                        if h03.get() == True:
                            frame_tmp = frame.loc[frame.TESTERNO == 6].copy()
                            frame_tmp.rename(columns={'RXDATA4':'T3 DVF'},inplace=True)
                            frame_tmp.index = range(len(frame_tmp))
                            finaldata = pd.concat([finaldata,frame_tmp['T3 DVF']],axis = 1)

                #print(finaldata)
                finaldata['TR'] = file[:-4]
                mm = finaldata.columns
                mm = pd.DataFrame(mm).T
                mm.columns = finaldata.columns
                finaldata = pd.concat([mm,finaldata], axis=0 ,ignore_index=True)
                finaldata.to_csv(fp + '/' + file[:-4] + '.csv',index=0,header=0)



                
                


    
        
    
    tellyou = "轉檔完成\n"+"檔案在桌面，不關閉可繼續轉檔\n"
    text.insert('insert',tellyou)
            

root=tk.Tk()
root.title("TR合併")
root.geometry("600x500")

label = tk.Label(root,
                text = '這個程式可以整理德微TR\n轉成JMP能夠分析的格式\n請選擇需要的項目後\n按下確認轉檔即可\n(請丟入資料夾)',
                bg = '#66B3FF',
                font = (12),
                width = 55, height = 6)
label.grid(row = 0, column = 0,columnspan=2)

h01 = tk.BooleanVar() # 设置选择框对象
cb01 = tk.Checkbutton(root,text='T1',variable=h01)
cb01.grid(row = 1, column = 0)

h02 = tk.BooleanVar() # 设置选择框对象
cb02 = tk.Checkbutton(root,text='T2',variable=h02)
cb02.grid(row = 2, column = 0)

h03 = tk.BooleanVar() # 设置选择框对象
cb03 = tk.Checkbutton(root,text='T3',variable=h03)
cb03.grid(row = 3, column = 0)

h04 = tk.BooleanVar() # 设置选择框对象
cb04 = tk.Checkbutton(root,text='T4',variable=h04)
cb04.grid(row = 4, column = 0)

h05 = tk.BooleanVar() # 设置选择框对象
cb05 = tk.Checkbutton(root,text='T5',variable=h05)
cb05.grid(row = 5, column = 0)

h06 = tk.BooleanVar() # 设置选择框对象
cb06 = tk.Checkbutton(root,text='T6',variable=h06)
cb06.grid(row = 6, column = 0)

h1 = tk.BooleanVar() # 设置选择框对象
cb1 = tk.Checkbutton(root,text='TRR',variable=h1)
cb1.grid(row = 1, column = 1)

h2 = tk.BooleanVar()
cb2 = tk.Checkbutton(root,text='VF',variable=h2)
cb2.grid(row = 2, column = 1)

h3 = tk.BooleanVar()
cb3 = tk.Checkbutton(root,text='VB',variable=h3)
cb3.grid(row = 3, column = 1)

h4 = tk.BooleanVar()
cb4 = tk.Checkbutton(root,text='IR',variable=h4)
cb4.grid(row = 4, column = 1)

h5 = tk.BooleanVar()
cb5 = tk.Checkbutton(root,text='DVR',variable=h5)
cb5.grid(row = 5, column = 1)

h6 = tk.BooleanVar()
cb6 = tk.Checkbutton(root,text='VC',variable=h6)
cb6.grid(row = 6, column = 1)

h7 = tk.BooleanVar()
cb7 = tk.Checkbutton(root,text='DVF',variable=h7)
cb7.grid(row = 7, column = 1)


ttk.Button(root, text="確認轉檔", command=open_file).grid(row = 8, column = 0,columnspan=2)
text = tk.Text(root,width=45,height=10)
text.grid(row = 9, column = 0,columnspan=2)
#ttk.Button(root, text="確認轉檔", command=save_file).pack()
root.mainloop()