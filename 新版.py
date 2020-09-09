from tkinter import filedialog
from tkinter import *
from tkinter import messagebox as mb
import tkinter as tk
import os
import math
import pandas as pd
import numpy as np
from tkinter import filedialog as fd

def DAT_one():
    openfilename=fd.askopenfilename(initialdir="/",title="Select file",\
                filetypes = (('DAT File', '*.DAT'),("all files","*.*")))
    p = 0
    for y in range(0,len(openfilename)):
        if openfilename[y] == "/":
            p = y
    f_name = openfilename[p+1:]
    Lot_No = openfilename[p+1:-4]
    print(f_name)
    Newdir = openfilename[:-4] + '.csv'
    file_test = open(openfilename,'r')
    file_test2 = open(Newdir,'w')
    for lines in file_test.readlines():
        strdata = ",".join(lines.split('\t'))
        file_test2.write(strdata)
    file_test.close()
    file_test2.close()

    data = pd.read_csv(Newdir,header = None)
    for x in range(0,data.shape[0]):
        if data[0][x][0:1] == '=':
            newdata = data[0][x]
            break

    z = []
    z_min = ['MIN',0,0]
    z_max = ['MAX',0,0]

    for x in range(536,548):
        num = 0
        num1 = 0
        z.append(data[0][x])
        if x == 536 or x == 537 or x == 539 or x == 540 or x == 547:
            for y in range(0,len(data[0][x])):
                if data[0][x][y:y+4] == 'Min=':
                    num = y + 4
                    while data[0][x][num] != 'V':
                        num = num + 1
                    #print(data[0][x][y+4:num])
                    z_min.append(data[0][x][y+4:num])
                if data[0][x][y:y+4] == 'Max=':
                    num1 = y + 4
                    while data[0][x][num1] != 'V':
                        num1 = num1 + 1
                    #print(data[0][x][y+4:num1])
                    z_max.append(data[0][x][y+4:num1])
            if num == 0:
                z_min.append(0)
            if num1 == 0:
                z_max.append(0)
        elif x == 542:
            for y in range(0,len(data[0][x])):
                if data[0][x][y:y+4] == 'Max=':
                    num1 = y + 4
                    while data[0][x][num1] != 'u':
                        num1 = num1 + 1
                    #print(data[0][x][y+4:num1])
                    z_max.append(data[0][x][y+4:num1])
            if num == 0:
                z_min.append(0)
            if num1 == 0:
                z_max.append(0)
        elif x == 546:
            for y in range(0,len(data[0][x])):
                if data[0][x][y:y+4] == 'Min=':
                    num = y + 4
                    while data[0][x][num] != 'n':
                        num = num + 1
                    #print(data[0][x][y+4:num])
                    z_min.append(data[0][x][y+4:num])
                if data[0][x][y:y+4] == 'Max=':
                    num1 = y + 4
                    while data[0][x][num1] != 'n':
                        num1 = num1 + 1
                    #print(data[0][x][y+4:num1])
                    z_max.append(data[0][x][y+4:num1])
            if num == 0:
                z_min.append(0)
            if num1 == 0:
                z_max.append(0)
        else:
            z_max.append(0)
            z_min.append(0)
    z.append('')
    z = pd.DataFrame(z)
    

    a = 0
    b = []
    for x in range(0,len(newdata)):
        if newdata[x] == '=' and x != 0:
            #print(newdata[a:x])
            #finaldata.insert(newdata[a:x])
            b.append(newdata[a:x])
            a = x
    if a == 0:
        b.append(newdata)
    a = pd.DataFrame(b)
    a = a[0].str.split('|',expand=True)
    #x = 15
    #y = a.shape[1]-1
    #while x <= y:
    #    del a[x]
    #    x = x + 1
    a.rename(columns={0:'No',1:'Bin',2:'unknow',3:'VF1',4:'VF2',5:'SG1',6:'VR1',7:'VR2',8:'dVR1',9:'IR1',10:'dIR1',11:'VR3',12:'dVR2',13:'TRR1',14:'SOP1'},inplace=True)
    if a.shape[1] > 15:
        a.rename(columns={17:'VB2',20:'IR2'},inplace=True)
        z_max.extend(z_max[4:])
        z_min.extend(z_min[4:])
    max_data = pd.DataFrame(z_max).T
    min_data = pd.DataFrame(z_min).T
    max_data.columns = a.columns
    min_data.columns = a.columns
    a = pd.concat([min_data,a], axis=0 ,ignore_index=True)
    a = pd.concat([max_data,a], axis=0 ,ignore_index=True)
    #print(a)
    for x in a.columns.values.tolist():
        a[x] = pd.to_numeric(a[x], errors='ignore')
    #a['VF1'] = pd.to_numeric(a['VF1'], errors='ignore')
    #a['VF2'] = pd.to_numeric(a['VF2'], errors='ignore')
    #a['SG1'] = pd.to_numeric(a['SG1'], errors='ignore')
    #a['VR1'] = pd.to_numeric(a['VR1'], errors='ignore')
    #a['VR2'] = pd.to_numeric(a['VR2'], errors='ignore')
    #a['dVR1'] = pd.to_numeric(a['dVR1'], errors='ignore')
    #a['IR1'] = pd.to_numeric(a['IR1'], errors='ignore')
    #a['dIR1'] = pd.to_numeric(a['dIR1'], errors='ignore')
    #a['VR3'] = pd.to_numeric(a['VR3'], errors='ignore')
    #a['dVR2'] = pd.to_numeric(a['dVR2'], errors='ignore')
    #a['TRR1'] = pd.to_numeric(a['TRR1'], errors='ignore')
    #a['SOP1'] = pd.to_numeric(a['SOP1'], errors='ignore')

    total = len(a) - 2
    GO = []
    NG = []
    for x in a.columns:
        if x == 'No':
            GO.append('GO')
            NG.append('NG')
        elif x != 'No' and x != 'Bin' and x != 'unknow' and x != 'SOP1':
            if sum(np.isnan(a[x]) != True) >= 3:
                #print(a[x][0])
                #print(a[x][1])
                ng = sum(a[x][2:] > a[x][0]) + sum(a[x][2:] < a[x][1])
                GO.append(sum(np.isnan(a[x]) != True)-2 - ng)
                NG.append(ng)
            else:
                GO.append(0)
                NG.append(0)
        else:
            GO.append(0)
            NG.append(0)
    YIELD = ['YIELD']
    for x in range(1,a.shape[1]):
        if GO[x] + NG[x] != 0:
            #print(GO[x]/total)
            YIELD.append(round(100*GO[x]/(GO[x] + NG[x]),2))
        else:
            YIELD.append(0)
    YIELD_data = pd.DataFrame(YIELD).T
    NG_data = pd.DataFrame(NG).T
    GO_data = pd.DataFrame(GO).T
    YIELD_data.columns = a.columns
    NG_data.columns = a.columns
    GO_data.columns = a.columns
    a = pd.concat([YIELD_data,a], axis=0 ,ignore_index=True)
    a = pd.concat([NG_data,a], axis=0 ,ignore_index=True)
    a = pd.concat([GO_data,a], axis=0 ,ignore_index=True)
    a.drop(index=[3, 4],inplace=True)
    #print(a)

    #aa = ['No','Bin','unknow','VF1','VF2','SG1','VR1','VR2','dVR1','IR1','dIR1','VR3','dVR2','TRR1','SOP1']
    aa = a.columns.values.tolist()
    aa_data = pd.DataFrame(aa).T
    aa_data.columns = a.columns
    a = pd.concat([aa_data,a], axis=0 ,ignore_index=True)

    for x in a.columns:
        if x != 'Bin':
            if a[x][1] == 0 and a[x][2] == 0 and a[x][3] == 0:
                del a[x]
                    
    w = []
    for x in range(0,a.shape[1]):
        w.append(x)
    a.columns = w
    #print(a)

    a_new = a[a[1] == 1]
    #print(a_new)

    bin1_name = ['Item']
    bin1_len = ['Count']
    bin1_min = ['Min']
    bin1_max = ['Max']
    bin1_mean = ['Avg']
    bin1_std = ['Sigma']
    bin1_mean_std = ['X+3Sigma']
    bin1_mean_stdown = ['X-3Sigma']

    if a_new.shape[0] == 0:
        bin1_name.append(a[x][0])
        bin1_len.append(len(a_new[x]))
        bin1_min.append(0)
        bin1_max.append(0)
        bin1_mean.append(0)
        bin1_std.append(0)
        bin1_mean_stdown.append(0)
        bin1_mean_std.append(0)
    else:
        for x in a_new.columns:
            if x == 0 or x == 1:
                0
            else:
                bin1_name.append(a[x][0])
                bin1_len.append(len(a_new[x]))
                bin1_min.append(round(min(a_new[x]),4))
                bin1_max.append(round(max(a_new[x]),4))
                bin1_mean.append(round(np.mean(a_new[x]),4))
                bin1_std.append(round(np.std(a_new[x],ddof=1),4))
                bin1_mean_stdown.append(round(np.mean(a_new[x]) - 3*np.std(a_new[x],ddof=1),4))
                bin1_mean_std.append(round(np.mean(a_new[x]) + 3*np.std(a_new[x],ddof=1),4))

    bin1_name.append('')
    bin1_len.append('')
    bin1_min.append('')
    bin1_max.append('')
    bin1_mean.append('')
    bin1_std.append('')
    bin1_mean_stdown.append('')
    bin1_mean_std.append('')

    k = [bin1_name,bin1_len,bin1_min,bin1_max,bin1_mean,bin1_std,bin1_mean_stdown,bin1_mean_std]
    k = pd.DataFrame(k)
    k = k.T
    #print(k)

    x = 7
    while a.shape[1] != k.shape[1]:
        if a.shape[1] > 8:
            k[x] = ''
            x = x+1
        if a.shape[1] < 8:
            a[x] = ''
            x = x-1
    x = 1
    while z.shape[1] != a.shape[1]:
        z[x] = ''
        x = x + 1

    finaldata = pd.concat([k,a], axis=0 ,ignore_index=True)
    finaldata = pd.concat([z,finaldata], axis=0 ,ignore_index=True)

    Lot_No = ['Lot_No:'+Lot_No]
    Lot_No = pd.DataFrame(Lot_No).T
    x = 1
    while Lot_No.shape[1] != a.shape[1]:
        Lot_No[x] = ''
        x = x + 1
    Lot_No.columns = a.columns
    finaldata = pd.concat([Lot_No,finaldata], axis=0 ,ignore_index=True)
    finaldata.to_csv(Newdir,index=0,header=0)
    tellyou = "執行完成，不關閉可繼續轉檔\n"
    text.insert('insert',tellyou)
    

def DAT_two():
    folder_selected = filedialog.askdirectory()
    allFileList = os.listdir(folder_selected)
    p = 0
    for y in range(0,len(folder_selected)):
        if folder_selected[y] == "/":
            p = y
    f_name = folder_selected[p+1:]
    #print(f_name)

    for file in allFileList:
        if os.path.isdir(folder_selected + '/' + file):
            smallfile_place = folder_selected + '/' + file
            if os.path.exists(folder_selected + '/' + file + '/' + file) == False:
                os.mkdir(folder_selected + '/' + file + '/' + file)
            final_place = folder_selected + '/' + file + '/' + file
            for file2 in os.listdir(smallfile_place):
                if file2[-4:] == '.dat' or file[-4:] == '.DAT':
                    print(file2)
                    openfilename = smallfile_place + '/' + file2
                    a = 0
                    for x in range(0,len(openfilename)):
                        if openfilename[x] == "/":
                            a = x
                    Lot_No = openfilename[a+1:-4]
                    Newdir = final_place + '/' + file2[:-4] + '.csv'

                    file_test = open(openfilename,'r')
                    file_test2 = open(Newdir,'w')
                    for lines in file_test.readlines():
                        strdata = ",".join(lines.split('\t'))
                        file_test2.write(strdata)
                    file_test.close()
                    file_test2.close()

                    data = pd.read_csv(Newdir,header = None)
                    for x in range(0,data.shape[0]):
                        if data[0][x][0:1] == '=':
                            newdata = data[0][x]
                            break

                    z = []
                    z_min = ['MIN',0,0]
                    z_max = ['MAX',0,0]

                    for x in range(536,548):
                        num = 0
                        num1 = 0
                        z.append(data[0][x])
                        if x == 536 or x == 537 or x == 539 or x == 540 or x == 547:
                            for y in range(0,len(data[0][x])):
                                if data[0][x][y:y+4] == 'Min=':
                                    num = y + 4
                                    while data[0][x][num] != 'V':
                                        num = num + 1
                                    #print(data[0][x][y+4:num])
                                    z_min.append(data[0][x][y+4:num])
                                if data[0][x][y:y+4] == 'Max=':
                                    num1 = y + 4
                                    while data[0][x][num1] != 'V':
                                        num1 = num1 + 1
                                    #print(data[0][x][y+4:num1])
                                    z_max.append(data[0][x][y+4:num1])
                            if num == 0:
                                z_min.append(0)
                            if num1 == 0:
                                z_max.append(0)
                        elif x == 542:
                            for y in range(0,len(data[0][x])):
                                if data[0][x][y:y+4] == 'Max=':
                                    num1 = y + 4
                                    while data[0][x][num1] != 'u':
                                        num1 = num1 + 1
                                    #print(data[0][x][y+4:num1])
                                    z_max.append(data[0][x][y+4:num1])
                            if num == 0:
                                z_min.append(0)
                            if num1 == 0:
                                z_max.append(0)
                        elif x == 546:
                            for y in range(0,len(data[0][x])):
                                if data[0][x][y:y+4] == 'Min=':
                                    num = y + 4
                                    while data[0][x][num] != 'n':
                                        num = num + 1
                                    #print(data[0][x][y+4:num])
                                    z_min.append(data[0][x][y+4:num])
                                if data[0][x][y:y+4] == 'Max=':
                                    num1 = y + 4
                                    while data[0][x][num1] != 'n':
                                        num1 = num1 + 1
                                    #print(data[0][x][y+4:num1])
                                    z_max.append(data[0][x][y+4:num1])
                            if num == 0:
                                z_min.append(0)
                            if num1 == 0:
                                z_max.append(0)
                        else:
                            z_max.append(0)
                            z_min.append(0)
                    z.append('')
                    z = pd.DataFrame(z)

                    a = 0
                    b = []
                    for x in range(0,len(newdata)):
                        if newdata[x] == '=' and x != 0:
                            #print(newdata[a:x])
                            #finaldata.insert(newdata[a:x])
                            b.append(newdata[a:x])
                            a = x
                    if a == 0:
                        b.append(newdata)
                    a = pd.DataFrame(b)
                    a = a[0].str.split('|',expand=True)
                    #x = 15
                    #y = a.shape[1]-1
                    #while x <= y:
                    #    del a[x]
                    #    x = x + 1

                    a.rename(columns={0:'No',1:'Bin',2:'unknow',3:'VF1',4:'VF2',5:'SG1',6:'VR1',7:'VR2',8:'dVR1',9:'IR1',10:'dIR1',11:'VR3',12:'dVR2',13:'TRR1',14:'SOP1'},inplace=True)
                    if a.shape[1] > 15:
                        a.rename(columns={17:'VB2',20:'IR2'},inplace=True)
                        z_max.extend(z_max[4:])
                        z_min.extend(z_min[4:])
                    max_data = pd.DataFrame(z_max).T
                    min_data = pd.DataFrame(z_min).T
                    max_data.columns = a.columns
                    min_data.columns = a.columns
                    a = pd.concat([min_data,a], axis=0 ,ignore_index=True)
                    a = pd.concat([max_data,a], axis=0 ,ignore_index=True)
                    for x in a.columns.values.tolist():
                        a[x] = pd.to_numeric(a[x], errors='ignore')
                    #a['VF1'] = pd.to_numeric(a['VF1'], errors='ignore')
                    #a['VF2'] = pd.to_numeric(a['VF2'], errors='ignore')
                    #a['SG1'] = pd.to_numeric(a['SG1'], errors='ignore')
                    #a['VR1'] = pd.to_numeric(a['VR1'], errors='ignore')
                    #a['VR2'] = pd.to_numeric(a['VR2'], errors='ignore')
                    #a['dVR1'] = pd.to_numeric(a['dVR1'], errors='ignore')
                    #a['IR1'] = pd.to_numeric(a['IR1'], errors='ignore')
                    #a['dIR1'] = pd.to_numeric(a['dIR1'], errors='ignore')
                    #a['VR3'] = pd.to_numeric(a['VR3'], errors='ignore')
                    #a['dVR2'] = pd.to_numeric(a['dVR2'], errors='ignore')
                    #a['TRR1'] = pd.to_numeric(a['TRR1'], errors='ignore')
                    #a['SOP1'] = pd.to_numeric(a['SOP1'], errors='ignore')

                    total = len(a) - 2
                    GO = []
                    NG = []
                    for x in a.columns:
                        if x == 'No':
                            GO.append('GO')
                            NG.append('NG')
                        elif x != 'No' and x != 'Bin' and x != 'unknow' and x != 'SOP1':
                            if sum(np.isnan(a[x]) != True) >= 3:
                                #print(a[x][0])
                                #print(a[x][1])
                                ng = sum(a[x][2:] > a[x][0]) + sum(a[x][2:] < a[x][1])
                                GO.append(sum(np.isnan(a[x]) != True)-2 - ng)
                                NG.append(ng)
                            else:
                                GO.append(0)
                                NG.append(0)
                        else:
                            GO.append(0)
                            NG.append(0)
                    YIELD = ['YIELD']
                    for x in range(1,a.shape[1]):
                        if GO[x] + NG[x] != 0:
                            #print(GO[x]/total)
                            YIELD.append(round(100*GO[x]/(GO[x] + NG[x]),2))
                        else:
                            YIELD.append(0)
                    YIELD_data = pd.DataFrame(YIELD).T
                    NG_data = pd.DataFrame(NG).T
                    GO_data = pd.DataFrame(GO).T
                    YIELD_data.columns = a.columns
                    NG_data.columns = a.columns
                    GO_data.columns = a.columns
                    a = pd.concat([YIELD_data,a], axis=0 ,ignore_index=True)
                    a = pd.concat([NG_data,a], axis=0 ,ignore_index=True)
                    a = pd.concat([GO_data,a], axis=0 ,ignore_index=True)
                    a.drop(index=[3, 4],inplace=True)

                    #aa = ['No','Bin','unknow','VF1','VF2','SG1','VR1','VR2','dVR1','IR1','dIR1','VR3','dVR2','TRR1','SOP1']
                    aa = a.columns.values.tolist()
                    aa_data = pd.DataFrame(aa).T
                    aa_data.columns = a.columns
                    a = pd.concat([aa_data,a], axis=0 ,ignore_index=True)

                    for x in a.columns:
                        if x != 'Bin':
                            if a[x][1] == 0 and a[x][2] == 0 and a[x][3] == 0:
                                del a[x]
                    
                    w = []
                    for x in range(0,a.shape[1]):
                        w.append(x)
                    a.columns = w

                    a_new = a[a[1] == 1]

                    bin1_name = ['Item']
                    bin1_len = ['Count']
                    bin1_min = ['Min']
                    bin1_max = ['Max']
                    bin1_mean = ['Avg']
                    bin1_std = ['Sigma']
                    bin1_mean_std = ['X+3Sigma']
                    bin1_mean_stdown = ['X-3Sigma']

                    if a_new.shape[0] == 0:
                        bin1_name.append(a[x][0])
                        bin1_len.append(len(a_new[x]))
                        bin1_min.append(0)
                        bin1_max.append(0)
                        bin1_mean.append(0)
                        bin1_std.append(0)
                        bin1_mean_stdown.append(0)
                        bin1_mean_std.append(0)
                    else:
                        for x in a_new.columns:
                            if x == 0 or x == 1:
                                0
                            else:
                                bin1_name.append(a[x][0])
                                bin1_len.append(len(a_new[x]))
                                bin1_min.append(round(min(a_new[x]),4))
                                bin1_max.append(round(max(a_new[x]),4))
                                bin1_mean.append(round(np.mean(a_new[x]),4))
                                bin1_std.append(round(np.std(a_new[x],ddof=1),4))
                                bin1_mean_stdown.append(round(np.mean(a_new[x]) - 3*np.std(a_new[x],ddof=1),4))
                                bin1_mean_std.append(round(np.mean(a_new[x]) + 3*np.std(a_new[x],ddof=1),4))
                    bin1_name.append('')
                    bin1_len.append('')
                    bin1_min.append('')
                    bin1_max.append('')
                    bin1_mean.append('')
                    bin1_std.append('')
                    bin1_mean_stdown.append('')
                    bin1_mean_std.append('')

                    k = [bin1_name,bin1_len,bin1_min,bin1_max,bin1_mean,bin1_std,bin1_mean_stdown,bin1_mean_std]
                    k = pd.DataFrame(k)
                    k = k.T

                    x = 7
                    while a.shape[1] != k.shape[1]:
                        if a.shape[1] > 8:
                            k[x] = ''
                            x = x+1
                        if a.shape[1] < 8:
                            a[x] = ''
                            x = x-1
                    x = 1
                    while z.shape[1] != a.shape[1]:
                        z[x] = ''
                        x = x + 1

                    finaldata = pd.concat([k,a], axis=0 ,ignore_index=True)
                    finaldata = pd.concat([z,finaldata], axis=0 ,ignore_index=True)

                    Lot_No = ['Lot_No:'+Lot_No]
                    Lot_No = pd.DataFrame(Lot_No).T
                    x = 1
                    while Lot_No.shape[1] != a.shape[1]:
                        Lot_No[x] = ''
                        x = x + 1
                    Lot_No.columns = a.columns
                    finaldata = pd.concat([Lot_No,finaldata], axis=0 ,ignore_index=True)
                    finaldata.to_csv(Newdir,index=0,header=0)

        if file[-4:] == '.dat' or file[-4:] == '.DAT':
            if os.path.exists(folder_selected + '/' + f_name) == False:
                os.mkdir(folder_selected + '/' + f_name)
            print(file)
            openfilename = folder_selected + '/' + file
            a = 0
            for x in range(0,len(openfilename)):
                if openfilename[x] == "/":
                    a = x
            Lot_No = openfilename[a+1:-4]
            Newdir = folder_selected + '/' + f_name + '/' + file[:-4] + '.csv'

            file_test = open(openfilename,'r')
            file_test2 = open(Newdir,'w')
            for lines in file_test.readlines():
                strdata = ",".join(lines.split('\t'))
                file_test2.write(strdata)
            file_test.close()
            file_test2.close()

            data = pd.read_csv(Newdir,header = None)
            for x in range(0,data.shape[0]):
                if data[0][x][0:1] == '=':
                    newdata = data[0][x]
                    break
            #print(newdata)

            z = []
            z_min = ['MIN',0,0]
            z_max = ['MAX',0,0]

            for x in range(536,548):
                num = 0
                num1 = 0
                z.append(data[0][x])
                if x == 536 or x == 537 or x == 539 or x == 540 or x == 547:
                    for y in range(0,len(data[0][x])):
                        if data[0][x][y:y+4] == 'Min=':
                            num = y + 4
                            while data[0][x][num] != 'V':
                                num = num + 1
                            #print(data[0][x][y+4:num])
                            z_min.append(data[0][x][y+4:num])
                        if data[0][x][y:y+4] == 'Max=':
                            num1 = y + 4
                            while data[0][x][num1] != 'V':
                                num1 = num1 + 1
                            #print(data[0][x][y+4:num1])
                            z_max.append(data[0][x][y+4:num1])
                    if num == 0:
                        z_min.append(0)
                    if num1 == 0:
                        z_max.append(0)
                elif x == 542:
                    for y in range(0,len(data[0][x])):
                        if data[0][x][y:y+4] == 'Max=':
                            num1 = y + 4
                            while data[0][x][num1] != 'u':
                                num1 = num1 + 1
                            #print(data[0][x][y+4:num1])
                            z_max.append(data[0][x][y+4:num1])
                    if num == 0:
                        z_min.append(0)
                    if num1 == 0:
                        z_max.append(0)
                elif x == 546:
                    for y in range(0,len(data[0][x])):
                        if data[0][x][y:y+4] == 'Min=':
                            num = y + 4
                            while data[0][x][num] != 'n':
                                num = num + 1
                            #print(data[0][x][y+4:num])
                            z_min.append(data[0][x][y+4:num])
                        if data[0][x][y:y+4] == 'Max=':
                            num1 = y + 4
                            while data[0][x][num1] != 'n':
                                num1 = num1 + 1
                            #print(data[0][x][y+4:num1])
                            z_max.append(data[0][x][y+4:num1])
                    if num == 0:
                        z_min.append(0)
                    if num1 == 0:
                        z_max.append(0)
                else:
                    z_max.append(0)
                    z_min.append(0)
            z.append('')
            z = pd.DataFrame(z)

            a = 0
            b = []
            for x in range(0,len(newdata)):
                if newdata[x] == '=' and x != 0:
                    #print(newdata[a:x])
                    #finaldata.insert(newdata[a:x])
                    b.append(newdata[a:x])
                    a = x
            if a == 0:
                b.append(newdata)
            a = pd.DataFrame(b)
            a = a[0].str.split('|',expand=True)
            #x = 15
            #y = a.shape[1]-1
            #while x <= y:
            #    del a[x]
            #    x = x + 1
            #print(a)
            a.rename(columns={0:'No',1:'Bin',2:'unknow',3:'VF1',4:'VF2',5:'SG1',6:'VR1',7:'VR2',8:'dVR1',9:'IR1',10:'dIR1',11:'VR3',12:'dVR2',13:'TRR1',14:'SOP1',17:'VB2',20:'IR2'},inplace=True)
            if a.shape[1] > 15:
                a.rename(columns={17:'VB2',20:'IR2'},inplace=True)
                z_max.extend(z_max[4:])
                z_min.extend(z_min[4:])
            max_data = pd.DataFrame(z_max).T
            min_data = pd.DataFrame(z_min).T
            #print(a)
            #print(max_data)
            max_data.columns = a.columns
            min_data.columns = a.columns
            a = pd.concat([min_data,a], axis=0 ,ignore_index=True)
            a = pd.concat([max_data,a], axis=0 ,ignore_index=True)
            for x in a.columns.values.tolist():
                a[x] = pd.to_numeric(a[x], errors='ignore')
            #a['VF1'] = pd.to_numeric(a['VF1'], errors='ignore')
            #a['VF2'] = pd.to_numeric(a['VF2'], errors='ignore')
            #a['SG1'] = pd.to_numeric(a['SG1'], errors='ignore')
            #a['VR1'] = pd.to_numeric(a['VR1'], errors='ignore')
            #a['VR2'] = pd.to_numeric(a['VR2'], errors='ignore')
            #a['dVR1'] = pd.to_numeric(a['dVR1'], errors='ignore')
            #a['IR1'] = pd.to_numeric(a['IR1'], errors='ignore')
            #a['dIR1'] = pd.to_numeric(a['dIR1'], errors='ignore')
            #a['VR3'] = pd.to_numeric(a['VR3'], errors='ignore')
            #a['dVR2'] = pd.to_numeric(a['dVR2'], errors='ignore')
            #a['TRR1'] = pd.to_numeric(a['TRR1'], errors='ignore')
            #a['SOP1'] = pd.to_numeric(a['SOP1'], errors='ignore')
            #a[15] = pd.to_numeric(a[15], errors='ignore')
            #a[16] = pd.to_numeric(a[16], errors='ignore')
            #a['VB2'] = pd.to_numeric(a['VB2'], errors='ignore')
            #a[18] = pd.to_numeric(a[18], errors='ignore')
            #a[19] = pd.to_numeric(a[19], errors='ignore')
            #a['IR2'] = pd.to_numeric(a['IR2'], errors='ignore')
            #a[21] = pd.to_numeric(a[21], errors='ignore')
            #a[22] = pd.to_numeric(a[22], errors='ignore')
            #a[23] = pd.to_numeric(a[23], errors='ignore')
            #a[24] = pd.to_numeric(a[24], errors='ignore')
            #a[25] = pd.to_numeric(a[25], errors='ignore')


            #print(a)
            total = len(a) - 2
            GO = []
            NG = []
            for x in a.columns:
                if x == 'No':
                    GO.append('GO')
                    NG.append('NG')
                elif x != 'No' and x != 'Bin' and x != 'unknow' and x != 'SOP1':
                    if sum(np.isnan(a[x]) != True) >= 3:
                        #print(a[x][0])
                        #print(a[x][1])
                        ng = sum(a[x][2:] > a[x][0]) + sum(a[x][2:] < a[x][1])
                        GO.append(sum(np.isnan(a[x]) != True)-2 - ng)
                        NG.append(ng)
                    else:
                        GO.append(0)
                        NG.append(0)
                else:
                    GO.append(0)
                    NG.append(0)
            YIELD = ['YIELD']
            for x in range(1,a.shape[1]):
                if GO[x] + NG[x] != 0:
                    #print(GO[x]/total)
                    YIELD.append(round(100*GO[x]/(GO[x] + NG[x]),2))
                else:
                    YIELD.append(0)
            YIELD_data = pd.DataFrame(YIELD).T
            NG_data = pd.DataFrame(NG).T
            GO_data = pd.DataFrame(GO).T
            YIELD_data.columns = a.columns
            NG_data.columns = a.columns
            GO_data.columns = a.columns
            a = pd.concat([YIELD_data,a], axis=0 ,ignore_index=True)
            a = pd.concat([NG_data,a], axis=0 ,ignore_index=True)
            a = pd.concat([GO_data,a], axis=0 ,ignore_index=True)
            a.drop(index=[3, 4],inplace=True)

            aa = ['No','Bin','unknow','VF1','VF2','SG1','VR1','VR2','dVR1','IR1','dIR1','VR3','dVR2','TRR1','SOP1']
            aa = a.columns.values.tolist()
            aa_data = pd.DataFrame(aa).T
            aa_data.columns = a.columns
            a = pd.concat([aa_data,a], axis=0 ,ignore_index=True)


            for x in a.columns:
                if x != 'Bin':
                    if a[x][1] == 0 and a[x][2] == 0 and a[x][3] == 0:
                        del a[x]

            w = []
            for x in range(0,a.shape[1]):
                w.append(x)
            a.columns = w
            #print(a)

            a_new = a[a[1] == 1]

            bin1_name = ['Item']
            bin1_len = ['Count']
            bin1_min = ['Min']
            bin1_max = ['Max']
            bin1_mean = ['Avg']
            bin1_std = ['Sigma']
            bin1_mean_std = ['X+3Sigma']
            bin1_mean_stdown = ['X-3Sigma']

            if a_new.shape[0] == 0:
                bin1_name.append(a[x][0])
                bin1_len.append(len(a_new[x]))
                bin1_min.append(0)
                bin1_max.append(0)
                bin1_mean.append(0)
                bin1_std.append(0)
                bin1_mean_stdown.append(0)
                bin1_mean_std.append(0)
            else:
                for x in a_new.columns:
                    if x == 0 or x == 1:
                        0
                    else:
                        bin1_name.append(a[x][0])
                        bin1_len.append(len(a_new[x]))
                        bin1_min.append(round(min(a_new[x]),4))
                        bin1_max.append(round(max(a_new[x]),4))
                        bin1_mean.append(round(np.mean(a_new[x]),4))
                        bin1_std.append(round(np.std(a_new[x],ddof=1),4))
                        bin1_mean_stdown.append(round(np.mean(a_new[x]) - 3*np.std(a_new[x],ddof=1),4))
                        bin1_mean_std.append(round(np.mean(a_new[x]) + 3*np.std(a_new[x],ddof=1),4))
            bin1_name.append('')
            bin1_len.append('')
            bin1_min.append('')
            bin1_max.append('')
            bin1_mean.append('')
            bin1_std.append('')
            bin1_mean_stdown.append('')
            bin1_mean_std.append('')

            k = [bin1_name,bin1_len,bin1_min,bin1_max,bin1_mean,bin1_std,bin1_mean_stdown,bin1_mean_std]
            k = pd.DataFrame(k)
            k = k.T

            x = 7
            while a.shape[1] != k.shape[1]:
                if a.shape[1] > 8:
                    k[x] = ''
                    x = x+1
                if a.shape[1] < 8:
                    a[x] = ''
                    x = x-1
            x = 1
            while z.shape[1] != a.shape[1]:
                z[x] = ''
                x = x + 1

            finaldata = pd.concat([k,a], axis=0 ,ignore_index=True)
            finaldata = pd.concat([z,finaldata], axis=0 ,ignore_index=True)

            Lot_No = ['Lot_No:'+Lot_No]
            Lot_No = pd.DataFrame(Lot_No).T
            x = 1
            while Lot_No.shape[1] != a.shape[1]:
                Lot_No[x] = ''
                x = x + 1
            Lot_No.columns = a.columns
            finaldata = pd.concat([Lot_No,finaldata], axis=0 ,ignore_index=True)

            finaldata.to_csv(Newdir,index=0,header=0)
    tellyou = "執行完成，不關閉可繼續轉檔\n"
    text.insert('insert',tellyou)


root=tk.Tk()
root.title("整理DAT檔案丟入JMP")
root.geometry("600x400")

label = tk.Label(root,
                text = '這個程式可以把DAT檔轉成可讀的格式\n有測試條件跟讀值之類的\n按下按鈕即可轉檔\n(可放入檔案or資料夾，資料夾裡面為檔案or資料夾皆可)',
                bg = '#EEBB00',
                font = (13),
                width = 55, height = 6)
label.grid(row = 0, column = 0,columnspan=1)




#tk.Button(root, text="我要轉德微", command=ERIS_DAT).grid(row = 6, column = 0)
tk.Button(root, text="我要轉一個檔案", command=DAT_one).grid(row = 1, column = 0)
tk.Button(root, text="我要轉資料夾", command=DAT_two).grid(row = 2, column = 0)
text = tk.Text(root,width=30,height=10)
text.grid(row = 8, column = 0,columnspan=1)
#ttk.Button(root, text="確認轉檔", command=save_file).pack()
root.mainloop()