from tkinter import filedialog as fd
from tkinter import messagebox as mb
import tkinter as tk
import zipfile
import os
import math
from dbfread import DBF
import pandas as pd
import numpy as np
import winreg

def rename_duplicate(list):
    new_list=[v + str(list[:i].count(v) + 1) if list.count(v) > 1 else v for i, v in enumerate(list)]
    return new_list

def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key, "Desktop")[0]

def BY_BIG():
    folder_selected = fd.askdirectory()
    allFileList = os.listdir(folder_selected)
    if os.path.exists(get_desktop() + '/' + '可以丟到JMP裡面的程式') == False:
        os.mkdir(get_desktop() + '/' + '可以丟到JMP裡面的程式')

    for file in allFileList:
        if os.path.isdir(folder_selected + '/' + file):
            smallfile_place = folder_selected + '/' + file
            print(file)
            #dd = {"VB1":h1.get(),
            #        "VB2":h2.get(),
            #        "IR":h3.get(),
            #        "TRR":h4.get(),
            #        "TRR OFFSET":h5.get()}

            #last = ["NO","name","BIN"]
            #for x in dd:
            #    if dd[x] == True:
            #        last.append(x)
            #final_data = pd.DataFrame(columns = last)
            final_data = pd.DataFrame()
            
            for file2 in os.listdir(smallfile_place):
                if file2[-4:] == '.zip':
                    print(file2[:-4])
                    z = zipfile.ZipFile(smallfile_place + '/' + file2, "r")
                    #print(z.namelist())
                    z.extract(z.namelist()[0])
                    data_cfg = open(z.namelist()[0], mode='r')
                    data_cfg = pd.DataFrame(data_cfg)
                    os.remove(z.namelist()[0])
                    if data_cfg[0][0] == 'Ver             2.0             \n':
                        x = 11
                        y = ["NO","BIN"]
                        while data_cfg[0][x][:2] != '\n':
                            y.append(data_cfg[0][x][:2])
                            x = x + 1
                        for x in range(0,len(y)):
                            if y[x] == '05':
                                y[x] = 'VB'
                            elif y[x] == '21':
                                y[x] = 'Delta'
                            elif y[x] == '06':
                                y[x] = 'IR'
                            elif y[x] == '02':
                                y[x] = 'TRR'
                            elif y[x] == '23':
                                y[x] = 'TRR OFFSET'
                            elif y[x] == '01':
                                y[x] = 'KELVIN'
                            elif y[x] == '22':
                                y[x] = 'ABS'
                            elif y[x] == '04':
                                y[x] = 'VF(A)'
                        y = rename_duplicate(y)

                        z.extract(z.namelist()[2])
                        dbf = DBF(z.namelist()[2])
                        frame = pd.DataFrame(iter(dbf))
                        os.remove(z.namelist()[2])
                        if frame.columns.values.tolist() != []:
                            data = frame.drop(columns=['X          ', 'Y          '])
                            data.columns = y
                            #print(data)

                            dd = {"VB":h1.get(),
                                    "VB1":h1.get(),
                                    "VB2":h2.get(),
                                    "IR":h3.get(),
                                    "IR1":h3.get(),
                                    "IR2":h3.get(),
                                    "TRR":h4.get(),
                                    "TRR OFFSET":h5.get(),
                                    "TRR OFFSET1":h5.get(),
                                    "TRR OFFSET2":h5.get()}

                            last = ["NO","NAME","BIN"]
                            for x in dd:
                                if dd[x] == True:
                                    if np.any(x in data) == True:
                                        last.append(x)
                            for y in range(0,len(file2)):
                                if file2[y] == '.':
                                    data['NAME'] = file2[:y]
                                    break
                            data = data[last]
                            #print(data)
                            j = data.shape[1]
                            
                            
                                                            
                            #print(data)
                            final_data = pd.concat([final_data,data], axis=0 ,ignore_index=True)
                    else:
                        #print(z.namelist())
                        z.extract(z.namelist()[1])
                        dbf = DBF(z.namelist()[1])
                        frame = pd.DataFrame(iter(dbf))
                        os.remove(z.namelist()[1])

                        if frame.columns.values.tolist() != []:
                            data = frame.drop(columns=['X          ', 'Y          '])
                            for y in range(0,len(file2)):
                                if file2[y] == '.':
                                    data['NAME'] = file2[:y]
                                    break
                            #print(data.columns)
                            data = data[['RECORDNO   ','NAME','BINNO      ','RXDATA6    ','RXDATA11   ']]
                            data.rename(columns={'RECORDNO   ':'NO','BINNO      ':'BIN','RXDATA6    ':'VB','RXDATA11   ':'IR'},inplace=True)
                            final_data = pd.concat([final_data,data], axis=0 ,ignore_index=True)

                #os.remove(file_name)

                if file2[-4:] == '.dat' or file2[-4:] == '.DAT':
                    print(file2)
                    tmpdir = smallfile_place + '/tmp.csv'
                    openfilename = smallfile_place + '/' + file2

                    file_test = open(openfilename,'r')
                    file_test2 = open(tmpdir,'w')
                    for lines in file_test.readlines():
                        strdata = ",".join(lines.split('\t'))
                        file_test2.write(strdata)
                    file_test.close()
                    file_test2.close()

                    data = pd.read_csv(tmpdir,header = None)
                    for x in range(0,data.shape[0]):
                        if data[0][x][0:1] == '=':
                            newdata = data[0][x]
                            break
                    #print(newdata)

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
                    j = a.shape[1]
                    #print(a)
                    #x = 15
                    #y = a.shape[1]-1
                    #while x <= y:
                    #    del a[x]
                    #    x = x + 1

                    a.rename(columns={0:'NO',1:'BIN',2:'unknow',3:'VF1',4:'VF2',5:'SG1',6:'VR1',7:'VR2',8:'dVR1',9:'IR1',10:'dIR1',11:'VR3',12:'dVR2',13:'TRR1',14:'SOP1',17:'VB2',20:'IR2'},inplace=True)
                    if j > 15:
                        a.rename(columns={17:'VB2',20:'IR2'},inplace=True)
                    a['VF1'] = pd.to_numeric(a['VF1'], errors='ignore')
                    a['VF2'] = pd.to_numeric(a['VF2'], errors='ignore')
                    a['SG1'] = pd.to_numeric(a['SG1'], errors='ignore')
                    a['VR1'] = pd.to_numeric(a['VR1'], errors='ignore')
                    a['VR2'] = pd.to_numeric(a['VR2'], errors='ignore')
                    a['dVR1'] = pd.to_numeric(a['dVR1'], errors='ignore')
                    a['IR1'] = pd.to_numeric(a['IR1'], errors='ignore')
                    a['dIR1'] = pd.to_numeric(a['dIR1'], errors='ignore')
                    a['VR3'] = pd.to_numeric(a['VR3'], errors='ignore')
                    a['dVR2'] = pd.to_numeric(a['dVR2'], errors='ignore')
                    a['TRR1'] = pd.to_numeric(a['TRR1'], errors='ignore')
                    a['SOP1'] = pd.to_numeric(a['SOP1'], errors='ignore')
                    if j > 15:
                        a['VB2'] = pd.to_numeric(a['VB2'], errors='ignore')
                        a['IR2'] = pd.to_numeric(a['IR2'], errors='ignore')
                    a.rename(columns={'VR1':'VB','VR2':'VB1','TRR1':'TRR OFFSET'},inplace=True)
                    #print(a)
                    #aa = ['No','Bin','unknow','VF1','VF2','SG1','VR1','VR2','dVR1','IR1','dIR1','VR3','dVR2','TRR1','SOP1']
                    #aa_data = pd.DataFrame(aa).T
                    #aa_data.columns = a.columns
                    #a = pd.concat([aa_data,a], axis=0 ,ignore_index=True)
                    
                    dd = {"VB":h1.get(),
                            "VB1":h2.get(),
                            "VB2":h1.get(),
                            "IR":h3.get(),
                            "IR1":h3.get(),
                            "IR2":h3.get(),
                            "TRR":h4.get(),
                            "TRR OFFSET":h5.get(),
                            "TRR OFFSET1":h5.get(),
                            "TRR OFFSET2":h5.get()}

                    last = ["NO","NAME","BIN"]
                    for x in dd:
                        if dd[x] == True:
                            if np.any(x in a) == True:
                                last.append(x)
                    for y in range(0,len(file2)):
                        if file2[y] == ' ':
                            #print(file2[0:y])
                            a["NAME"] = file2[0:y]
                            break
                        if file2[y] == '.':
                            #print(file2[0:y])
                            a["NAME"] = file2[0:y]
                            break
                    a = a[last]
                    #print(a)

                    
                    #print(a)
                    final_data = pd.concat([final_data,a], axis=0 ,ignore_index=True)
                    os.remove(tmpdir)
            t = list(range(1,final_data.shape[0]+1))
            final_data['NO'] = t

            #if j <= 15:
            #    dd = {"VB1":h1.get(),
            #            "VR2":h2.get(),
            #            "IR":h3.get(),
            #            "TRR":h4.get(),
            #            "TRR OFFSET":h5.get()}
            #else:
            #    dd = {"VB1":h1.get(),
            #            "VR2":h2.get(),
            #            "IR":h3.get(),
            #            "TRR":h4.get(),
            #            "TRR OFFSET":h5.get(),
            #            "VB2":h1.get(),
            #            "IR2":h3.get()}

            #last = ["NO","name","BIN"]
            #for x in dd:
            #    if dd[x] == True:
            #        last.append(x)
            last = final_data.columns
            last_data = pd.DataFrame(last).T
            last_data.columns = final_data.columns
            final_data = pd.concat([last_data,final_data], axis=0 ,ignore_index=True)

            final_data.to_csv(get_desktop() + '/' + '可以丟到JMP裡面的程式' + '/' + file + '.csv',index=0,header=0)
    
                        
    tellyou = "執行完成，檔案在桌面喔\n不關閉可繼續轉檔\n"
    text.insert('insert',tellyou) 
                                
    

root=tk.Tk()
root.title("整理DAT檔案丟入JMP")
root.geometry("600x500")

label = tk.Label(root,
                text = '這個程式可以整理亞昕DAT檔orZIP檔\n轉成JMP能夠分析的格式\n請選擇需要的項目後\n按下我要轉檔即可\n(請丟入資料夾)',
                bg = '#EEBB00',
                font = (12),
                width = 55, height = 6)
label.grid(row = 0, column = 0,columnspan=1)

h1 = tk.BooleanVar() # 设置选择框对象
cb1 = tk.Checkbutton(root,text='VB1(VR)',variable=h1)
cb1.grid(row = 1, column = 0)

h2 = tk.BooleanVar()
cb2 = tk.Checkbutton(root,text='VB2',variable=h2)
cb2.grid(row = 2, column = 0)

h3 = tk.BooleanVar()
cb3 = tk.Checkbutton(root,text='IR',variable=h3)
cb3.grid(row = 3, column = 0)

h4 = tk.BooleanVar()
cb4 = tk.Checkbutton(root,text='TRR',variable=h4)
cb4.grid(row = 4, column = 0)

h5 = tk.BooleanVar()
cb5 = tk.Checkbutton(root,text='TRR OFFSET(TRR1)',variable=h5)
cb5.grid(row = 5, column = 0)




tk.Button(root, text="我要轉檔", command=BY_BIG).grid(row = 6, column = 0)
text = tk.Text(root,width=40,height=8)
text.grid(row = 7, column = 0,columnspan=1)
root.mainloop()