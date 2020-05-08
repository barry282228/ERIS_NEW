import sys
import getopt
import pymysql
import numpy as np
import pandas as pd
#import matplotlib.pyplot as plt
import datetime



args = sys.argv[1:]
def usage():
    print("Usage:{0} [-h name | --host=name] [-P # | --port=#] [-u name | --user=name] [-p password | --password=password] [-m mmcode | --mmcode=name] [-? | --help]".format(sys.argv[0]))

if (sys.argv.__len__()) == 1:
    usage()

try:
    opts, args = getopt.getopt(args, 'h:P:u:p:m:?', ["host=", "port=", "user=", "password=", "mmcode=", "help"])

except getopt.GetoptError:
    usage()
    sys.exit(1)

for opt, arg in opts:
    if opt in ('-?','--help'):
        usage()
    elif opt in ('-h', '--host'):
        host = arg
    elif opt in ('-P', '--port'):
        port = arg
    elif opt in ('-u', '--user'):
        user = arg
    elif opt in ('-p', '--password'):
        passwd = arg
    elif opt in ('-m', '--mmcode'):
        mmcode = arg

host = "{1}".format('host',eval('host'))
port = "{1}".format('port',eval('port'))
user = "{1}".format('user',eval('user'))
passwd = "{1}".format('passwd',eval('passwd'))
mmcode = "{1}".format('mmcode',eval('mmcode'))


db = pymysql.connect(host=host, user=user, passwd=passwd, db='erp_db')
cursor = db.cursor()

sql = """SELECT
  A.id,
  A.mm_exam_no,
  A.mm_date,
  A.mm_type,
  A.mm_userid,
  A.audit_user,
  B.mm_main_exam,
  B.mm_exam_no,
  C.mm_exam_no,
  C.mm_no_gp,
  C.mm_name,
  B.mm_num,
  D.mm_catego
  
FROM mm_warehouse_main A
LEFT JOIN mm_warehouse_desc B on B.mm_main_exam = A.mm_exam_no
LEFT JOIN mm_material_base C on C.mm_exam_no = B.mm_exam_no
LEFT JOIN mm_material_base_belong_category D on D.mm_exam_no = B.mm_exam_no
WHERE 1=1
and C.mm_no_gp=%s
and A.mm_type=2
and D.token=1
and DATE_SUB(CURDATE(), INTERVAL 550 DAY) <= date(A.mm_date)"""
cursor.execute(sql,mmcode)
data = cursor.fetchall()
db.close()

newdata = pd.DataFrame(data)
newdata = newdata.sort_values(by=2)
trydata = newdata[[2,11,12]]
trydata = trydata.reset_index(drop=True)
trydata = trydata.rename(columns={2:'time',11:'number',12:'mmlevel'})
trydata['days'] = trydata['time'].diff()
offnumber = ((len(trydata)*0.025)//1)+1
if offnumber == 1:
    trydata = trydata.drop([trydata['days'].idxmax()])
    trydata = trydata.drop([trydata['days'].idxmin()])
    trydata = trydata.reset_index(drop=True)
else:
    for x in range(1,offnumber):
        trydata = trydata.drop([trydata['days'].idxmax()])
        trydata = trydata.drop([trydata['days'].idxmin()])
        trydata = trydata.reset_index(drop=True)
take_time = trydata.at[len(trydata)-1,'time']

today = take_time.date() + datetime.timedelta(days=1)
rangeday = datetime.timedelta(days=30)
limitday = today - rangeday
a = []

while today >= min(trydata["time"]):
    data= trydata[(trydata["time"] <= pd.Timestamp(today)) & (trydata["time"] >= pd.Timestamp(limitday))]
    a.append(sum(data["number"]))
    today = today - rangeday
    limitday = today - rangeday


db = pymysql.connect(host=host, user=user, passwd=passwd, db='erp_db')
cursor = db.cursor()

sql = """SELECT 
mm_stock_comb,
mm_recommend,
mm_safety_stock_max
FROM mm_material_base
WHERE mm_no_gp = %s
"""
cursor.execute(sql,mmcode)
data = cursor.fetchall()
db.close()
numbers = pd.DataFrame(data)
stock = numbers[0][0]
#print(stock)
recommend = numbers[1][0]
#print(recommend)
if recommend < 1:
    recommend = 1
#print(recommend)

if trydata.at[0,'mmlevel'] =="S" or trydata.at[0,'mmlevel'] =="A" or trydata.at[0,'mmlevel'] =="B":
    buy = float(round(np.mean(a)+np.std(a)*3.08+max(a),0)-stock)
    if stock >= float(round(2*np.mean(a)+np.std(a)*3.08,0)):
        buy = 0
    if buy%recommend != 0:
        buy = ((buy//recommend)+1)*recommend


    #db = pymysql.connect(host='192.168.8.156', user='view1', passwd='q123456', db='erp_db')
    #cursor = db.cursor()

    #sql = "UPDATE mm_material_base SET mm_alert = %s, mm_rop = %s, mm_recommend = %s WHERE mm_no_gp = %s"

    #changelist = (float(round(np.mean(a)+np.std(a)*3.08,0)),float(round(2*np.mean(a)+np.std(a)*3.08,0)),float(round(np.mean(a)+np.std(a)*3.08+max(a),0)-stock),mmcode)
    #cursor.execute(sql,changelist)
    changelist = [float(round(np.mean(a)+np.std(a)*3.08,0)),float(round(2*np.mean(a)+np.std(a)*3.08,0)),buy,round(buy/np.mean(a[0:12]))]
else:
    buy = float(round(np.mean(a)+np.std(a)*1.65+max(a),0)-stock)
    if stock >= float(round(2*np.mean(a)+np.std(a)*1.65,0)):
        buy = 0
    if buy%recommend != 0:
        buy = ((buy//recommend)+1)*recommend


    #db = pymysql.connect(host='192.168.8.156', user='view1', passwd='q123456', db='erp_db')
    #cursor = db.cursor()

    #sql = "UPDATE mm_material_base SET mm_alert = %s, mm_rop = %s, mm_recommend = %s WHERE mm_no_gp = %s"

    #changelist = (float(round(np.mean(a)+np.std(a)*3.08,0)),float(round(2*np.mean(a)+np.std(a)*3.08,0)),float(round(np.mean(a)+np.std(a)*3.08+max(a),0)-stock),mmcode)
    #cursor.execute(sql,changelist)
    changelist = [float(round(np.mean(a)+np.std(a)*1.65,0)),float(round(2*np.mean(a)+np.std(a)*1.65,0)),buy,round(buy/np.mean(a[0:12]))]

if numbers[2][0] != 0:
    mm_rop_max = numbers[2][0]-1
    mm_propose_max = numbers[2][0]-stock
    changelist.append(mm_rop_max,mm_propose_max)
else:
    changelist.append(0,0)


print(changelist)