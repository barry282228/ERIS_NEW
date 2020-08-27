from werkzeug.utils import secure_filename
from flask import Flask,app,request,g,render_template,redirect,url_for,abort,send_file
from flask import jsonify,make_response,send_from_directory,flash
from contextlib import closing
import time
import sqlite3
import os
import base64
import random

app = Flask(__name__)

UPLOAD_FOLDER = 'C:\\Users\\minghsuantu\\Desktop\\flask try\\static\\upload'
app.secret_key = "secret key"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

app.config['MAX_CONTENT_LENGTH'] = 16*1024*1024
basedir = os.path.abspath(os.path.dirname(__file__))

ALLOWED_EXTENSIONS=set(['txt','png','PNG','jpg','xls','xlsx','DAT','dat','zip'])

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.',1)[1] in ALLOWED_EXTENSIONS

def get_FileSize(filePath):
    filePath = unicode(filePath,'utf8')
    fsize = os.path.getsize(filePath)
    fsize = fsize/float(1024*1024)
    return round(fsize,2)

@app.route('/')
def upload_test():
    return render_template('Upload.html')


@app.route('/',methods=['GET','POST'])
def api_upload():
    file = request.files['myfile']
    if file and allowed_file(file.filename):
	    filename = secure_filename(file.filename)
	    file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
	    flash('File successfully uploaded')
	    return redirect('/')
    

def getListFiles(path):
    ret=[]
    for root,dirs,files in os.walk(path):
        for filespath in files:
            ret.append(os.path.join(filespath))
    return ret

@app.route("/download")
def download_list():
    file_dir=os.path.join(basedir,app.config['UPLOAD_FOLDER'])
    f_list = getListFiles(file_dir)
    len_file=len(f_list)
    return render_template('Download_list.html',f_list=f_list,len_file=len_file)


@app.route("/download/<filename>",methods=['GET'])
def download(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'],filename,as_attachment=True)

@app.route("/delete/<filename>",methods=['GET'])
def delet(filename):
    file_dir = os.path.join(basedir,app.config['UPLOAD_FOLDER'])
    delfile = os.path.join(file_dir,filename)
    if os.path.exists(delfile):
        os.remove(delfile)
        return "ok<a href=\"..\download\">File List</a>"
    else:
        return "delete fail"



if __name__ == '__main__':
    app.run(debug=True,host="0.0.0.0",processes = 1)