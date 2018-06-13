from datetime import datetime
from flask import Flask,session, request, flash, url_for, redirect, render_template, abort ,g,send_from_directory,jsonify
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
from flask_login import login_user , logout_user , current_user , login_required
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
from flask_login import UserMixin
from wtforms import StringField, PasswordField, BooleanField, SubmitField
import xlrd,xlwt
import mysql.connector
import numpy as np
import pandas as pd
from flask import send_file,make_response
from io import BytesIO
import re
import time
import os
from xlrd import xldate_as_tuple


app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql://root@localhost:3306/fddrca?charset=utf8'
app.config['SQLALCHEMY_COMMIT_ON_TEARDOWN'] = True
app.config['SQLALCHEMY_ECHO'] = False
app.config['SECRET_KEY'] = 'secret_key'
app.config['DEBUG'] = True
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False


UPLOAD_FOLDER = 'upload'
UPLOAD_FOLDER1 = 'apupload'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
basedir = os.path.abspath(os.path.dirname(__file__))
ALLOWED_EXTENSIONS = set(['txt', 'png', 'jpg', 'xls', 'JPG', 'PNG', 'xlsx', 'gif', 'GIF','xlsm'])


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS


@app.route('/test/upload')
def upload_test():
    return render_template('upload.html')

@app.route('/api/upload', methods=['POST'], strict_slashes=False)
def api_upload():
    file_dir = os.path.join(basedir, app.config['UPLOAD_FOLDER'])
    if not os.path.exists(file_dir):
        os.makedirs(file_dir)
    f=request.files['fileField']
    print f
    print 'fakepath**************'
    if f and allowed_file(f.filename):
        fname=f.filename
        ext = fname.rsplit('.', 1)[1]
        unix_time = int(time.time())
        new_filename = str(unix_time)+'.'+ext
        filename=os.path.join(file_dir, new_filename)
        print filename
        f.save(os.path.join(file_dir, new_filename))
        val=importfromexcel(filename)
        if val==0:
            flash('PRID has been there !', 'error')
            todo=Todo.query.order_by(Todo.PRID.desc()).first()
            a = re.sub("\D", "", todo.PRID)
                    #a=filter(str.isdigit, todo.PRID)
            a=int(a)
            a=a+1

            b=len(str(a))
                    #sr=sr+'m'*(9-len(sr))
            PRID='MAC'+'0'*(6-b)+ str(a)
            return render_template('new.html'
                                    ,PRID=PRID)
        flash('Pronto item has been successfully imported')
        return redirect(url_for('index'))
    else:
        flash('Invalid Filename!')
        return redirect(url_for('index'))
"""
        return jsonify({"errno": 0, "errmsg": "upload ok"})
    else:
        return jsonify({"errno": 1001, "errmsg": "upload fail"})
"""




db = SQLAlchemy(app)

Base=declarative_base()
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

class User(db.Model):
    __tablename__ = "users"
    id = db.Column('user_id',db.Integer , primary_key=True)
    username = db.Column('username', db.String(20), unique=True , index=True)
    password = db.Column('password' , db.String(250))
    email = db.Column('email',db.String(50),unique=True , index=True)
    registered_on = db.Column('registered_on' , db.DateTime)
    todos = db.relationship('Todo' , backref='user',lazy='select')
    todoaps= db.relationship('TodoAP' , backref='user',lazy='select')
    todolongcycletimercas = db.relationship('TodoLongCycleTimeRCA' , backref='user',lazy='select')
    def __init__(self , username ,password , email):
        self.username = username
        self.set_password(password)
        self.email = email
        self.registered_on = datetime.utcnow()

    def set_password(self , password):
        self.password = generate_password_hash(password)

    def check_password(self , password):
        return check_password_hash(self.password , password)

    def is_authenticated(self):
        return True

    def is_active(self):
        return True

    def is_anonymous(self):
        return False

    def get_id(self):
        return unicode(self.id)

    def __repr__(self):
        return '<User %r>' % (self.username)

class Todo(db.Model):
    __tablename__ = 'rcastatus'
    PRID = db.Column('PRID', db.String(64), primary_key=True)
    PRTitle = db.Column(db.String(1024))
    PRReportedDate = db.Column(db.String(64))
    PRClosedDate=db.Column(db.String(64))
    PROpenDays=db.Column(db.Integer)
    PRRcaCompleteDate = db.Column(db.String(64))
    PRRelease = db.Column(db.String(128))
    PRAttached = db.Column(db.String(128))
    IsLongCycleTime = db.Column(db.String(32))
    IsCatM = db.Column(db.String(32))
    IsRcaCompleted = db.Column(db.String(32))
    NoNeedDoRCAReason = db.Column(db.String(64))
    RootCauseCategory=db.Column(db.String(1024))
    FunctionArea = db.Column(db.String(1024))
    CodeDeficiencyDescription = db.Column(db.String(1024))
    CorrectionDescription=db.Column(db.String(1024))
    RootCause = db.Column(db.String(1024))
    IntroducedBy = db.Column(db.String(128))
    Handler = db.Column(db.String(64))
    rca5whys = db.relationship('Rca5Why' , backref='todo',lazy='select')
    user_id = db.Column(db.Integer, db.ForeignKey('users.user_id'))

    def __init__(self, PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,PRRelease,PRAttached,IsLongCycleTime,\
                 IsCatM,IsRcaCompleted,NoNeedDoRCAReason,RootCauseCategory,FunctionArea,CodeDeficiencyDescription,\
		 CorrectionDescription,RootCause,IntroducedBy,Handler):
        self.PRID = PRID
        self.PRTitle = PRTitle
        self.PRReportedDate = PRReportedDate
        self.PRClosedDate = PRClosedDate
        self.PROpenDays = PROpenDays
        self.PRRcaCompleteDate = PRRcaCompleteDate
        self.PRRelease = PRRelease
        self.PRAttached = PRAttached
        self.IsLongCycleTime = IsLongCycleTime
        self.IsCatM = IsCatM
        self.IsRcaCompleted = IsRcaCompleted
        self.NoNeedDoRCAReason = NoNeedDoRCAReason
        self.RootCauseCategory = RootCauseCategory
        self.FunctionArea = FunctionArea
        self.CodeDeficiencyDescription = CodeDeficiencyDescription
        self.CorrectionDescription = CorrectionDescription
        self.RootCause = RootCause
        self.IntroducedBy = IntroducedBy
        self.Handler = Handler

    conn=mysql.connector.connect(host='localhost',user='root',passwd='',port=3306)
    cur=conn.cursor()
    cur.execute('create database if not exists fddrca')
    conn.commit()
    cur.close()
    conn.close()

class TodoAP(db.Model):
    __tablename__ = 'apstatus'
    APID = db.Column('APID', db.String(64), primary_key=True)
    PRID = db.Column(db.String(64))
    APDescription = db.Column(db.String(1024))
    APCreatedDate = db.Column(db.String(64))
    APDueDate = db.Column(db.String(64))
    APCompletedOn = db.Column(db.String(64))
    IsApCompleted = db.Column(db.String(32))
    APAssingnedTo = db.Column(db.String(128))
    QualityOwner = db.Column(db.String(128))
    user_id = db.Column(db.Integer, db.ForeignKey('users.user_id'))

    def __init__(self, APID,PRID,APDescription,APCreatedDate,APDueDate,APCompletedOn,IsApCompleted,APAssingnedTo,QualityOwner):
        self.APID = APID
        self.PRID = PRID
        self.APDescription = APDescription
        self.APCreatedDate = APCreatedDate
        self.APDueDate = APDueDate
        self.APCompletedOn = APCompletedOn
        self.IsApCompleted = IsApCompleted
        self.APAssingnedTo = APAssingnedTo
        self.QualityOwner = QualityOwner

	#user_id = db.Column(db.Integer, db.ForeignKey('users.user_id'))

class Rca5Why(db.Model):
    __tablename__ = 'rca5why'
    id = db.Column('why_id', db.Integer, primary_key=True)
    PRID=db.Column(db.String(64))
    Why1 = db.Column(db.String(1024))
    Why2 = db.Column(db.String(1024))
    Why3=db.Column(db.String(1024))
    Why4 = db.Column(db.String(1024))
    Why5 = db.Column(db.String(1024))
    pr_id = db.Column(db.String(64), db.ForeignKey('rcastatus.PRID'))

    def __init__(self, PRID,Why1,Why2,Why3,Why4,Why5):
        self.PRID = PRID
        self.Why1 = Why1
        self.Why2 = Why2
        self.Why3 = Why3
        self.Why4 = Why4
        self.Why5 = Why5
        #pr_id = db.Column(db.String(64), db.ForeignKey('rcastatus.PRID'))

class TodoLongCycleTimeRCA(db.Model):
    __tablename__ = 'longcycletimercastatus'
    PRID = db.Column('PRID', db.String(64), primary_key=True)
    PRTitle = db.Column(db.String(1024))
    PRReportedDate = db.Column(db.String(64))
    PRClosedDate=db.Column(db.String(64))
    PROpenDays=db.Column(db.Integer)
    PRRcaCompleteDate = db.Column(db.String(64))
    IsLongCycleTime = db.Column(db.String(32))
    IsCatM = db.Column(db.String(32))
    LongCycleTimeRcaIsCompleted = db.Column(db.String(32))
    LongCycleTimeRootCause = db.Column(db.String(1024))
    NoNeedDoRCAReason = db.Column(db.String(64))
    Handler = db.Column(db.String(64))
    user_id = db.Column(db.Integer, db.ForeignKey('users.user_id'))

    def __init__(self, PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,IsLongCycleTime,\
                 IsCatM,LongCycleTimeRcaIsCompleted,LongCycleTimeRootCause,NoNeedDoRCAReason,Handler):
        self.PRID = PRID
        self.PRTitle = PRTitle
        self.PRReportedDate = PRReportedDate
        self.PRClosedDate = PRClosedDate
        self.PRRcaCompleteDate = PRRcaCompleteDate
        self.PROpenDays = PROpenDays
        self.IsLongCycleTime = IsLongCycleTime
        self.IsCatM = IsCatM
        self.LongCycleTimeRcaIsCompleted = LongCycleTimeRcaIsCompleted
        self.LongCycleTimeRootCause = LongCycleTimeRootCause
        self.NoNeedDoRCAReason = NoNeedDoRCAReason
        self.Handler = Handler

db.create_all()

app.config['dbconfig'] = {'host': '127.0.0.1',
                          'user': 'root',
                          'password': '',
                          'database': 'fddrca', }

class UseDatabase:
    def __init__(self, config):
        """Add the database configuration parameters to the object.

        This class expects a single dictionary argument which needs to assign
        the appropriate values to (at least) the following keys:

            host - the IP address of the host running MySQL/MariaDB.
            user - the MySQL/MariaDB username to use.
            password - the user's password.
            database - the name of the database to use.

        For more options, refer to the mysql-connector-python documentation.
        """
        self.configuration = config

    def __enter__(self):
        """Connect to database and create a DB cursor.

        Return the database cursor to the context manager.
        """
        self.conn = mysql.connector.connect(**self.configuration)
        self.cursor = self.conn.cursor()
        return self.cursor

    def __exit__(self, exc_type, exc_value, exc_traceback):
        """Destroy the cursor as well as the connection (after committing).
        """
        self.conn.commit()
        self.cursor.close()
        self.conn.close()


def modifyColumnType(fieldname):
    with UseDatabase(app.config['dbconfig']) as cursor:
        #alter table user MODIFY new1 VARCHAR(1) -->modify field type
        _SQL = "alter table rcastatus MODIFY `"+fieldname+"` VARCHAR(512)"
        cursor.execute(_SQL)

def compare_time(start_t,end_t):
    s_time = time.mktime(time.strptime(start_t,'%Y-%m-%d'))
    #get the seconds for specify date
    e_time = time.mktime(time.strptime(end_t,'%Y-%m-%d'))
    if float(s_time) >= float(e_time):
        return True
    return False

def comparetime(start_t,end_t):
    s_time = time.mktime(time.strptime(start_t,'%Y-%m-%d'))
    #get the seconds for specify date
    e_time = time.mktime(time.strptime(end_t,'%Y-%m-%d'))
    if(float(e_time)- float(s_time)) > float(86400):
        print ("@@@float(e_time)- float(s_time))=%f"%(float(e_time)- float(s_time)))
        return True
    return False

def leap_year(y):
    if (y % 4 == 0 and y % 100 != 0) or y % 400 == 0:
        return True
    else:
        return False

def days_in_month(y, m):
    if m in [1, 3, 5, 7, 8, 10, 12]:
        return 31
    elif m in [4, 6, 9, 11]:
        return 30
    else:
        if leap_year(y):
            return 29
        else:
            return 28

def days_this_year(year):
    if leap_year(year):
        return 366
    else:
        return 365

def days_passed(year, month, day):
    m = 1
    days = 0
    while m < month:
        days += days_in_month(year, m)
        m += 1
    return days + day

def dateIsBefore(year1, month1, day1, year2, month2, day2):
    """Returns True if year1-month1-day1 is before year2-month2-day2. Otherwise, returns False."""
    if year1 < year2:
        return True
    if year1 == year2:
        if month1 < month2:
            return True
        if month1 == month2:
            return day1 < day2
    return False

def daysBetweenDates(year1, month1, day1, year2, month2, day2):
    if year1 == year2:
        return days_passed(year2, month2, day2) - days_passed(year1, month1, day1)
    else:
        sum1 = 0
        y1 = year1
        while y1 < year2:
            sum1 += days_this_year(y1)
            y1 += 1
        return sum1-days_passed(year1,month1,day1)+days_passed(year2,month2,day2)

"""
    ip_set = [int(i) for i in ip_addr.split('.')]
    ip_number = (ip_set[0] << 24) + (ip_set[1] << 16) + (ip_set[2] << 8) + ip_set[3]
    return ip_number
    ext = fname.rsplit('.', 1)[1]
    
"""
def daysBetweenDate(start,end):
    year1=int(start.split('-',2)[0])
    month1=int(start.split('-',2)[1])
    day1=int(start.split('-',2)[2])

    year2=int(end.split('-',2)[0])
    month2=int(end.split('-',2)[1])
    day2=int(end.split('-',2)[2])
    print ("daysBetweenDates(year1, month1, day1, year2, month2, day2)=%d"%daysBetweenDates(year1, month1, day1, year2, month2, day2))
    return daysBetweenDates(year1, month1, day1, year2, month2, day2)

def insert_item(team,internaltask_sheet,i):
    PRID = internaltask_sheet.cell_value(i+1,2)
    print PRID
    PRTitle = internaltask_sheet.cell_value(i+1,9)
    PRReportedDate = internaltask_sheet.cell_value(i+1,5)
    PRClosedDate =internaltask_sheet.cell_value(i+1,44)
    PROpenDays=daysBetweenDate(PRReportedDate,PRClosedDate)
    PRRcaCompleteDate =''

    PRRelease = internaltask_sheet.cell_value(i+1,6)
    PRAttached = internaltask_sheet.cell_value(i+1,27)

    if daysBetweenDate(PRReportedDate,PRClosedDate) > 14:
        IsLongCycleTime = 'Yes'
    else:
        IsLongCycleTime = 'No'

    #IsLongCycleTime= internaltask_sheet.cell_value(i+1,34)
    IsCatM =''
    IsRcaCompleted='No'
    LongCycleTimeRcaIsCompleted='No'
    NoNeedDoRCAReason =''
    RootCauseCategory = ''
    FunctionArea = ''

    CodeDeficiencyDescription = ''
    CorrectionDescription = ''
    RootCause=''
    LongCycleTimeRootCause=''

    IntroducedBy=''
    Handler = team

    todo_item = Todo.query.get(PRID)

    registered_user = Todo.query.filter_by(PRID=PRID).all()
    if len(registered_user) ==0:
	print 'OK#################################################'
	todo = Todo(PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,PRRelease,PRAttached,IsLongCycleTime,\
                 IsCatM,IsRcaCompleted,NoNeedDoRCAReason,RootCauseCategory,FunctionArea,CodeDeficiencyDescription,\
		 CorrectionDescription,RootCause,IntroducedBy,Handler)
        #g.user=Todo.query.get(team)
	todo.user = g.user
	print("todo.user = g.user=%s"%todo.user)
	hello = User.query.filter_by(username=team).first()
	todo.user_id=hello.id
	print("todo.user_id=hello.user_id=%s"%hello.id)
	db.session.add(todo)
	db.session.commit()
    else:
        print ("registered_user.PRTitle=%s"%PRID)

    registered_user = TodoLongCycleTimeRCA.query.filter_by(PRID=PRID).all()
    if len(registered_user) ==0 and IsLongCycleTime=='Yes':
	print 'OK#################################################'
	todo = TodoLongCycleTimeRCA(PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,IsLongCycleTime,\
                 IsCatM,LongCycleTimeRcaIsCompleted,LongCycleTimeRootCause,NoNeedDoRCAReason,Handler)
        #g.user=Todo.query.get(team)
	todo.user = g.user
	print("todo.user = g.user=%s"%todo.user)
	hello = User.query.filter_by(username=team).first()
	todo.user_id=hello.id
	print("todo.user_id=hello.user_id=%s"%hello.id)
	db.session.add(todo)
	db.session.commit()
    else:
        print ("registered_user.PRTitle=%s"%PRID)

def importfromexcel(filename):
    workbook = xlrd.open_workbook(filename)
    internaltask_sheet=workbook.sheet_by_name(r'PR List')
    rows=internaltask_sheet.row_values(0)
    nrows=internaltask_sheet.nrows
    ncols=internaltask_sheet.ncols
    #modifyColumnType('PRTitle')
    print str(nrows)+"*********"
    print ncols
    k=0
    teams=['chenlong','xiezhen','yangjinyong','hanbing','lanshenghai','liumingjing','lizhongyuan','caizhichao']
    for i in range(nrows-1):
        ClosedEnter= internaltask_sheet.cell_value(i+1,44)
        PRState= internaltask_sheet.cell_value(i+1,8)
        team= internaltask_sheet.cell_value(i+1,63)
        print ClosedEnter
        if PRState=='Closed':
            time='2018-1-1'
            if compare_time (ClosedEnter,time):
                if team in teams:
                    print team
                    k=k+1
                    teamname=teams[teams.index(team)]
                    print teamname
                    insert_item(team,internaltask_sheet,i)

    val=1
    print k
    print 'closed pr'
    return val

def get_apid():
    todoap=TodoAP.query.order_by(TodoAP.APID.desc()).first()
    if todoap is not None:
        a = re.sub("\D", "", todoap.APID)
        a=int(a)
        a=a+1
        b=len(str(a))
    else:
        a=1
        b=1

    APID='AP'+'0'*(6-b)+ str(a)
    return APID
def update_rca(PRID,internaltask_sheet):
    todo_item = Todo.query.get(PRID)
    if todo_item is None:
        flash('Please check the PR, seems it is not in the Formal RCA PR list!', 'error')
        return False
    todo_item.PRRcaCompleteDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    todo_item.IsRcaCompleted = 'Yes'
    todo_item.RootCause  = internaltask_sheet.cell_value(20,11)
    """
    with UseDatabase(app.config['dbconfig']) as cursor:
        _SQL = "alter table rcastatus MODIFY column FunctionArea VARCHAR(1024)"
        cursor.execute(_SQL)
    """
    todo_item.CodeDeficiencyDescription = internaltask_sheet.row_values(6)[3]
    todo_item.CorrectionDescription  = internaltask_sheet.row_values(7)[3]
    db.session.commit()
    return True

def update_longcycletimerca(PRID,internaltask_sheet):
    todo_item = TodoLongCycleTimeRCA.query.get(PRID)
    if todo_item is None:
        flash('No in LongCycleTimeRca PR list!', 'error')
        return False
    todo_item.PRRcaCompleteDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    todo_item.LongCycleTimeRcaIsCompleted = 'Yes'
    todo_item.LongCycleTimeRootCause  = internaltask_sheet.cell_value(21,11)

    db.session.commit()
    return True

def insert_ap(internaltask_sheet,index):
    PRID = internaltask_sheet.cell_value(2,1).strip()
    APCompletedOn=''
    IsApCompleted='No'
    APDescription = internaltask_sheet.cell_value(index,12)
    if len(APDescription)!=0:
	APCreatedDate=time.strftime('%Y-%m-%d',time.localtime(time.time()))
	APAssingnedTo=internaltask_sheet.cell_value(index,15)
	#ctype:0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
	ctype= internaltask_sheet.cell(index,17).ctype
	APDueDate= internaltask_sheet.cell_value(index,17)
	if ctype==3:
            date = datetime(*xldate_as_tuple(APDueDate, 0))
            APDueDate = date.strftime('%Y-%m-%d')
        else:
            date=APDueDate
	#APDueDate=APDueDate.strftime('%Y-%d-%m')
	APID=get_apid()
	QualityOwner=User.query.get(g.user.id).username #g.user.id
	todoap = TodoAP(APID,PRID,APDescription,APCreatedDate,APDueDate,APCompletedOn,IsApCompleted,APAssingnedTo,QualityOwner)
	todoap.user = g.user
	db.session.add(todoap)
	db.session.commit()

def insert_rca5why(internaltask_sheet,index):
    PRID = internaltask_sheet.cell_value(2,1).strip()
    Why1 = internaltask_sheet.cell_value(index,2)
    Why2 = internaltask_sheet.cell_value(index,4)
    Why3= internaltask_sheet.cell_value(index,6)
    Why4 = internaltask_sheet.cell_value(index,8)
    Why5 = internaltask_sheet.cell_value(index,10)
    rca5why = Rca5Why(PRID,Why1,Why2,Why3,Why4,Why5)
    rca5why.pr_id = PRID
    db.session.add(rca5why)
    db.session.commit()

def findIndex(index_start,target_String,internaltask_sheet):
    APDescription = internaltask_sheet.cell_value(index_start,12)
    for i in range(10):
        APDescription = internaltask_sheet.cell_value(index_start+i,12)
        if APDescription==target_String:
            return index_start+i+1

def find5whyIndex(index_start,target_String,internaltask_sheet):
    APDescription = internaltask_sheet.cell_value(index_start,1)
    for i in range(10):
        APDescription = internaltask_sheet.cell_value(index_start+i,1)
        if APDescription==target_String:
            return index_start+i+1

def import_ap_fromexcel(filename):
    workbook = xlrd.open_workbook(filename)
    internaltask_sheet=workbook.sheet_by_name(r'RcaEda')
    rows=internaltask_sheet.row_values(0)
    nrows=internaltask_sheet.nrows
    ncols=internaltask_sheet.ncols
    print ("nrows==%d"%nrows)
    print ("#############nrows==%d"%ncols)
    PRID = internaltask_sheet.cell_value(2,1).strip()

    APCompletedOn=''
    IsApCompleted=''
    QualityOwner=User.query.get(g.user.id).username #g.user.id
    todo=TodoAP.query.filter_by(PRID = PRID).order_by(TodoAP.APID.asc()).first()
    state=0
    a=update_rca(PRID,internaltask_sheet)
    if a is False:
        print ("PRID=%s is not in the PR list of RCA candidate"%PRID)
        return False

    a=update_longcycletimerca(PRID,internaltask_sheet)
    if a is False:
        print ("PRID=%s is not in the LongCycleTime PR list"%PRID)
        #return False
    b=find5whyIndex(16,'Root Cause Analysis',internaltask_sheet)
    insert_rca5why(internaltask_sheet,b)
    insert_rca5why(internaltask_sheet,b+1)
    insert_rca5why(internaltask_sheet,b+2)
    print("import_ap_fromexcel.PRID=%s"%PRID)
    prid=str(PRID)
    """
    with UseDatabase(app.config['dbconfig']) as cursor:
        _SQL = "select * from apstatus where PRID=`"+prid+""
        cursor.execute(_SQL)
        contents = cursor.fetchall()
        if contents
        _SQL = "delete from apstatus where PRID=`"+prid+""
        cursor.execute(_SQL)
    """
    todo_item = TodoAP.query.filter_by(PRID = prid).all()
    if len(todo_item)==0:
        print ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@iterms=======")
    else:
        for item in todo_item:
            db.session.delete(item)
            print("item.APID=%s,item.PRID=%s"%(item.APID,item.PRID)) +"&*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
        db.session.commit()
        #TodoAP.query.filter(PRID == PRID).delete()
        #db.session.query(TodoAP).filter(PRID == prid).delete()
	print ("*****TodoAP.query.filter(PRID == PRID).delete()***")

    if state==0:
        b=findIndex(16,'Action Proposal',internaltask_sheet)
        for i in range(4):
                index=b+i
                insert_ap(internaltask_sheet,index)

        b=findIndex(43,'Action Proposal',internaltask_sheet)
        for i in range(3):
                index=b+i
                insert_ap(internaltask_sheet,index)
        c=findIndex(b,'Action Proposal',internaltask_sheet)
        for i in range(3):
                index=c+i
                insert_ap(internaltask_sheet,index)
        d=findIndex(c,'Action Proposal',internaltask_sheet)
        for i in range(3):
                index=d+i
                insert_ap(internaltask_sheet,index)
        e=findIndex(d,'Action Proposal',internaltask_sheet)
        for i in range(3):
                index=e+i
                insert_ap(internaltask_sheet,index)
        f=findIndex(e,'Action Proposal',internaltask_sheet)
        for i in range(3):
                index=f+i
                insert_ap(internaltask_sheet,index)
        gg=findIndex(f,'Action Proposal',internaltask_sheet)
        if gg is None:
            return True
        for i in range(3):
                index=gg+i
                insert_ap(internaltask_sheet,index)
        h=findIndex(gg,'Action Proposal',internaltask_sheet)
        if h is None:
            return True
        for i in range(3):
                index=h+i
                insert_ap(internaltask_sheet,index)
    return True

class Excel():
    def export(self):

        output = BytesIO()

        writer = pd.ExcelWriter(output, engine='xlwt')
        workbook = writer.book

        worksheet= workbook.add_sheet('sheet1',cell_overwrite_ok=True)
        col=0
        row=1
        pattern = xlwt.Pattern() # Create the Pattern
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
        #style = xlwt.XFStyle() # Create the Pattern
        font = xlwt.Font() # Create the Font
        font.name = 'Times New Roman'
        font.bold = True
        #font.underline = True
        #font.italic = True
        style = xlwt.XFStyle() # Create the Style
        style.font = font # Apply the Font to the Style
        style.pattern = pattern # Add Pattern to Style
        columns=['PRID','PRTitle','PRReportedDate','PRClosedDate','PROpenDays','IsLongCycleTime','PRRcaCompleteDate','PRRelease','PRAttached','IsCatM','IsRcaCompleted',\
                 'NoNeedDoRCAReason','RootCauseCategory','FunctionArea','CodeDeficiencyDescription','CorrectionDescription','RootCause','LongCycleTimeRootCause','IntroducedBy','Handler']
        for item in columns:
            worksheet.col(col).width = 4333 # 3333 = 1" (one inch).
            worksheet.write(0, col,item,style)
            col+=1
        style = xlwt.XFStyle()
        style.num_format_str = 'M/D/YY' # Other options: D-MMM-YY, D-MMM, MMM-YY, h:mm, h:mm:ss, h:mm, h:mm:ss, M/D/YY h:mm, mm:ss, [h]:mm:ss, mm:ss.0
        alignment = xlwt.Alignment() # Create Alignment
        alignment.horz = xlwt.Alignment.HORZ_JUSTIFIED # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED,HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
        alignment.vert = xlwt.Alignment.VERT_TOP # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
        #style = xlwt.XFStyle() # Create Style
        style.alignment = alignment # Add Alignment to Style
        todos=Todo.query.filter_by(user_id = g.user.id).order_by(Todo.Handler.desc()).all()
        nrows=len(todos)
        print ('nrows==%s' %(nrows))
        for row in range(nrows):
            row1=row+1
            worksheet.write(row1,0,todos[row].PRID)
            worksheet.write(row1,1,todos[row].PRTitle,style)
            worksheet.write(row1,2,todos[row].PRReportedDate)
            worksheet.write(row1,3,todos[row].PRClosedDate)
            worksheet.write(row1,4,todos[row].PROpenDays)
            worksheet.write(row1,5,todos[row].IsLongCycleTime)
            worksheet.write(row1,6,todos[row].PRRcaCompleteDate)
            worksheet.write(row1,7,todos[row].PRRelease)
            worksheet.write(row1,8,todos[row].PRAttached,style)
            worksheet.write(row1,9,todos[row].IsCatM)
            worksheet.write(row1,10,todos[row].IsRcaCompleted)
            worksheet.write(row1,11,todos[row].NoNeedDoRCAReason)
            worksheet.write(row1,12,todos[row].RootCauseCategory)
            worksheet.write(row1,13,todos[row].FunctionArea)
            worksheet.write(row1,14,todos[row].CodeDeficiencyDescription)
            worksheet.write(row1,15,todos[row].CorrectionDescription)
            worksheet.write(row1,16,todos[row].RootCause)
            if todos[row].IsLongCycleTime is 'Yes':
                todo=TodoLongCycleTimeRCA.query.filter_by(PRID = todos[row].PRID).first()
                worksheet.write(row1,17,todo.LongCycleTimeRootCause)
            else:
                worksheet.write(row1,17,'N/A')
            worksheet.write(row1,18,todos[row].IntroducedBy)
            worksheet.write(row1,19,todos[row].Handler)

        writer.close()
        output.seek(0)
        return output

class apExcel():
    def export(self):

        output = BytesIO()

        writer = pd.ExcelWriter(output, engine='xlwt')
        workbook = writer.book

        worksheet= workbook.add_sheet('sheet1',cell_overwrite_ok=True)
        col=0
        row=1
        pattern = xlwt.Pattern() # Create the Pattern
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
        pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
        #style = xlwt.XFStyle() # Create the Pattern
        font = xlwt.Font() # Create the Font
        font.name = 'Times New Roman'
        font.bold = True
        #font.underline = True
        #font.italic = True
        style = xlwt.XFStyle() # Create the Style
        style.font = font # Apply the Font to the Style
        style.pattern = pattern # Add Pattern to Style

        columns=['APID','PRID','APDescription','APCreatedDate','APDueDate','APCompletedOn','IsApCompleted','APAssingnedTo','QualityOwner']
        for item in columns:
            worksheet.col(col).width = 4333 # 3333 = 1" (one inch).
            worksheet.write(0, col,item,style)
            col+=1
        style = xlwt.XFStyle()
        style.num_format_str = 'M/D/YY' # Other options: D-MMM-YY, D-MMM, MMM-YY, h:mm, h:mm:ss, h:mm, h:mm:ss, M/D/YY h:mm, mm:ss, [h]:mm:ss, mm:ss.0
        alignment = xlwt.Alignment() # Create Alignment
        alignment.horz = xlwt.Alignment.HORZ_JUSTIFIED # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED,HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
        alignment.vert = xlwt.Alignment.VERT_TOP # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
        #style = xlwt.XFStyle() # Create Style
        style.alignment = alignment # Add Alignment to Style
        todos=TodoAP.query.filter_by(user_id = g.user.id).order_by(TodoAP.APID.asc()).all()
        nrows=len(todos)
        print ('nrows==%s' %(nrows))
        for row in range(nrows):
            r=row+1
            worksheet.write(r,0,todos[row].APID)
            worksheet.write(r,1,todos[row].PRID,style)
            worksheet.write(r,2,todos[row].APDescription,style)
            worksheet.write(r,3,todos[row].APCreatedDate)
            worksheet.write(r,4,todos[row].APDueDate,style)
            worksheet.write(r,5,todos[row].APCompletedOn,style)
            worksheet.write(r,6,todos[row].IsApCompleted)
            worksheet.write(r,7,todos[row].APAssingnedTo,style)
            worksheet.write(r,8,todos[row].QualityOwner)

            """
            for co in columns:
                column=columns.index(co)
                cellvalue=todos[row][column]
                worksheet.write(row,column,cellvalue)
            print('row===%s,index===%s'%(column,cellvalue))
            """

        #worksheet.set_column('A:E', 20)

        writer.close()
        output.seek(0)
        return output

@app.route('/dashboard')
def dashboard():
    return redirect(url_for('home1'))

@app.route('/api/apupload', methods=['POST'], strict_slashes=False)
def api_ap_upload():
    app.config['UPLOAD_FOLDER'] = 'apupload'
    file_dir = os.path.join(basedir, app.config['UPLOAD_FOLDER'])
    if not os.path.exists(file_dir):
        os.makedirs(file_dir)
    f=request.files['fileField']
    print f
    print 'fakepath**************'
    if f and allowed_file(f.filename):
        fname=f.filename
        ext = fname.rsplit('.', 1)[1]
        unix_time = int(time.time())
        new_filename = str(unix_time)+'.'+ext
        filename=os.path.join(file_dir, new_filename)
        print filename
        f.save(os.path.join(file_dir, new_filename))
        a=import_ap_fromexcel(filename)
        if a is True:
            flash('AP item have been successfully imported')
        else:
            flash('APs importing failed')
        return redirect(url_for('ap_home'))
    else:
        flash('Invalid Filename!')
        return redirect(url_for('ap_home'))



@app.route('/fromexcel',methods=['GET', 'POST'])
@login_required
def fromexcel():
    if g.user.username not in admin:
        flash('You are not permitted to import before you have passed the Audit!', 'error')
        return redirect(url_for('rca_home'))
    if request.method == 'GET':
        return render_template('fromexcel.html')
    else:
        print request.form['textfield']
        filename=request.form['textfield']
        val=importfromexcel(filename)
        if val==0:
            flash('PRID has been used! please use the recommanded one', 'error')
            todo=Todo.query.order_by(Todo.PRID.desc()).first()
            a = re.sub("\D", "", todo.PRID)
                    #a=filter(str.isdigit, todo.PRID)
            a=int(a)
            a=a+1

            b=len(str(a))
                    #sr=sr+'m'*(9-len(sr))
            PRID='MAC'+'0'*(6-b)+ str(a)
            return render_template('new.html'
                                    ,PRID=PRID)
        flash('Todo item was successfully imported')
        return redirect(url_for('index'))

@app.route('/importapfromexcel',methods=['GET', 'POST'])
@login_required
def importapfromexcel():
    if request.method == 'GET':
        return render_template('ap_fromexcel.html')
    else:
        filename=request.files['fileField']
        print filename
        import_ap_fromexcel(filename)
        flash('AP item have been successfully imported RCA status also has been updated!')
        return redirect(url_for('apindex'))
    return render_template('ap_new.html')

@app.route('/toexcel',methods=['GET', 'POST'])
@login_required
def toexcel():
    if request.method == 'GET':
        return render_template('toexcel.html')
    else:
        output=Excel().export()
        resp = make_response(output.getvalue())
        resp.headers["Content-Disposition"] ="attachment; filename=rca_pronto_list.xls"
        resp.headers['Content-Type'] = 'application/x-xlsx'
        flash('Export to excel successfully,Getting the result at bottom left!')
        return resp

@app.route('/toapexcel',methods=['GET', 'POST'])
@login_required
def toapexcel():
    if request.method == 'GET':
        return render_template('ap_toexcel.html')
    else:
        output=apExcel().export()
        resp = make_response(output.getvalue())
        resp.headers["Content-Disposition"] ="attachment; filename=rca_ap_list.xls"
        resp.headers['Content-Type'] = 'application/x-xlsx'
        flash('Export to excel successfully,Getting the result at bottom left!')
        return resp

@app.route('/',methods=['GET','POST'])
#@login_required
def home1():
    if request.method=='GET':
        return render_template('home.html')
    else:
        return redirect(url_for('rca_home'))


admin=['leienqing',]
@app.route('/rca_home',methods=['GET','POST'])
@login_required
def rca_home():
    if request.method=='GET':
        if g.user.username in admin:
            count = Todo.query.filter_by( NoNeedDoRCAReason='').order_by(Todo.PRClosedDate.asc()).count()

            return render_template('index.html',count= count,
                                   todos=Todo.query.filter_by( \
                                                              NoNeedDoRCAReason='').order_by(
                                       Todo.PRClosedDate.asc()).all(), \
                                   user=User.query.get(g.user.id).username + '  Logged in')
        else:
            count=Todo.query.filter_by(user_id = g.user.id,\
                               NoNeedDoRCAReason='').order_by(Todo.PRClosedDate.asc()).count()

            return render_template('index.html',count=count,
                               todos=Todo.query.filter_by(user_id = g.user.id,\
                               NoNeedDoRCAReason='').order_by(Todo.PRClosedDate.asc()).all(),\
                               user=User.query.get(g.user.id).username + '  Logged in')
    else:
        output=Excel().export()
        resp = make_response(output.getvalue())
        resp.headers["Content-Disposition"] ="attachment; filename=rca_pronto_list.xls"
        resp.headers['Content-Type'] = 'application/x-xlsx'
        return resp

@app.route('/rca_done',methods=['GET','POST'])
@login_required
def rca_done():
    if request.method=='GET':
        if g.user.username in admin:
            count = Todo.query.filter_by( NoNeedDoRCAReason='',IsRcaCompleted='Yes').order_by(Todo.PRClosedDate.asc()).count()

            return render_template('index.html',count= count,
                                   todos=Todo.query.filter_by( \
                                                              NoNeedDoRCAReason='',IsRcaCompleted='Yes').order_by(
                                       Todo.PRClosedDate.asc()).all(), \
                                   user=User.query.get(g.user.id).username + '  Logged in')
        else:
            count=Todo.query.filter_by(user_id = g.user.id,\
                               NoNeedDoRCAReason='',IsRcaCompleted='Yes').order_by(Todo.PRClosedDate.asc()).count()

            return render_template('index.html',count=count,
                               todos=Todo.query.filter_by(user_id = g.user.id,\
                               NoNeedDoRCAReason='',IsRcaCompleted='Yes').order_by(Todo.PRClosedDate.asc()).all(),\
                               user=User.query.get(g.user.id).username + '  Logged in')

@app.route('/rca_undone',methods=['GET','POST'])
@login_required
def rca_undone():
    if request.method=='GET':
        if g.user.username in admin:
            count = Todo.query.filter_by( NoNeedDoRCAReason='',IsRcaCompleted='No').order_by(Todo.PRClosedDate.asc()).count()

            return render_template('index.html',count= count,
                                   todos=Todo.query.filter_by( \
                                                              NoNeedDoRCAReason='',IsRcaCompleted='No').order_by(
                                       Todo.PRClosedDate.asc()).all(), \
                                   user=User.query.get(g.user.id).username + '  Logged in')
        else:
            count=Todo.query.filter_by(user_id = g.user.id,\
                               NoNeedDoRCAReason='',IsRcaCompleted='No').order_by(Todo.PRClosedDate.asc()).count()

            return render_template('index.html',count=count,
                               todos=Todo.query.filter_by(user_id = g.user.id,\
                               NoNeedDoRCAReason='',IsRcaCompleted='No').order_by(Todo.PRClosedDate.asc()).all(),\
                               user=User.query.get(g.user.id).username + '  Logged in')


@app.route('/longcycletimerca_home',methods=['GET','POST'])
@login_required
def longcycletimerca_home():
    if request.method=='GET':
        if g.user.username in admin:
            return render_template('longcycletimerca_index.html',count= TodoLongCycleTimeRCA.query.count(),
                               todos=TodoLongCycleTimeRCA.query.filter_by(NoNeedDoRCAReason='').order_by(TodoLongCycleTimeRCA.PRClosedDate.asc()).all(),user=User.query.get(g.user.id).username + '  Logged in')
        else:
            return render_template('longcycletimerca_index.html',
                                   todos=TodoLongCycleTimeRCA.query.filter_by(user_id=g.user.id,
                                                                              NoNeedDoRCAReason='').order_by(
                                       TodoLongCycleTimeRCA.PRClosedDate.asc()).all(),
                                   user=User.query.get(g.user.id).username + '  Logged in')


@app.route('/ap_home',methods=['GET','POST'])
@login_required
def ap_home():
    if request.method=='GET':
        if g.user.username in admin:
            return render_template('ap_index.html',
                               todos=TodoAP.query.order_by(TodoAP.APID.asc()).all())
        else:
            return render_template('ap_index.html',
                                   todos=TodoAP.query.filter_by(user_id=g.user.id).order_by(TodoAP.APID.asc()).all())
    else:
        output=Excel().export()
        resp = make_response(output.getvalue())
        resp.headers["Content-Disposition"] ="attachment; filename=ap_list.xls"
        resp.headers['Content-Type'] = 'application/x-xlsx'
        return resp

@app.route('/index',methods=['GET','POST'])
@login_required
def index():
    return redirect(url_for('rca_home'))

@app.route('/longcycletimercaindex',methods=['GET','POST'])
@login_required
def longcycletimercaindex():
    return redirect(url_for('longcycletimerca_home'))

@app.route('/apindex',methods=['GET','POST'])
@login_required
def apindex():
    return redirect(url_for('ap_home'))

@app.route('/new', methods=['GET', 'POST'])
@login_required
def new():
    if request.method == 'POST':
        if not request.form['PRID']:
            flash('PRID is required', 'error')
        elif not request.form['PRTitle']:
            flash('PRTitle is required', 'error')
        else:
            PRID=request.form['PRID'].strip()

            todo=Todo.query.filter_by(PRID=PRID).first()
            if todo is not None:
                flash('PRID has been used!', 'error')
                return render_template('new.html'
                                       ,PRID=PRID)
            PRID = request.form['PRID']
            PRTitle = request.form['PRTitle']
            PRClosedDate  = request.form['PRClosedDate']
            PRReportedDate = request.form['PRReportedDate']
            PROpenDays=daysBetweenDate(PRReportedDate,PRClosedDate)
            PRRcaCompleteDate = request.form['PRRcaCompleteDate']
            PRRelease  = request.form['PRRelease']
            PRAttached = request.form['PRAttached']
            IsCatM  = request.form['IsCatM']
            IsRcaCompleted = request.form['IsRcaCompleted']
            IsLongCycleTime  = request.form['IsLongCycleTime']
            NoNeedDoRCAReason  = request.form['NoNeedDoRCAReason']
            RootCauseCategory  = request.form['RootCauseCategory']
            FunctionArea = request.form['FunctionArea']

            CodeDeficiencyDescription  = request.form['CodeDeficiencyDescription']
            CorrectionDescription  = request.form['CorrectionDescription']
            RootCause = request.form['RootCause']

            IntroducedBy  = request.form['IntroducedBy']
            Handler  = request.form['Handler']

            todo = Todo(PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,PRRelease,PRAttached,IsLongCycleTime,\
                     IsCatM,IsRcaCompleted,NoNeedDoRCAReason,RootCauseCategory,FunctionArea,CodeDeficiencyDescription,CorrectionDescription,RootCause,IntroducedBy,Handler)
            todo.user = g.user
            db.session.add(todo)
            db.session.commit()
            flash('RCA item was successfully created')
            return redirect(url_for('index'))

    return render_template('new.html')

@app.route('/newlongcycletimerca', methods=['GET', 'POST'])
@login_required
def newlongcycletimerca():
    if request.method == 'POST':
        if not request.form['PRID']:
            flash('PRID is required', 'error')
        elif not request.form['PRTitle']:
            flash('PRTitle is required', 'error')
        else:
            PRID=request.form['PRID'].strip()

            todo=TodoLongCycleTimeRCA.query.filter_by(PRID=PRID).first()
            if todo is not None:
                flash('PRID has been used!', 'error')
                return render_template('longcycletimerca_new.html'
                                       ,PRID=PRID)
            PRID = request.form['PRID']
            PRTitle = request.form['PRTitle']
            PRClosedDate  = request.form['PRClosedDate']
            PRReportedDate = request.form['PRReportedDate']
            PROpenDays=daysBetweenDate(PRReportedDate,PRClosedDate)
            PRRcaCompleteDate = request.form['PRRcaCompleteDate']
            IsCatM  = request.form['IsCatM']
            LongCycleTimeRcaIsCompleted = request.form['LongCycleTimeRcaIsCompleted']
            if daysBetweenDate(PRReportedDate,PRClosedDate) > 14:
                IsLongCycleTime = 'Yes'
            else:
                IsLongCycleTime = 'No'
            NoNeedDoRCAReason  = request.form['NoNeedDoRCAReason']

            LongCycleTimeRootCause = request.form['LongCycleTimeRootCause']

            Handler  = request.form['Handler']

            todo = TodoLongCycleTimeRCA(PRID,PRTitle,PRReportedDate,PRClosedDate,PROpenDays,PRRcaCompleteDate,IsLongCycleTime,\
                                         IsCatM,LongCycleTimeRcaIsCompleted,LongCycleTimeRootCause,NoNeedDoRCAReason,Handler)
            todo.user = g.user
            db.session.add(todo)
            db.session.commit()
            flash('LongCycleTime RCA item was successfully created')
            return redirect(url_for('longcycletimercaindex'))

    return render_template('longcycletimerca_new.html')


@app.route('/newap', methods=['GET', 'POST'])
@login_required
def newap():
    if request.method == 'POST':
        if not request.form['APID']:
            flash('APID is required', 'error')
        elif not request.form['PRID']:
            flash('PRID is required', 'error')
        elif not request.form['APDescription']:
            flash('APDescription is required', 'error')
        else:
            PRID=request.form['PRID']
            APID=request.form['APID']
            todo=TodoAP.query.filter_by(APID=APID).first()
            if todo is not None:
                flash('APID has been used!', 'error')
                return render_template('ap_new.html',APID=get_apid())

            APDescription = request.form['APDescription']
            APCreatedDate = request.form['APCreatedDate']
            APDueDate  = request.form['APDueDate']
            APCompletedOn = request.form['APCompletedOn']
            IsApCompleted  = request.form['IsApCompleted']
            APAssingnedTo = request.form['APAssingnedTo']
            QualityOwner  = request.form['QualityOwner']

            todo = TodoAP(APID,PRID,APDescription,APCreatedDate,APDueDate,APCompletedOn,IsApCompleted,APAssingnedTo,QualityOwner)
            todo.user = g.user
            db.session.add(todo)
            db.session.commit()
            flash('AP item was successfully created')
            return redirect(url_for('apindex'))

    return render_template('ap_new.html',APID=get_apid())

def update_rca_team(PRID,internaltask_sheet):
    todo_item = Todo.query.get(PRID)
    todo_item.PRRcaCompleteDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    todo_item.IsRcaCompleted = 'Yes'
    todo_item.CodeDeficiencyDescription = internaltask_sheet.row_values(6)[3]
    todo_item.CorrectionDescription  = internaltask_sheet.row_values(7)[3]
    todo_item.RootCause = internaltask_sheet.cell_value(20,11)
    db.session.commit()

def update_longcycletimercatable(PRID,request):
    todo_item = TodoLongCycleTimeRCA.query.get(PRID)

    todo_item.PRTitle = request.form['PRTitle']
    todo_item.PRClosedDate  = request.form['PRClosedDate']
    todo_item.PRReportedDate = request.form['PRReportedDate']
    todo_item.PROpenDays = daysBetweenDate(todo_item.PRReportedDate,todo_item.PRClosedDate)

    todo_item.PRRcaCompleteDate = request.form['PRRcaCompleteDate']

    if daysBetweenDate(todo_item.PRReportedDate,todo_item.PRClosedDate) > 14:
        IsLongCycleTime = 'Yes'
    else:
        IsLongCycleTime = 'No'
    todo_item.IsCatM  = request.form['IsCatM']
    todo_item.LongCycleTimeRcaIsCompleted = request.form['IsRcaCompleted']
    todo_item.NoNeedDoRCAReason  = request.form['NoNeedDoRCAReason']
    todo_item.Handler  = request.form['Handler']
    team=request.form['Handler']

    hello = User.query.filter_by(username=team).first()
    todo_item.user_id=hello.id
    db.session.commit()





@app.route('/todos/<PRID>', methods = ['GET' , 'POST'])
@login_required
def show_or_update(PRID):
    todo_item = Todo.query.get(PRID)
    if request.method == 'GET':
        if todo_item.user_id == g.user.id or g.user.username in admin:
            print ("todo_item.user_id=%d"%todo_item.user_id)
            print ("g.user.id=%d"%g.user.id)
            print "before"
            return render_template('view.html',todo=todo_item)
        else:
            print "after"
            print ("todo_item.user_id=%d"%todo_item.user_id)
            print ("g.user.id=%d"%g.user.id)
            flash('This PR is not under your account,You are not authorized to edit this item, Please login with correct account', 'error')
            return redirect(url_for('logout'))
    elif request.method == 'POST':
        value = request.form['button']
        if value == 'Update':
            #if g.user.username in admin:
            if todo_item.user.id == g.user.id:
                todo_item.PRID = request.form['PRID'].strip()
                PRID=todo_item.PRID
                todo_item.PRTitle = request.form['PRTitle']

                todo_item.PRClosedDate  = request.form['PRClosedDate']

                todo_item.PRReportedDate = request.form['PRReportedDate']

                PRReportedDate= request.form['PRReportedDate']
                PRClosedDate= request.form['PRClosedDate']

                todo_item.PROpenDays = daysBetweenDate(PRReportedDate,PRClosedDate)

                todo_item.PRRcaCompleteDate = request.form['PRRcaCompleteDate']
                if todo_item.PROpenDays > 14:
                    IsLongCycleTime = 'Yes'
                else:
                    IsLongCycleTime = 'No'
                todo_item.IsLongCycleTime  = IsLongCycleTime
                todo_item.IsCatM  = request.form['IsCatM']
                todo_item.IsRcaCompleted = request.form['IsRcaCompleted']
                todo_item.NoNeedDoRCAReason  = request.form['NoNeedDoRCAReason']
                todo_item.RootCauseCategory  = request.form['RootCauseCategory']
                todo_item.FunctionArea = request.form['FunctionArea']
                todo_item.CodeDeficiencyDescription  = request.form['CodeDeficiencyDescription']
                todo_item.CorrectionDescription  = request.form['CorrectionDescription']
                todo_item.RootCause = request.form['RootCause']
                todo_item.IntroducedBy  = request.form['IntroducedBy']
                todo_item.Handler  = request.form['Handler']
                team=request.form['Handler']
                print("team=%s"%team)
                hello = User.query.filter_by(username=team).first()
                todo_item.user_id=hello.id
                db.session.commit()
                if IsLongCycleTime =='Yes':
                    update_longcycletimercatable(PRID,request)
                return redirect(url_for('index'))
            else:
                flash('You are not authorized to edit this item','error')
                return redirect(url_for('show_or_update',PRID=PRID))
        elif value == 'Delete'and g.user.username in admin:
            todo_item = Todo.query.get(PRID)
            db.session.delete(todo_item)
            todo_longcycle_item = TodoLongCycleTimeRCA.query.get(PRID)
            if todo_longcycle_item:
                db.session.delete(todo_longcycle_item)
            db.session.commit()
            return redirect(url_for('index'))
        flash('You are not authorized to Delete this item', 'error')
        return redirect(url_for('show_or_update', PRID=PRID))

def update_rcatable(PRID,request):
    todo_item = Todo.query.get(PRID)
    todo_item.PRTitle = request.form['PRTitle']
    todo_item.PRClosedDate  = request.form['PRClosedDate']
    todo_item.PRReportedDate = request.form['PRReportedDate']
    todo_item.PROpenDays = daysBetweenDate(todo_item.PRReportedDate,todo_item.PRClosedDate)

    todo_item.PRRcaCompleteDate = request.form['PRRcaCompleteDate']


    if daysBetweenDate(todo_item.PRReportedDate,todo_item.PRClosedDate) > 14:
        IsLongCycleTime = 'Yes'
    else:
        IsLongCycleTime = 'No'
    todo_item.IsCatM  = request.form['IsCatM']
    todo_item.IsRcaCompleted = request.form['LongCycleTimeRcaIsCompleted']
    todo_item.NoNeedDoRCAReason  = request.form['NoNeedDoRCAReason']

    todo_item.Handler  = request.form['Handler']
    team=request.form['Handler']

    hello = User.query.filter_by(username=team).first()
    todo_item.user_id=hello.id
    db.session.commit()

@app.route('/todolongcycletimercas/<PRID>', methods = ['GET' , 'POST'])
@login_required
def show_or_updatelongcycletimerca(PRID):
    todo_item = TodoLongCycleTimeRCA.query.get(PRID)
    if request.method == 'GET':
        return render_template('longcycletimerca_view.html',todo=todo_item)
    if todo_item.user.id == g.user.id:
        todo_item.PRID = request.form['PRID'].strip()
        PRID=todo_item.PRID
        todo_item.PRTitle = request.form['PRTitle']
        todo_item.PRClosedDate  = request.form['PRClosedDate']
        todo_item.PRReportedDate = request.form['PRReportedDate']
        todo_item.PROpenDays = daysBetweenDate(todo_item.PRReportedDate,todo_item.PRClosedDate)
        todo_item.PROpenDays = daysBetweenDate(todo_item.PRReportedDate,todo_item.PRClosedDate)

        todo_item.PRRcaCompleteDate = request.form['PRRcaCompleteDate']

        todo_item.IsLongCycleTime  = request.form['IsLongCycleTime']

        todo_item.IsCatM  = request.form['IsCatM']
        todo_item.LongCycleTimeRcaIsCompleted = request.form['LongCycleTimeRcaIsCompleted']

        todo_item.NoNeedDoRCAReason  = request.form['NoNeedDoRCAReason']

        todo_item.LongCycleTimeRootCause = request.form['LongCycleTimeRootCause']

        todo_item.Handler  = request.form['Handler']
        team=request.form['Handler']
	hello = User.query.filter_by(username=team).first()
	todo_item.user_id=hello.id

        db.session.commit()

        update_rcatable(PRID,request)

        return redirect(url_for('longcycletimercaindex'))

    flash('You are not authorized to edit this todo item','error')
    return redirect(url_for('show_or_updatelongcycletimerca',PRID=PRID))

@app.route('/todoaps/<APID>', methods = ['GET' , 'POST'])
@login_required
def show_or_updateap(APID):
    todo_item = TodoAP.query.get(APID)
    if request.method == 'GET':
        return render_template('ap_view.html',todo=todo_item)
    if todo_item.user.id == g.user.id:
        todo_item = TodoAP.query.get(APID)
        todo_item.PRID = request.form['PRID'].strip()
        todo_item.APDescription  = request.form['APDescription']
        todo_item.APCreatedDate = request.form['APCreatedDate']
        todo_item.APDueDate = request.form['APDueDate']
        todo_item.APCompletedOn  = request.form['APCompletedOn']
        todo_item.IsApCompleted = request.form['IsApCompleted']
        todo_item.APAssingnedTo  = request.form['APAssingnedTo']
        todo_item.QualityOwner  = request.form['QualityOwner']
        db.session.commit()
        return redirect(url_for('apindex'))
    flash('You are not authorized to edit this todo item','error')
    return redirect(url_for('show_or_updateap',APID=APID))


@app.route('/register' , methods=['GET','POST'])
def register():
    if request.method == 'GET':
        return render_template('register.html')
    username = request.form['username']
    password = request.form['password']
    user = User(request.form['username'] , request.form['password'],request.form['email'])
    registered_user = User.query.filter_by(username=username).first()
    if registered_user is None:
        db.session.add(user)
        db.session.commit()
        flash('User successfully registered')
        return redirect(url_for('login'))
    else:
        flash('User name has been used,please try other one')
        return redirect(url_for('register'))


@app.route('/login',methods=['GET','POST'])
def login():
    if request.method == 'GET':
        return render_template('login.html')

    username = request.form['username']
    password = request.form['password']
    remember_me = False
    if 'remember_me' in request.form:
        #remember_me = BooleanField('Keep me logged in')
        remember_me = True
    registered_user = User.query.filter_by(username=username).first()
    if registered_user is None:
        flash('Username is invalid' , 'error')
        return redirect(url_for('login'))
    if not registered_user.check_password(password):
        flash('Password is invalid','error')
        return redirect(url_for('login'))
    login_user(registered_user, remember = remember_me)
    flash('Logged in successfully')
    return redirect(request.args.get('next') or url_for('index'))

@app.route('/logout')
def logout():
    logout_user()
    return redirect(url_for('index'))

@login_manager.user_loader
def load_user(id):
    return User.query.get(int(id))

@app.before_request
def before_request():
    g.user = current_user

def sleeptime(hour,min,sec):
    return hour*3600+min*60+sec
def test():
    second=sleeptime(0,0,1)
    while True:
        time.sleep(second)
        print 'Time delay!!!!!!!!!!!!!!!!!!!!'

if __name__ == '__main__':
    app.run(debug=True,host='0.0.0.0',port=3344)
