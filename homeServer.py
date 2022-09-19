from flask import Flask, jsonify, request, send_from_directory, render_template
from pprint import pprint
from numpy import True_
import pymysql
from fileinput import filename
from docxWritter import docxWritter
from upp import FromExcelToMysql
import sys
import pathlib
from typing import Type, TypeVar, Union, Optional


HOST_LISTEN: str = "0.0.0.0"
PORT_LISTEN: int = 8080
Connection = TypeVar('Connection')
Cursor = TypeVar('Cursor')


class MysqlLover():
    def __init__(self, host: Optional[str] = 'ip_address', user: Optional[str] = 'user_name',
                 passwd: Optional[str] = 'password', db: Optional[str] = 'db_name',
                 port: Optional[int] = 'port') -> object:

        self.__host: str = host
        self.__user: str = user
        self.__passwd: str = passwd
        self.__db: str = db
        self.__port: int = port
        self.__address: str = 'ip_address'

    def Connect(self):
        self.__con__: Connection = pymysql.connect(host=self.__host, user=self.__user, passwd=self.__passwd,
                                                   db=self.__db, port=self.__port)
        self.cur: Cursor = self.__con__.cursor()

    def __getConn(self) -> str:
        return self.__address + " " + self.__user + " " + self.__passwd + \
            " " + self.__db + " " + str(self.__port)

    def end(self) -> None:
        self.__con__.commit()
        self.cur.close()
        self.__con__.close()

    def entr(self, log, pas):
        buf: str = 'bad'
        self.sql: str = "SELECT * FROM users WHERE userName = %s"
        print(self.sql)

        self.Connect()
        self.cur.execute(self.sql, (log))
        self.data = self.cur.fetchall()

        for row in self.data:
            if row[2] == pas:
                print('сработало ------------------')
                buf = self.__getConn()
                print(row[1])
                if row[1] == 'Admin':
                    buf += " "+"root"
                    print('buf = ', buf)
            else:
                buf = 'bad'
        self.end()
        return buf

    def SingUp(self, log, pas, key):
        buf = 'bad'
        self.Connect()
        self.sql = f"SELECT userName FROM users WHERE userKey = %s"
        print(self.sql)
        self.cur.execute(self.sql, (key))
        self.data = self.cur.fetchall()
        print(self.data[0][0])
        if str(self.data[0][0]) == 'None':
            self.sql = f"UPDATE users SET userName = %s, userPassword = %s WHERE userKey = %s"
            self.cur.execute(self.sql, (log, pas, key))
            self.end()
            buf = "succes"
        else:
            buf = 'bad'
        return buf


ExcelReader = FromExcelToMysql()


json = TypeVar('json')

pechat = docxWritter()
msql = MysqlLover()

app = Flask(__name__)

filename = 'file'


@app.route('/DBHome', methods=['GET', 'POST'])
def requestGet():
    content = request.headers.get('Content-Type')
    if (content == 'application/json'):
        data = request.json
        #r = request.json.load(data)
        pprint(data)
        return msql.entr(data['log'], data['pas'])


@app.route('/GetDataSQL', methods=['GET', 'POST'])
def GetDataSQL() -> str:
    content: str = request.headers.get('Content-Type')
    if (content == 'application/json'):
        data: json = request.json
        #r = request.json.load(data)
        pprint(data)
        return msql.SingUp(data['log'], data['pas'], data['key'])


@app.route('/FlaskMaster', methods=['GET', 'POST'])
def request1Get():
    content = request.headers.get('Content-Type')
    if (content == 'application/json'):
        data = request.json

        if data['method'] == 'pechat':
            pprint(data)
            r = pechat.pechat(data)
            print(r)
            return r[1]

        elif data['method'] == 'read':
            return msql.readFile(data['path'])
        else:
            return 'bad'
    return jsonify("hi")



@app.route('/FlaskMaster/Download', methods=['GET', 'POST'])
def request2Get():
    print(request.args.get('docx_name'))
    desktop_path = '/home/Sanya/filesDocx/'
    return send_from_directory(desktop_path, request.args.get('docx_name'), as_attachment=True)


@app.route('/UploadFile', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
 
        ExcelReader.readFile(request.data)


    if request.method == "GET":
        print("Get")
    return jsonify("Ok")


@app.route('/DR', methods=['GET', 'POST'])
def DR():
    return render_template('DR.html')


if __name__ == "__main__":
    app.run(host=HOST_LISTEN, port=PORT_LISTEN, debug=True)
