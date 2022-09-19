import pymysql
import pandas as pd

class FromExcelToMysql():


    def __init__(self):

        self.CollumnName = ['Работник', 'Адрес', 'Кв№', 'Прибор', 'Счётчик', 'Заводской номер', 'Показания начало',
                            'Показания конец', 'единицы измерения', 'Сумма договора', 'Оплата', 'способ оплаты', 'банк', 'Заказчик',
                            'Клиент', 'Хозяин', 'номер телефона', 'Годность', 'ФИО', 'Новая сумма', 'Новый прибор', 'Батарея была',
                            'Номер накладой', 'Батарея стала', 'Дата снятия', 'Дата замены', 'Дата отправления',
                            'Дата монтажа', 'Дата поступления', 'ссылка']

    def readFile(self, file):

        k = pd.read_excel(file, dtype=str, index_col=0)
        res = pd.DataFrame()
        for i in list(k):
            for j in self.CollumnName:
                if i.lower() == j.lower():
                    res[j] = k[i]
        print(res)
        try:
            res['Адрес'] = res['Адрес'] + ' Кв№ ' + res['Кв№']
            res = res.drop(columns='Кв№')
        except:
            try:
                res = res.drop(columns='Кв№')
            except:
                pass
        res = res.fillna('')
        res.insert(1, 'DateCreate', pd.to_datetime(
            'today').strftime("%Y-%m-%d"))
        for i in self.CollumnName[-6:-1]:
            try:
                d1 = pd.to_datetime(res[i], format='%Y-%m-%d', errors='coerce')
                d2 = pd.to_datetime(pd.to_numeric(
                    res[i], errors='coerce'),  unit='d', origin='1900-01-01')
                res[i] = d1.combine_first(d2).astype(str)
            except:
                print("прошло")
                continue
            k = None
            print("Первое хорошо")
        return self.Insert(res)

    def Connect(self):

        self.__con__ = pymysql.connect(
            host='ip_address', user='user_name', passwd='password', db='db_name', port="port")
        self.cur = self.__con__.cursor()

    def pechat(self):

        self.Connect()
        self.sql = "SELECT * FROM clients"
        self.cur.execute(self.sql)
        self.data = self.cur.fetchall()
        data = pd.read_sql('SELECT * FROM clients', self.__con__)
        print(data)
        # for row in self.data:
        #     print(row)
        self.__con__.commit()
        self.cur.close()
        self.__con__.close()

    def Insert(self, frame):

        try:
            self.Connect()
            CollumnName = {'Работник': 'Rabotnik', 'Дата снятия': 'DataSnyatya', 'Адрес': 'Adress',
                           'Прибор': 'Pribor', 'Счётчик': 'Pribor', 'Заводской номер': 'Zavodnumber',
                           'Показания начало': 'PokazBegin', 'Показания конец': 'PokazEnd', 'единицы измерения': 'Izmer',
                           'Сумма договора': 'Price', 'Оплата': 'Price', 'способ оплаты': 'Pay',
                           'банк': 'Pay', 'Заказчик': 'Client', 'Клиент': 'Client', 'Хозяин': 'Client', 'номер телефона': 'phone',
                           'Годность': 'GoodOrBad', 'ФИО': 'Client', 'DateCreate': 'DateCreate',  'Дата замены': 'DataZameny', 'Дата отправления': 'DataSend',
                           'Дата монтажа': 'Montaj', 'Дата поступления': 'DateEntrace', 'ссылка': 'LincDoc', 'Новая сумма': 'NewSum', 'Новый прибор': 'NewPribor', 'Батарея была': 'BataryBegin',
                           'Номер накладой': 'NumberNaklad', 'Батарея стала': 'BataryEnd'}

            db_request: str = " ".join(
                [f"{CollumnName[i]}," for i in frame])[:-1]
            self.ins = ("INSERT INTO work.clients (" + db_request + ')' +
                        ' VALUES (' + '%s,' * frame.shape[1])[:-1] + ')'
            for i in range(frame.shape[0]):
                print(self.ins)
                print(frame.iloc[i])
                self.cur.execute(self.ins, tuple(frame.iloc[i]))
            self.__con__.commit()
            self.cur.close()
            self.__con__.close()
            frame = None
            print('good')
            return 'good'
        except:
            print('bad')
            return 'bad'
