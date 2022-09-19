from docxtpl import DocxTemplate
from dateutil.parser import parse
import pathlib
import uuid


class docxWritter():
    def __init__(self):
        self.month = {'01': 'января', '02': 'февраля',
                      '03': 'марта', '04': 'апреля', '05': 'мая',
                      '06': 'июня', '07': 'июля', '08': 'августа',
                      '09': 'сентября', '10': 'октября',
                      '11': 'ноября', '12': 'декабря', '00': 'Ошибка'}
        self.izm = 0

    def pechat(self, data):

        try:
            data['DataSnyatya'] = parse(data['DataSnyatya'])
            data['DataSnyatya'] = data['DataSnyatya'].strftime('%d.%m.%Y')
        except BaseException:
            data['DataSnyatya'] = '000000 Пустая дата'
        if data['dogovor'] == 'МВтч':
            if data['Izmer'] == 'кВтч':
                try:
                    self.izm = float(data['PokazEnd'])*0.00085984523
                except Exception:
                    pass
            elif data['Izmer'] == 'МВтч':
                try:
                    self.izm = float(data['PokazEnd'])*0.8598452278589853
                except Exception:
                    pass
            doc = DocxTemplate("МВтч.docx")
            context = {'day': data['DataSnyatya'][0:2], 'date': self.month[data['DataSnyatya'][3:5]],
                       'year': data['DataSnyatya'][6:10]+' г.', 'pribor': data['Pribor'],
                       'zavodNumber': data['Zavodnumber'], 'GoodorBad': data['GoodOrBad'],
                       'pBegin': data['PokazBegin'], 'pEnd': data['PokazEnd'],
                       'pEndGkall': float('{:.4f}'.format(self.izm)),
                       'dataSny': data['DataSnyatya'], 'client': data['Client'], 'dataMontaj': data['Montaj'], 'adress': data['address'], 'Izmer': data['Izmer']}

        elif data['dogovor'] == 'Гкал':
            doc = DocxTemplate("Гкал.docx")

            context = {'day': data['DataSnyatya'][0:2], 'date': self.month[data['DataSnyatya'][3:5]],
                       'year': data['DataSnyatya'][6:10]+' г.', 'pribor': data['Pribor'],
                       'zavodNumber': data['Zavodnumber'], 'GoodorBad': data['GoodOrBad'],
                       'pBegin': data['PokazBegin'], 'pEnd': data['PokazEnd'], 'dataSny': data['DataSnyatya'],
                       'client': data['Client'], 'dataMontaj': data['Montaj'], 'adress': data['address'], 'Izmer': data['Izmer']}
        doc.render(context)
        desktop_path = '/home/Sanya/filesDocx/'

        file_id = str(uuid.uuid1())
        dcs = desktop_path + file_id + ".docx"
        print(dcs)
        full_path_for_docx = desktop_path + "/" + \
            data['Client'] + "_" + data['dogovor'] + ".docx"
        print(full_path_for_docx)
        doc.save(dcs)

        filename = file_id + ".docx"
        return ('good', file_id + ".docx")
