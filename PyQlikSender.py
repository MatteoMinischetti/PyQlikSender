from win32com.client import Dispatch  # pywin32 module
import smtplib
from email.message import EmailMessage
import os.path
from datetime import datetime
import configparser
import logging

current_path = os.getcwd()
config = configparser.ConfigParser()
config.read_file(open(r'PyQlikSenderConf.txt'))
QLIKVIEW_DOCUMENT = config.get('configuration','QLIKVIEW_DOCUMENT')
SMTP_ADDRESS = config.get('configuration','SMTP_ADDRESS')
SMTP_PORT = config.get('configuration','SMTP_PORT')
EMAIL_FROM = config.get('configuration', 'EMAIL_FROM')
EMAIL_LOGIN = config.get('configuration', 'EMAIL_LOGIN')
EMAIL_PASSWORD = config.get('configuration', 'EMAIL_PASSWORD')
LOGFILE = config.get('configuration', 'logfile')
QVRELOAD = config.get('configuration', 'QVRELOAD')


class QlikView:
    def __init__(self):
        self.app = Dispatch('QlikTech.QlikView')

    def opendoc(self, docname, username, password):
        doc = self.app.OpenDoc(docname, username, password)
        return doc

    def reload(self,doc):
        doc.Reload()

    def clearall(self,doc):
        doc.ClearAll()

    def apply_field_filter(self, doc, filter_field, filter_value):
        doc.Fields(filter_field).Select(filter_value)

    def closedoc(self, doc):
        doc.CloseDoc()



def sendmail(email, subject,  name, surname, filename):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = EMAIL_FROM
    msg['To'] = email
    content = 'Dear ' + name + ' ' + surname + '  your report is attached here'
    msg.set_content(content)
    try:
        with open(filename, 'rb') as xls:
            msg.add_attachment(xls.read(), maintype='application', subtype='octet-stream', filename=xls.name)
    except Exception as e:
        print('cannot open attachment to send, error :' + str(e))
    try:
        with smtplib.SMTP_SSL(SMTP_ADDRESS, SMTP_PORT) as smtp:
            smtp.login(EMAIL_LOGIN, EMAIL_PASSWORD)
            smtp.send_message(msg)
    except Exception as e:
        logging.info('cannot send email to : ' + email + ' , error :' + str(e))


def manage_document(docname, current_path, tb_email):
    username = None
    password = None
    q = QlikView()
    version = q.app.QvVersion()
    print(version)
    doc = q.opendoc(docname, username, password)
    if QVRELOAD == 'Y':
        q.reload(doc)
    email_table = doc.GetSheetObject(tb_email)  # the object table containing email to send and filter to apply
    today_date = str(datetime.today().date()).replace('-', '_')
    rowiter = 0
    rows = email_table.GetRowCount()

    while rowiter < rows:
        title = email_table.GetCell(rowiter, 0).Text
        name = email_table.GetCell(rowiter, 1).Text
        surname = email_table.GetCell(rowiter, 2).Text
        email = email_table.GetCell(rowiter, 3).Text
        company = email_table.GetCell(rowiter, 4).Text
        filter_field = email_table.GetCell(rowiter, 5).Text
        filter_value = email_table.GetCell(rowiter, 6).Text
        tb1 = email_table.GetCell(rowiter, 7).Text
        print(title, name, surname, email, company)
        rowiter = rowiter + 1
        subject = "Qlikview report service for: " + company

        if rowiter > 1:
            filename = tb1 + '_' + name + '_' + surname + '_' + today_date + '.xls'
            q.clearall(docname)  # clear all filter
            doc.Fields(filter_field).Select(filter_value)  # apply filter
            try:
                chart = doc.GetSheetObject(tb1)
            except Exception as e:
                logging.info('error: ' + str(e))
            chart_path = current_path + '/' + filename
            chart.ExportBiff(chart_path)
            sendmail(email, subject, name, surname, filename)

    q.closedoc(doc)
    q.app.Quit()


if __name__ == '__main__':
    doc = current_path + '/' + QLIKVIEW_DOCUMENT
    logging.basicConfig(format='%(asctime)s - %(message)s', filename=LOGFILE, encoding='utf-8',
                        level=logging.DEBUG)
    tb_email = 'email'   # qlikview table containing email address and info
    manage_document(doc, current_path, tb_email)
