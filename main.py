import requests
import yaml
from bs4 import BeautifulSoup
import os
import arrow as Aw
import csv
import xlrd
from queue import Queue


def main():
    global cfg
    session = requests.session()
    headers = {
        'User-Agent': cfg['options']['user-agent']
    }
    payload = {
        'LOGIN': cfg['account']['user'],
        'PASSWD': cfg['account']['pass']
    }

    session.post("https://finecobank.com/portalelogin", data=payload, headers=headers)
    print("Logged in")
    session.get("https://finecobank.com/home/myfineco", headers=headers)
    print("Requested Homapage")
    session.get("https://finecobank.com/conto-e-carte/movimenti/movimenti-conto", headers=headers)
    print("Request transaction page")

    last_date = Aw.get(date())
    current_date = Aw.utcnow()
    to_date = last_date.shift(months=+2)

    while last_date.timestamp <= current_date.timestamp:
        print(str(last_date.format('DD/MM/YYYY')))
        print(str(to_date.format('DD/MM/YYYY')))

        date_payload = {
            'filtroAperto': 'true',
            'keyword': '',
            'tipoImporto': 'tutti',
            'dataDal': str(last_date.format('DD/MM/YYYY')),
            'dataAl': str(to_date.format('DD/MM/YYYY'))
        }

        page = session.post("https://finecobank.com/conto-e-carte/movimenti/movimenti-conto/ricerca-filtro",
                            data=date_payload, headers=headers)
        soup = BeautifulSoup(page.text, 'html.parser')
        transactions = soup.find(name='span', attrs='txt14 bold')

        if not transactions:
            print("No new transaction!")
            break
        else:
            print("Number of transactions: " + transactions.text)
            xls = session.get("https://finecobank.com/conto-e-carte/movimenti/movimenti-conto/excel")
            convert_csv(xls.content)
            if to_date.shift(months=+2, days=+1).timestamp >= current_date.timestamp:
                last_date = to_date.shift(days=+1)
                to_date = current_date
            else:
                last_date = last_date.shift(months=+2, days=+1)
                to_date = to_date.shift(months=+2, days=+1)

    write("data.csv")


def convert_csv(xls):
    global header
    wb = xlrd.open_workbook(file_contents=xls)
    sheet = wb.sheet_by_index(0)
    header = [cell.value for cell in sheet.row(7)]
    for row_idx in reversed(range(8, sheet.nrows)):
        row = [cell.value for cell in sheet.row(row_idx)]
        row[0] = (Aw.get(xlrd.xldate_as_datetime(row[0], 0))).format(cfg['options']['data_format'])
        row[1] = (Aw.get(xlrd.xldate_as_datetime(row[1], 0))).format(cfg['options']['data_format'])
        q.put(row)


def write(filename):
    # TODO: Make a diff of the file.
    global header
    with open(filename, "w") as file:
        writer = csv.writer(file, delimiter=",")
        writer.writerow(header)
        while not q.empty():
            writer.writerow(q.get())
            q.task_done()


def date():
    global cfg
    if os.path.isfile('latest-check'):
        with open('latest-check', 'r+') as check_file:
            check = check_file.read()
            check_file.seek(0)
            check_file.write(str(Aw.utcnow().timestamp))
            check_file.truncate()
            check_file.close()
            return int(check)
    else:
        with open('latest-check', 'w') as check_file:
            check_file.write(str(Aw.utcnow().timestamp))
            check_file.close()
            return int(cfg['options']['oldest_date'])


def config():
    try:
        with open('config.yml', 'r', encoding='utf-8') as config_file:
            conf = yaml.load(config_file, Loader=yaml.FullLoader)
        return conf
    except Exception as e:
        print("Error: " + str(e))


if __name__ == "__main__":
    cfg = config()
    q = Queue(maxsize=0)
    main()
