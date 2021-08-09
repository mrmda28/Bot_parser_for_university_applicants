import requests
from bs4 import BeautifulSoup
import datetime
import openpyxl
from pathlib import Path
from prettytable import PrettyTable

URL = 'https://nnov.hse.ru/bakvospo/abiturspo'
HEADERS = {'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) '
                         'Chrome/91.0.4472.114 Safari/537.36',
           'accept': '* / *'}


def get_xlsx():
    html = requests.get(URL, HEADERS)

    if html.status_code == 200:
        link = get_link(html)

        file = requests.get(link['href'])
        filename = f'files/{link["name"]} ({link["date"]}).xlsx'

        with open(filename, 'wb') as output:
            output.write(file.content)

        return filename
    else:
        print('Error: status_code ' + str(html.status_code))


def get_link(html):
    soup = BeautifulSoup(html.text, 'html.parser')
    items = soup.find_all('div', class_='wdj-plashka__card')

    for item in items:
        if item.findChild('h3', text='Программная инженерия (очно-заочная форма обучения)', recursive=False):
            a = item.findChild('a', recursive=True)
            name = a.contents[0]
            href = 'https://nnov.hse.ru' + str(a['href'])
            date = datetime.date.today().strftime("%d_%m")

            link = {'name': name,
                    'href': href,
                    'date': date}

            return link

    print('Error: link not found')


def parse_xlsx(path):
    xlsx = Path(path)
    wb = openpyxl.load_workbook(xlsx)

    sheet = wb.active

    rows = sheet.max_row

    list = []

    for row in sheet['B23':f'F{rows}']:
        abitur = []
        for cell in row:
            if cell.value is None:
                abitur.append(0)
            else:
                abitur.append(cell.value)
        list.append(abitur)

    def short_name(full_name):
        """u'Иванов Иван Петрович' -> u'Иванов И.П.'"""
        last, name, patronymic = full_name.split()
        return u'{last} {name[0]}.{patronymic[0]}.'.format(**vars())

    for item in list:
        item[0] = short_name(item[0])

    return list, path


def view_table(list, path):
    table = PrettyTable()

    table.field_names = ['Место', 'ФИО', 'Сумма', 'Математика', 'Информатика', 'Русский']

    sorted_list = sorted(list, key=lambda k: (k[1], k[2], k[3], k[4]), reverse=True)

    index = 0
    for item in sorted_list:
        index += 1
        table.add_row([index, item[0], item[1], item[2], item[3], item[4]])

    print(table.get_string(title=f'Количество: {len(sorted_list)}, {path[6:]}'))


xlsx_path = get_xlsx()

list, path = parse_xlsx(xlsx_path)

view_table(list, path)
