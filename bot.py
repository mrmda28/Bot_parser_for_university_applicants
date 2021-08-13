import openpyxl
import requests
from bs4 import BeautifulSoup
from pathlib import Path
import config
import logging
import asyncio
import datetime
from aiogram import Bot, Dispatcher, executor, types


URL = 'https://nnov.hse.ru/bakvospo/abiturspo'
HEADERS = {'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) '
                         'Chrome/91.0.4472.114 Safari/537.36',
           'accept': '* / *'}

logging.basicConfig(level=logging.INFO)

bot = Bot(token=config.API_TOKEN)
dp = Dispatcher(bot)


async def scheduled(wait_for):
    while True:
        await asyncio.sleep(wait_for)

        html = requests.get(URL, HEADERS)
        last_id = open('files/id.txt', 'r').read()

        if html.status_code == 200:
            soup = BeautifulSoup(html.text, 'html.parser')
            items = soup.find_all('div', class_='wdj-plashka__card')

            for item in items:
                if item.findChild('h3', text='Программная инженерия (очно-заочная форма обучения)', recursive=False):
                    a = item.findChild('a', recursive=True)
                    new_id = a['href'][19:-5]

            if last_id != new_id:
                with open('files/id.txt', 'w') as f:
                    f.write(new_id)

                name = a.contents[0]
                href = 'https://nnov.hse.ru' + str(a['href'])
                date_t = datetime.date.today().strftime("%d_%m")

                link = {'name': name,
                        'href': href,
                        'date': date_t}

                file = requests.get(link['href'])
                filename = f'files/{link["name"]} ({link["date"]}).xlsx'

                with open(filename, 'wb') as output:
                    output.write(file.content)

                xlsx = Path(filename)
                wb = openpyxl.load_workbook(xlsx)

                sheet = wb.active

                rows = sheet.max_row

                new_list = []

                for row in sheet['B23':f'F{rows}']:
                    abitur = []
                    for cell in row:
                        if cell.value is None:
                            abitur.append(0)
                        else:
                            abitur.append(cell.value)
                    new_list.append(abitur)

                def short_name(full_name):
                    last, name, patronymic = full_name.split()
                    return u'{last} {name[0]}.{patronymic[0]}.'.format(**vars())

                for item in new_list:
                    item[0] = short_name(item[0])

                sorted_list = sorted(new_list, key=lambda k: (k[1], k[2], k[3], k[4]), reverse=True)

                datee = f'Список на {filename[6:11]}\n'
                header = f'Количество: {len(sorted_list)}\n'

                content = ''
                index = 0
                for item in sorted_list:
                    index += 1
                    content += f'{index} - {item[0]} - {item[1]} ({item[2]}, {item[3]}, {item[4]})\n'

                text = datee + header + content

                for user in config.USER_ID:
                    await bot.send_message(user, text)


if __name__ == '__main__':
    loop = asyncio.get_event_loop()
    loop.create_task(scheduled(config.TIME))
    executor.start_polling(dp, skip_updates=True)
