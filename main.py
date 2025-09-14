import urllib.parse
from datetime import datetime
import requests
import urllib3
import telebot
import xlsxwriter

urllib3.disable_warnings()

from config import BOT_TOKEN, CHAT_ID_SERVICE, CHAT_ID, TOKEN_FIRST, TOKEN_SECOND, ID_FIRST, ID_SECOND

bot = telebot.TeleBot(BOT_TOKEN)
chatId = CHAT_ID
chatId_service = CHAT_ID_SERVICE
tokenFIRST = TOKEN_FIRST
tokenSECOND = TOKEN_SECOND
idFirst = ID_FIRST
idSecond = ID_SECOND

import random

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64)...",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)...",
    "Mozilla/5.0 (X11; Linux x86_64)...",
]


def check_comment(order_id, market_id):
    # проверяем есть ли отзыв

    headers = {
        "Content-Type": "application/vnd.api+json",
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
        'origin': 'https://kaspi.kz',
        'priority': 'u=1, i',
        'referer': 'https://kaspi.kz/',
        'sec-ch-ua': '"Chromium";v="128", "Not;A=Brand";v="24", "Google Chrome";v="128"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"macOS"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-site',
        'user-agent': random.choice(USER_AGENTS)
    }

    params = {
        'id': 'all',
        'limit': '300',
        'days': '365',
        'isCommented': 'false',
        'orderCode': order_id,
        'filterByOrder': 'true',
    }

    response = requests.get(
        f'https://kaspi.kz/yml/creview/rest/misc/merchant/{market_id}/reviews/period',
        params=params,
        headers=headers,
    )
    if response.json()['data'] == []:
        return True
    else:
        return False


def create_exel(start_time, finish_time, token, market_name, id_market, page_number):
    url = "https://kaspi.kz/shop/api/v2/orders"
    params = {
        "page[number]": page_number,
        "page[size]": 99,
        "filter[orders][state]": "ARCHIVE",
        "filter[orders][creationDate][$ge]": start_time,
        "filter[orders][creationDate][$le]": finish_time,
        "filter[orders][status]": "COMPLETED",
        "filter[orders][signatureRequired]": "false",
        "include[orders]": "user"
    }

    headers = {
        "Content-Type": "application/vnd.api+json",
        "X-Auth-Token": token,
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
        'origin': 'https://kaspi.kz',
        'priority': 'u=1, i',
        'referer': 'https://kaspi.kz/',
        'sec-ch-ua': '"Chromium";v="128", "Not;A=Brand";v="24", "Google Chrome";v="128"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"macOS"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-site',
        'user-agent': random.choice(USER_AGENTS)

    }

    response = requests.get(url, params=params, headers=headers, verify=False)

    # Список для хранения ссылок
    links = []

    for i in response.json()['data']:

        if i['attributes']['totalPrice'] > 20000 and check_comment(order_id=i['attributes']['code'],
                                                                   market_id=id_market):
            id_order = i['id']
            url_order = f'https://kaspi.kz/shop/api/v2/orders/{id_order}/entries'
            response_order = requests.get(url_order, headers=headers)
            product_name = response_order.json()['data'][0]['attributes']['offer']['name']
            id_order2 = response_order.json()['data'][0]['id']  # айди заказа
            descrive_product_in_order = f'https://kaspi.kz/shop/api/v2/orderentries/{id_order2}/product'
            response_descride_product_in_order = requests.get(descrive_product_in_order, headers=headers)
            productCode = response_descride_product_in_order.json()['data']['attributes']['code']  # SKU для клиента

            order_code = i['attributes']['code']
            full_name = i['attributes']['customer']['lastName'] + ' ' + i['attributes']['customer']['firstName']
            phone = i['attributes']['customer']['cellPhone']
            link_to_review = f'https://kaspi.kz/shop/review/productreview?orderCode={order_code}&productCode={productCode}&rating=5'
            link_to_review = urllib.parse.quote(link_to_review, safe='')
            answer = (
                    'Здравствуйте, ' + full_name + '!\n\n'
                    + 'Меня зовут Евгений, менеджер ' + market_name + ' ' + 'с Kaspi магазина.'
                    + 'Недавно Вы сделали заказ'' ' + '*(' + product_name + ')*'
                    + ' и мы бы хотели предложить Вам получить кешбек в размере *2 000 тнг!*\n\n'
                    + 'Что необходимо сделать?\n\n'
                    + ' 1️⃣⁠ *По ссылке ниже написать отзыв о товаре с упоминанием ' + market_name + ' и 5🌟.*\n'
                    + ' 2️⃣⁠ ⁠*Написать отзыв в 2ГИС* \n'
                    + ' 3️⃣⁠ ⁠*Дождаться, пока отзыв опубликуют.*\n\n'
                    + 'После публикации отзыва напишите нам номер телефона и мы переведем *Вам 2 000 тнг!*\n\n'
                    + 'KASPI.KZ: ' + link_to_review + '\n\n'
                    + '2ГИС: link_to_2gis \n\n'
            )

            massage_to_client = f'https://api.whatsapp.com/send?phone={phone}&text={answer}'
            links.append(massage_to_client)
        if 19999 > i['attributes']['totalPrice'] > 2000 and check_comment(order_id=i['attributes']['code'],
                                                                          market_id=id_market):
            id_order = i['id']
            url_order = f'https://kaspi.kz/shop/api/v2/orders/{id_order}/entries'
            response_order = requests.get(url_order, headers=headers)
            product_name = response_order.json()['data'][0]['attributes']['offer']['name']
            id_order2 = response_order.json()['data'][0]['id']  # айди заказа
            descrive_product_in_order = f'https://kaspi.kz/shop/api/v2/orderentries/{id_order2}/product'
            response_descride_product_in_order = requests.get(descrive_product_in_order, headers=headers)
            productCode = response_descride_product_in_order.json()['data']['attributes']['code']  # SKU для клиента

            order_code = i['attributes']['code']
            full_name = i['attributes']['customer']['lastName'] + ' ' + i['attributes']['customer']['firstName']
            phone = i['attributes']['customer']['cellPhone']
            link_to_review = f'https://kaspi.kz/shop/review/productreview?orderCode={order_code}&productCode={productCode}&rating=5'
            link_to_review = urllib.parse.quote(link_to_review, safe='')
            answer = (
                    'Здравствуйте, ' + full_name + '!\n\n'
                    + 'Меня зовут Евгений, менеджер ' + market_name + ' ' + 'с Kaspi магазина.'
                    + 'Недавно Вы сделали заказ'' ' + '*(' + product_name + ')*'
                    + ' и мы бы хотели предложить Вам получить кешбек в размере *1 000 тнг!*\n\n'
                    + 'Что необходимо сделать?\n\n'
                    + ' 1️⃣⁠ *По ссылке ниже написать отзыв о товаре с упоминанием ' + market_name + ' и 5🌟.*\n'
                    + ' 2️⃣⁠ ⁠*Написать отзыв в 2ГИС* \n'
                    + ' 3️⃣⁠ ⁠*Дождаться, пока отзыв опубликуют.*\n\n'

                    + 'После публикации отзыва напишите нам номер телефона и мы переведем *Вам 1 000 тнг!*\n\n'
                    + 'KASPI.KZ: ' + link_to_review + '\n\n'
                    + '2ГИС: link_to_2gis' + '\n\n'
            )

            massage_to_client = f'https://api.whatsapp.com/send?phone={phone}&text={answer}'
            links.append(massage_to_client)
            print(order_code)
    # Создание Excel файла
    excel_filename = f'{market_name}+{page_number}.xlsx'
    workbook = xlsxwriter.Workbook(excel_filename)
    worksheet = workbook.add_worksheet()

    # Запись ссылок в Excel
    for row_num, link in enumerate(links):
        worksheet.write(row_num, 0, link)

    workbook.close()
    # Отправка файла в Telegram
    with open(excel_filename, 'rb') as file:

        bot.send_document(chatId_service, file, timeout=120)


def main(startDate, finishDate):
    start_time = int(datetime.strptime(f'{startDate} 01:01:00', '%Y-%m-%d %H:%M:%S').timestamp() * 1000)
    finish_time = int(datetime.strptime(f'{finishDate} 23:59:00', '%Y-%m-%d %H:%M:%S').timestamp() * 1000)
    tookens_and_names = [
        [tokenFIRST, 'Название магазина', idFirst],
        [tokenSECOND, 'Название магазина', idSecond]

    ]

    for i in tookens_and_names:
        for j in range(4):
            create_exel(start_time=start_time, finish_time=finish_time, token=i[0], market_name=i[1], id_market=i[2],
                        page_number=j)


if __name__ == "__main__":
    startDate = input('Введите стартовую дату, по формату 2025-06-30, строго по формату: \n')
    finishDate = input('Введите стартовую дату, по формату 2025-06-30, строго по формату: \n')
    main(startDate, finishDate)
    input("\nНажмите Enter для выхода...")
