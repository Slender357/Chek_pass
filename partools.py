from python_rucaptcha.ImageCaptcha import ImageCaptcha
from config import CONFIG
from telebot import TeleBot
import apiclient
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials

import httplib2



def answer_captcha(captcha_file: str):
    """Решает капчу капчи.
    Принимает ссылку на картинку и отправляет на сервер рукапчи для решения
    :return: ответ сервера рукапча
    """
    image_captcha = ImageCaptcha(rucaptcha_key=CONFIG['RUCAPTCHA_KEY'])
    result = image_captcha.captcha_handler(captcha_file=captcha_file)
    return result


def quickstart_sheet(spreadsheet_id=CONFIG['spreadsheet_id'], credentials_file=CONFIG['credentials_file'],
                     type_conection='sheets',
                     version_conection='v4'):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        credentials_file,
        ['https://www.googleapis.com/auth/spreadsheets',
         'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build(type_conection, version_conection, http=httpAuth)
    return spreadsheet_id, service


def get_sheet_values(spreadsheet_id, service, sheet):
    sheet_values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=sheet + '!A2:Z9999999'
    ).execute()
    keys = sheet_values['values'][0]
    dic_values = []
    for k in sheet_values['values']:
        d = {}
        m = 0
        for i in keys:
            try:
                d.update({i: k[m]})
                m += 1
            except BaseException:
                d.update({i: ''})
                m += 1
        dic_values.append(d)
    dic_values.pop(0)
    return dic_values


def convert_tab(pars_tab):
    if pars_tab['day'] is None and pars_tab['night'] is None:
        body = [''] * 6
        body.append('ПРОПУСКОВ НЕТ')
        return body
    elif pars_tab['day'] is None:
        body = [''] * 3
        for i in pars_tab['night'].values():
            body.append(i)
        body.append('ДНЕВНОГО ПРОПУСКА НЕТ')
        return body
    elif pars_tab['night'] is None:
        body = []
        for i in pars_tab['day'].values():
            body.append(i)
        for i in range(3):
            body.append('')
        body.append('НОЧНОГО ПРОПУСКА НЕТ')
        return body
    else:
        body = []
        for i in pars_tab['day'].values():
            body.append(i)
        for i in pars_tab['night'].values():
            body.append(i)
        body.append('')
        return body


def chek_last_date(spreadsheet_id, service):
    date1 = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=CONFIG['sheet'] + '!J1'
    ).execute()['values'][0][0][25:]
    date = datetime.strptime(date1, '%d.%m.%Y') - datetime.strptime(datetime.strftime(datetime.today(), '%d.%m.%Y'),
                                                                    '%d.%m.%Y')
    if date == timedelta(days=0):
        sent_to_bot('Проверка', 'уже была сегодня')
        return False
    elif date + timedelta(days=1) == timedelta(days=0):
        sent_to_bot('Начата проверка ', datetime.strftime(datetime.today(), '%d.%m.%Y'))
        return True
    else:
        date2 = datetime.strftime(datetime.today(), '%d.%m.%Y')
        sent_to_bot(f'Старая дата проверки {date1}', f'Начата проверка {date2}')
        return True

def sent_to_bot(car, message):
    bot = TeleBot(CONFIG['telegram_bot_token'])
    bot.config['api_key'] = CONFIG['telegram_bot_token']
    bot.send_message(CONFIG['chat_id'], car+' '+message)


def safe_to_shet(car, worksheet, service, spreadsheet_id, index, pars_tab):
    body = convert_tab(pars_tab)
    worksheet.format(f"D{index}:F{index}", {
        "backgroundColor": {
            "red": 0.85,
            "green": 0.92,
            "blue": 0.83
        }
    }
                     )
    worksheet.format(f"G{index}:I{index}", {
        "backgroundColor": {
            "red": 0.79,
            "green": 0.85,
            "blue": 0.97
        }
    }
                     )
    if body[6] == 'ПРОПУСКОВ НЕТ':
        sent_to_bot(car, 'ПРОПУСКОВ НЕТ')
        worksheet.format(f"D{index}:I{index}", {
            "backgroundColor": {
                "red": 0.96,
                "green": 0.8,
                "blue": 0.8
            }
        }
                         )
        body[6] = ''
    elif body[6] == 'НОЧНОГО ПРОПУСКА НЕТ':
        sent_to_bot(car, 'НОЧНОГО ПРОПУСКА НЕТ')
        worksheet.format(f"G{index}:I{index}", {
            "backgroundColor": {
                "red": 0.96,
                "green": 0.8,
                "blue": 0.8
            }
        }
                         )
        body[6] = ''
    elif body[6] == 'ДНЕВНОГО ПРОПУСКА НЕТ':
        sent_to_bot(car, 'ДНЕВНОГО ПРОПУСКА НЕТ')
        worksheet.format(f"D{index}:F{index}", {
            "backgroundColor": {
                "red": 0.96,
                "green": 0.8,
                "blue": 0.8
            }
        }
                         )
        body[6] = ''
    if chek_date(body[1]):
        sent_to_bot(car, 'Дневной пропуск скоро закончится')
        worksheet.format(f"D{index}:F{index}", {
            "backgroundColor": {
                "red": 1.0,
                "green": 0.95,
                "blue": 0.80
            }
        })
    if chek_date(body[4]):
        sent_to_bot(car, 'Ночной пропуск скоро закончится')
        worksheet.format(f"G{index}:I{index}", {
            "backgroundColor": {
                "red": 1.0,
                "green": 0.95,
                "blue": 0.80
            }
        })
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=CONFIG['sheet'] + '!D' + str(index),
        valueInputOption='RAW',
        body={'values': [body]
              }
    ).execute()


def update_last_date(spreadsheet_id, service):
    date1=datetime.strftime(datetime.today(), '%d.%m.%Y')
    body = f'Дата последней проверки: {date1}'
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=CONFIG['sheet'] + '!J1',
        valueInputOption='RAW',
        body={'values': [[body]]
              }).execute()
    sent_to_bot('Проверка завершена ',date1)


def chek_date(fdate) -> bool:
    """
    Проверяет время до сегоднящней даты
    Если меньше двух недель до даты возвращает True
    Иначе False
    :param fdate:
    :return:
    """
    if fdate == '':
        return False
    fdate = datetime.strptime(fdate, '%d.%m.%Y') - timedelta(days=30)
    if datetime.today() > fdate:
        return True
    else:
        return False
