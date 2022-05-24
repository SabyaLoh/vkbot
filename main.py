import schedule
import time
from vk_api import VkApi
from threading import Thread
from confing import tokenvk, spreadsheetId, timet, namejson
from vkbottle.bot import Bot, Message
import httplib2
import apiclient
from oauth2client.service_account import ServiceAccountCredentials
vk_session = VkApi(token=tokenvk)
vk = vk_session.get_api()
bot = Bot(tokenvk)
import datetime

def main_loop():
    thread = Thread(target=do_schedule)
    thread.start()
def NAME():
    CREDENTIALS_FILE = namejson
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                                   ['https://www.googleapis.com/auth/spreadsheets',
                                                                    'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    ranges = ["Лист номер один!A1:I"]  #
    results = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheetId,
                                                       ranges=ranges,
                                                       valueRenderOption='FORMATTED_VALUE',
                                                       dateTimeRenderOption='FORMATTED_STRING').execute()
    sheet_values = results['valueRanges'][0]['values']
    i = 0
    for i in range(int(len(sheet_values))):
        if str(sheet_values[i][6]) == 'да':

            if str(sheet_values[i][8]) == 'нет':
                date = datetime.datetime.strptime(str(sheet_values[i][3]), "%d.%m.%Y")
                res = date + datetime.timedelta(days=int(sheet_values[i][4]))
                if res.date() == datetime.date.today():
                    results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body={
                        "valueInputOption": "USER_ENTERED",
                        # Данные воспринимаются, как вводимые пользователем (считается значение формул)
                        "data": [
                            {"range": f"I{i+1}",
                             "majorDimension": "ROWS",  # Сначала заполнять строки, затем столбцы
                             "values": [
                                 [f"{datetime.date.today()}"]
                             ]}
                        ]
                    }).execute()
                    results1 = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body={
                        "valueInputOption": "USER_ENTERED",
                        # Данные воспринимаются, как вводимые пользователем (считается значение формул)
                        "data": [
                            {"range": f"H{i + 1}",
                             "majorDimension": "ROWS",  # Сначала заполнять строки, затем столбцы
                             "values": [
                                 ["нет"]
                             ]}
                        ]
                    }).execute()
                    vk.messages.send(message = f'🔹Проект №{i}\n🔹Название проекта: {sheet_values[i][0]}\n🔹{sheet_values[i][2]}\n🔹{sheet_values[i][5]}',
                            chat_id = int(sheet_values[i][1]),
                                        random_id = 0)
            else:
                date = datetime.datetime.strptime(str(sheet_values[i][8]), "%Y-%m-%d")
                res = date + datetime.timedelta(days=int(sheet_values[i][4]))
                if res.date() == datetime.date.today():
                    results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body={
                        "valueInputOption": "USER_ENTERED",
                        # Данные воспринимаются, как вводимые пользователем (считается значение формул)
                        "data": [
                            {"range": f"I{i+1}",
                             "majorDimension": "ROWS",  # Сначала заполнять строки, затем столбцы
                             "values": [
                                 [f"{datetime.date.today()}"]
                             ]}
                        ]
                    }).execute()
                    results1 = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body={
                        "valueInputOption": "USER_ENTERED",
                        # Данные воспринимаются, как вводимые пользователем (считается значение формул)
                        "data": [
                            {"range": f"H{i+1}",
                             "majorDimension": "ROWS",  # Сначала заполнять строки, затем столбцы
                             "values": [
                                 ["нет"]
                             ]}
                        ]
                    }).execute()
                    vk.messages.send(
                        message=f'🔹Проект №{i}\n🔹Название проекта: {sheet_values[i][0]}\n🔹{sheet_values[i][2]}\n🔹{sheet_values[i][5]}',
                        chat_id=int(sheet_values[i][1]),
                        random_id=0)
        if str(sheet_values[i][7]) == 'нет':
            date = datetime.datetime.strptime(str(sheet_values[i][3]), "%d.%m.%Y")
            res = date + datetime.timedelta(days=int(sheet_values[i][4]))
            if res.date() > datetime.date.today():
                vk.messages.send(
                    message=f'🔹Проект №{i}\n🔹Название проекта: {sheet_values[i][0]}\n🔹{sheet_values[i][2]}\n🔹{sheet_values[i][5]}',
                    chat_id=int(sheet_values[i][1]),
                    random_id=0)
def do_schedule():
    schedule.every().day.at(timet).do(NAME)
    while True:
        schedule.run_pending()
        time.sleep(1)


@bot.on.chat_message(text="Какой id чата?")
async def id_handler(message: Message):
    users_info = await bot.api.users.get(message.from_id)
    await message.answer(f'{format(users_info[0].first_name)}, id чата - {message.chat_id}')

@bot.on.chat_message(text="Отчет готов<msg>")
async def hi_handler(message: Message):
    users_info = await bot.api.users.get(message.from_id)
    id = ((message.reply_message.text.split('\n'))[0].split('№'))[1]
    CREDENTIALS_FILE = namejson
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                                   ['https://www.googleapis.com/auth/spreadsheets',
                                                                    'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body={
        "valueInputOption": "USER_ENTERED",
        # Данные воспринимаются, как вводимые пользователем (считается значение формул)
        "data": [
            {"range": f"H{int(id)+1}",
             "majorDimension": "ROWS",  # Сначала заполнять строки, затем столбцы
             "values": [
                 ["да"]
             ]}
        ]
    }).execute()
    await message.answer("Ваш отчет принят, {}".format(users_info[0].first_name))
if __name__ == '__main__':
    main_loop()
    bot.run_forever()

