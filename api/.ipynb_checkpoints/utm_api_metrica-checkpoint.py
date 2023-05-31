# Импорт библиотек
import requests
import json
import pandas as pd
from io import StringIO
import gspread
from df2gspread import df2gspread as d2g
from oauth2client.service_account import ServiceAccountCredentials
import datetime


# ЯНДЕКС МЕТРИКА

# Передаю токен и авторизовываюсь
token = "y0_AgAEA7qkXqBwAAnr2AAAAADjYIRF_1qCnRGeTBWOBogBvzjGpd3erW4"
headers = {"Authorization": "OAuth " + token}

# Получение текущей даты и вычисление начальной и конечной дат для предыдущей недели
today = datetime.date.today()
end_date = today - datetime.timedelta(days=today.weekday()) - datetime.timedelta(days=1)
start_date = end_date - datetime.timedelta(days=6)

# Количество недель для отчета
weeks_to_report = 1

for i in range(weeks_to_report):
    params = {
        "ids": "50742883",
        "metrics": "ym:s:goal254504607reaches",
        "dimensions":  "ym:s:<attribution>UTMSource, \
                        ym:s:<attribution>UTMContent, \
                        ym:s:<attribution>UTMCampaign, \
                        ym:s:<attribution>UTMTerm",
        "date1": start_date.strftime("%Y-%m-%d"),
        "date2": end_date.strftime("%Y-%m-%d")
    }

    url = "https://api-metrika.yandex.net/stat/v1/data.csv?"
    response = requests.get(url, params=params, headers=headers)

    if response.status_code == 200:
        data = response.text
        df = pd.read_csv(StringIO(data))
        df = df.replace('"', "", regex=True) \
            .fillna("") \
            .rename(columns={
                "Goals reached (Оформление подписки)": 'Количество', 
                "Date interval of visit": "Week"
            })
        
        # Фильтрация данных по нужным UTM-источникам
        filtered_df = df.loc[
            (df['UTM Source'].isin([
                'podcast',
                'vk',
                'youtube',
                'yt',
                'zen',
                'tg',
                'blog',
                'email',
                'inst']))
        ]

    else:
        print("Error:", response.status_code)
        
        
# ГУГЛ ТАБЛИЦЫ

# Даю необходимые разрешения
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

my_mail = 'merabjejeia@gmail.com'

# Авторизовываюсь
credentials = ServiceAccountCredentials.from_json_keyfile_name('D:/Важное/API/credentials.json', scope)
gs = gspread.authorize(credentials)

# Передаю название таблицы. Также можно URL-адресом для open_by_url
# или id для open_by_key
table_name = 'test'

# Получаю таблицу
work_sheet = gs.open(table_name)

# Создание нового листа
date_range = f"{start_date.strftime('%Y.%m.%d')}-{end_date.strftime('%d')}"
worksheet = work_sheet.add_worksheet(title='подписки ' + date_range, rows="100", cols="20")


# Загрузка датафрейма в лист
import gspread_dataframe as gd
gd.set_with_dataframe(worksheet, filtered_df)

# Добавление фильтра к столбцам
worksheet.set_basic_filter()

# Сохранение изменений
header_range = f"A1:{chr(ord('A') + len(filtered_df.columns) - 1)}1"
header_cells = worksheet.range(header_range)
for cell, header in zip(header_cells, filtered_df.columns):
    cell.value = header
worksheet.update_cells(header_cells)