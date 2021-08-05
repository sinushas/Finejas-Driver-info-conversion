import os
import sys
import openpyxl
import pandas as pd
import numpy as np
from datetime import datetime

from pandas.io import excel


timestamp = datetime.now().strftime("%Y-%m-%d")

# Atidaromas Excel failas
excel_data = pd.read_excel(r'C:\Users\saulius.sinusas\OneDrive - Finejas, UAB\Code\Finejas-Driver-info-conversion\duomenys\Book 2.xlsx', index_col=None)

# Pasalinami nedalyvaujantys asmenys
# excel_data = excel_data[excel_data['Komentaras'].isna()] ## kuriama visiems net nedalyvaujantiems, tad nebeaktualu

# pakeiciami stulpeliu pavadinimai is Rusu i Anglu kalba
excel_data['firstname'] = excel_data['Имя']
excel_data['lastname'] = excel_data['Фамилия']
excel_data['count'] = np.arange(len(excel_data)) + 1
# excel_data['company'] = excel_data['Компания'] ## nebereikalingi duomenys

# Sukuriami email'ai
excel_data['email'] = 'vairuotojas' + excel_data['count'].astype(str) + '@finejas.lt'

# Sukuriamas sysrole1 stulpelis
excel_data['sysrole1'] = 'student'

# Sukuriamas vartotojo vardas
# excel_data['lastname_2'] = excel_data['lastname'].astype(str).str[:2] ## paimdavo 2 pirmas pavard4s raides
excel_data['username'] = excel_data['Имя пользователя']

# Ismetami nereikalingi stulpeliai
# excel_data.drop(['Пароль', 'Имя пользователя', 'Personalas', 'Уровень', 'Имя', 'Фамилия', 'Компания', 'Komentaras'], axis=1, inplace=True)

# Pašalinami tarpai po ar prieš žodį
excel_data['username'].strip()

# stulpeliu perdarymas eiles tvarka
excel_data = excel_data[['username', 'firstname', 'lastname', 'email', 'sysrole1']]

# dokumento issaugojimas
excel_data.to_csv(r'C:\Users\saulius.sinusas\OneDrive - Finejas, UAB\Code\Finejas-Driver-info-conversion\paruosta lentele\naujokai' + timestamp + '.txt')
# print(excel_data)
