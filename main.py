import os
import sys
import openpyxl
import pandas as pd
from datetime import datetime


timestamp = datetime.now().strftime("%Y-%m-%d")

# Atidaromas Excel failas
excel_data = pd.read_excel(r'C:\Users\sinus\OneDrive\Documents\GitHub\Finejas-Driver-info-conversion\duomenys\Book 2.xlsx')

# Pasalinami nedalyvaujantys asmenys
excel_data = excel_data[excel_data['Komentaras'].isna()]

# pakeiciami stulpeliu pavadinimai is Rusu i Anglu kalba
excel_data['firstname'] = excel_data['Имя']
excel_data['lastname'] = excel_data['Фамилия']
excel_data['company'] = excel_data['Компания']

# Sukuriami email'ai
excel_data['email'] = excel_data['firstname'] + '.' + excel_data['lastname'] + '@finejas.lt'

# Sukuriamas sysrole1 stulpelis
excel_data['sysrole1'] = 'student'

# Sukuriamas vartotojo vardas
excel_data['lastname_2'] = excel_data['lastname'].astype(str).str[:2]
excel_data['username'] = excel_data['firstname'] + excel_data['lastname_2']

# Ismetami nereikalingi stulpeliai
excel_data.drop(['Пароль', 'Имя пользователя', 'Personalas', 'Уровень', 'Имя', 'Фамилия', 'Компания', 'Komentaras'], axis=1, inplace=True)


# stulpeliu perdarymas eiles tvarka
excel_data = excel_data[['username', 'firstname', 'lastname', 'email', 'sysrole1']]

# dokumento issaugojimas
excel_data.to_csv(r'C:\Users\sinus\OneDrive\Documents\GitHub\Finejas-Driver-info-conversion\paruosta lentele\naujokai' + timestamp + '.txt')
print(excel_data)
