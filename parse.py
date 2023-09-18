import pandas as pd
import numpy as np
from pandas.io.excel import ExcelWriter
from copy import copy
from operator import itemgetter

xlsx = pd.ExcelFile('Абитуриенты для аналитики.xlsx')
df = pd.read_excel(xlsx, 'TDSheet')

list_of_dicts = []

python_dict = {
    'name_of_person': str,
    'name_of_index': str,
    'home_city': str
}


# def get_num_of_applications(dicts: list[dict], string: str):
#     j = 0
#     for i in range(len(dicts)):
#         if dicts[i]['name_of_index'] == string:
#             j += 1
#     return j

def get_num_of_applications(dicts: list[dict], list_of_indices: list):
    m = 0
    indices_dict = {}
    for i in range(len(list_of_indices)):
        for j in range(len(dicts)):
            if dicts[j]['name_of_index'] == list_of_indices[i]:
                m += 1
        indices_dict[list_of_indices[i]] = m
        m = 0
    return indices_dict


list_of_napr = ['Прикладная математика и информатика', 'Управление в технических системах', 'Наноинженерия', 'Геология',
                'Архитектура', 'Нефтегазовое дело', 'Эксплуатация', 'Строительство',
                'Инноватика', 'Конструкторско-технологическое', 'Энергетическое машиностроение']


def get_names(dataframe: pd.core.frame.DataFrame):
    names_list = []
    for cell in df['Unnamed: 0'].items():
        if cell[1] != 'Фамилия, имя, отчество':
            names_list.append(cell)
    return names_list


# names_list = []
# for cell in df['Unnamed: 0'].items():
#     print(cell)


def get_home_city(dataframe: pd.core.frame.DataFrame):
    if df['Unnamed: 2'][i] != 'Адрес проживания':
        try:
            data = df['Unnamed: 2'][i].split(',')[1].strip()
            parsed_data = data.replace('область', 'обл').replace(
                'Республика', 'Респ', 1).replace('Респ.', 'Респ', 1).replace(
                'респ.', 'Респ', 1).replace('г Москва', 'Московская обл').replace('Москва', 'Московская обл').replace(
                'Москва г', 'Московская обл').replace('Московская область', 'Московская обл').replace(
                'АО. Ханты-Мансийский Автономный округ - Югра', 'Ханты-Мансийский Автономный округ - Югра АО').replace(
                'Москва', 'Москва г').replace('Респ.', 'Республика').replace('край.', '').replace('край', '')
            return parsed_data.strip()
        except IndexError:
            data = df['Unnamed: 2'][i].strip()
            parsed_data = data.replace('область', 'обл').replace(
                'Республика', 'Респ', 1).replace('Респ.', 'Респ', 1).replace(
                'респ.', 'Респ', 1).replace('г Москва', 'Московская обл').replace('Москва', 'Московская обл').replace(
                'Москва г', 'Московская обл').replace('Московская область', 'Московская обл').replace(
                'АО. Ханты-Мансийский Автономный округ - Югра', 'Ханты-Мансийский Автономный округ - Югра АО').replace(
                'Москва', 'Москва г').replace('Респ.', 'Республика').replace('край.', '').replace('край', '')
            return parsed_data.strip()
        except AttributeError:
            return np.nan


names_tuples = get_names(df)

for i in range(len(names_tuples)):
    python_dict['name_of_person'] = names_tuples[i][1]
    python_dict['home_city'] = get_home_city(df)
    if df['Unnamed: 1'][i][1] == 'К':
        python_dict['name_of_index'] = 'Прикладная математика и информатика'
    if df['Unnamed: 1'][i][1] == 'У':
        python_dict['name_of_index'] = 'Управление в технических системах'
    if df['Unnamed: 1'][i][1] == 'И':
        python_dict['name_of_index'] = 'Наноинженерия'
    if df['Unnamed: 1'][i][1] == 'Г':
        python_dict['name_of_index'] = 'Геология'
    if df['Unnamed: 1'][i][1] == 'А':
        python_dict['name_of_index'] = 'Архитектура'
    if df['Unnamed: 1'][i][1] == 'Н':
        python_dict['name_of_index'] = 'Нефтегазовое дело'
    if df['Unnamed: 1'][i][1] == 'Х':
        python_dict['name_of_index'] = 'Эксплуатация'
    if df['Unnamed: 1'][i][1] == 'С':
        python_dict['name_of_index'] = 'Строительство'
    if df['Unnamed: 1'][i][1] == 'О':
        python_dict['name_of_index'] = 'Инноватика'
    if df['Unnamed: 1'][i][1] == 'М':
        python_dict['name_of_index'] = 'Конструкторско-технологическое'
    if df['Unnamed: 1'][i][1] == 'Д':
        python_dict['name_of_index'] = 'Энергетическое машиностроение'
    list_of_dicts.append(copy(python_dict))
    python_dict.clear()

# print(list_of_dicts)
work_list = [list_of_dicts[i]['home_city']
             for i in range(len(list_of_dicts)) if list_of_dicts[i]['home_city']
             if list_of_dicts[i]['name_of_index'] == 'Энергетическое машиностроение']
work_list = np.array(work_list)

# unique, counts = np.unique(work_list, return_counts=True)
# x = np.asarray((unique, counts)).T


# print(len(np.unique(work_list)))            # количество уникальных значений

x = get_num_of_applications(list_of_dicts, list_of_napr)

sorted_amount_of_applications = dict(sorted(x.items(), key=itemgetter(1)))

# dataframe = pd.DataFrame(x)
dataframe = pd.DataFrame(list(x.items()), columns=[
                         'Направлений', 'Кол-во заявлений'])

with ExcelWriter('mydata.xlsx', mode='a') as writer:
    dataframe.to_excel(
        writer, sheet_name='Кол-во каждое направ', index=False)
# print(sorted_amount_of_applications)


# dataframe.to_excel(r'mydata.xlsx', index=False,
#                    sheet_name='Кол-во заявлений по напр')
