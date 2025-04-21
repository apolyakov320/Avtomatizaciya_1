import pandas as pd
from datetime import datetime, timedelta
import re


# открываем файл
df = pd.read_excel(r'C:\Users\Artem\Desktop\Python\Codes_RTK\Задание 1\Исходник.xlsx') # сырые строки (необработанные строки). Нужны для того, чтобы слеш \ не вызывал экранирование символов.

#установим необходимые фильтры по столбцу резолюция
filters_resol =(
    (df["Резолюция"].str.contains(r'Нет\s*ТВ'))|
    (df["Резолюция"].str.contains(r'отсутствует\s*тех\s*\.\s*возможность', case=False, regex=True, na=False))
    ) 

# Текущая дата и 120 дней назад
today = datetime.now()
days_ago = today - timedelta(days=120)

#установим необходимые фильтры по столбцу дата и время
filters_date = df['Дата и время создания задания/всп. задания'] >= days_ago

# Вставляем столбец 'Тип дома' на конкретную позицию (после столбца 'Описание задания')
position_type = df.columns.get_loc('Описание задания') + 1
df.insert(position_type, 'Тип дома', 'МКД')  # По умолчанию значение 'МКД'

# Установим необходимые фильтры по столбцу описание
filters_descr =(
    df['Описание задания'].str.contains(r'кв\s*\.\s*0', case=False, regex=True, na=False)|
    df['Описание задания'].str.contains(r'частный', case=False, regex=True, na=False) 
    )

# Применим фильтр для столбца Тип дома
df.loc[filters_descr, 'Тип дома'] = 'ЧС'

# Функция для извлечения адреса (с обработкой NaN и не-строк)
def extract_address(description):
    if pd.isna(description) or not isinstance(description, str): # Проверка на NaN и нестроковые значения:
        return None
    match = (re.search(r'Адрес въезда(.*?)Диапазон', description, re.DOTALL) or 
    re.search(r'Новый адрес(.*?)МРФ', description, re.DOTALL) or
    re.search(r'Новый адрес(.*?)Старый адрес', description, re.DOTALL)) # Режим re.DOTALL позволяет . в регулярном выражении захватывать переносы строк (\n), что полезно, если адрес разбит на несколько строк.
    return match.group(1).strip() if match else None
# Вставляем столбец 'Новый адрес' на конкретную позицию (после столбца 'Тип дома')
position_newAddress = df.columns.get_loc('Тип дома') + 1
df.insert(position_newAddress, 'Новый адрес', '')
# Применяем функцию к столбцу
df['Новый адрес'] = df['Описание задания'].apply(extract_address)

#выделение Старого адреса
def extract_old_address (description_old): 
    if pd.isna(description_old) or not isinstance (description_old, str):
        return None 
    
    match = re.search(r'Адрес выезда(.*?)Адрес въезда', description_old, re.DOTALL) or re.search(r'Старый адрес(.*?)- Основной телефон', description_old, re.DOTALL) 
    return match.group(1).strip() if match else None 

position_old_address = df.columns.get_loc("Тип дома")+2
df.insert(position_old_address, "Старый адрес", "")
df["Старый адрес"] = df["Описание задания"].apply(extract_old_address) 

#выделение телефона
def extract_phone (description_phone): 
    if pd.isna(description_phone) or not isinstance (description_phone, str):
        return None 
    
    match = re.search(r'Основной телефон для связи:(.*?)ID запуска скрипта', description_phone, re.DOTALL) or re.search(r'Контактный телефон:(.*?)Email', description_phone, re.DOTALL) 
    return match.group(1).strip() if match else None 

position_phone = df.columns.get_loc("Тип дома")+3
df.insert(position_phone, "Телефон", "")
df["Телефон"] = df["Описание задания"].apply(extract_phone) 

# Объединим фильтры в одну переменную
filters_applied = df[filters_resol & filters_date]
print(filters_applied)

#сохранить в excel
filters_applied.to_excel("отфильтрованные_данные.xlsx", index=False)
