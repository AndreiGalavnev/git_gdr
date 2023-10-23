#this script will download files and folders from google drive
#It will perform some basic url cleaning to work with different types of links
#The users will also have access to a help function
import pandas as pd
import wget,gdown
import argparse
from docx import Document
import streamlit as st


def parser_headers(doc):
    # Переменная для хранения содержимого под заголовком
    headers = []
    # Флаг, указывающий, что мы находимся под заголовком
    under_heading = False
    heading_list = ['Heading 2', 'Heading 3', 'Heading 1']
    # Итерируемся по параграфам в документе
    for paragraph in doc.paragraphs:
        # Проверяем, является ли стиль параграфа заголовком 
        if paragraph.style.name in heading_list:
            under_heading = True
        else:
            under_heading = False

        # Если мы находимся под заголовком, добавляем содержимое параграфа в список
        if under_heading:
            headers.append(paragraph.text)


    # Пространство имен
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    # Удаляем пустые строки из списка
    while "" in headers:
        headers.remove("")

    # Список для хранения данных
    data = []

    # Проходимся по заголовкам
    for one_header in headers:
        # Словарь для хранения данных под заголовком
        one_header_data = {'Заголовок': one_header, 'Текст': '', 'Таблицы': []}

        # Флаг, указывающий, что мы находимся под нужным заголовком
        under_one_header = False

        # Итерируемся по элементам документа
        for element in doc.element.body:
            
            # Если находим параграф с текстом заголовка
            if element.tag.endswith('p') and element.text == one_header:
                under_one_header = True
                continue

            # Если мы находимся под заголовком, обрабатываем элемент
            if under_one_header:
                # Если элемент является параграфом, добавляем текст к заголовку_data['Текст']
                if element.tag.endswith('p'):
                    one_header_data['Текст'] += element.text + '\n'
                # Если элемент является таблицей, добавляем таблицу к заголовку_data['Таблицы']
                elif element.tag.endswith('tbl'):
                    # Создаем пустую строку для хранения текста таблицы
                    tabl_text = ""

                    # Проходимся по строкам таблицы
                    for row in element.iterfind('.//w:tr', namespaces=ns):
                        # Проходимся по ячейкам в каждой строке
                        for cell in row.iterfind('.//w:tc', namespaces=ns):
                            # Получаем текст из каждой ячейки и добавляем его к тексту таблицы
                            for paragraph in cell.iterfind('.//w:p', namespaces=ns):
                                for text_element in paragraph.iterfind('.//w:t', namespaces=ns):
                                    tabl_text += text_element.text + " "

                    # Добавляем текст таблицы к заголовку_data['Таблицы']
                    one_header_data['Таблицы'].append(tabl_text)
        data.append(one_header_data)

    # Создаем датафрейм на основе списка данных
    df = pd.DataFrame(data)
    # Чистим от повторов в таблицах (оставляем последние, да-да это не элегантно, но вроде работает)
    df['Таблицы'] = df['Таблицы'].mask(df['Таблицы'].duplicated(keep='last'))
    
    # Сохранение в Excel (просто для проверки)
    #df.to_excel('data.xlsx', index=False)
    return df


    
#https://drive.google.com/drive/folders/1Yba_EYXdkP7WTR4fkGU2gfv7zZ_Ui-NO?usp=share_link

if __name__=="__main__":
    st.header('G-drive folders downloader')
    url = st.text_input("Enter URL", value="")
    button = st.button('Отправить запрос')
    if button:
    #if url.split('/')[-1] == '?usp=sharing':
    #    url= url.replace('?usp=sharing','')
    #gdown.download_folder(url) # for local saving
        file_urls = gdown.download_folder(url) # все равно сохраняется на диск надо решить

        for file_url in file_urls:
            doc = Document(file_url)
            # Выполняйте операции с файлом docx
            # Например, получение содержимого документа:
            dfdoc = parser_headers(doc)
            st.dataframe(dfdoc, width=1300, hide_index=True)
