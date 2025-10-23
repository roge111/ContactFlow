import os
import pandas as pd
import csv
from django.shortcuts import render
from integration_utils.bitrix24.bitrix_user_auth.main_auth import main_auth

from django.http import HttpResponse, JsonResponse
from django.core.files.storage import FileSystemStorage
from django.views.decorators.csrf import csrf_exempt
import json
import datetime



def convert_to_csv(file_name: str):
    """
    Converts an Excel file to CSV format.
    
    This function reads an Excel file and converts it to CSV format,
    saving the result in the same directory.
    
    Args:
        file_name (str): Name of the Excel file to convert (should have .xlsx extension)
         
    Returns:
        None: Prints status messages to console
    """

    current_dir = os.getcwd()
    file_path = os.path.join(current_dir, 'main', 'static', file_name)
    output_path = os.path.join(current_dir,  'main', 'static', file_name.replace('xlsx', 'csv'))

    try:
        # Чтение Excel файла
        df = pd.read_excel(file_path, sheet_name=0)  # читаем первый лист
        
        # Сохранение в CSV
        df.to_csv(output_path, 
                index=False, 
                encoding='utf-8', 
                sep=',', 
                quoting=1)
        
        print(f"Файл успешно конвертирован в {output_path}")

    except UnicodeDecodeError:
        print("Ошибка декодирования. Попробуйте другую кодировку.")
    except FileNotFoundError:
        print(f"Файл не найден по пути: {file_path}")
    except Exception as e:
        print(f"Произошла ошибка: {e}")

def check_is_contact(token, name, last_name):
    """
    Проверяет наличие контакта в Bitrix24 по имени и фамилии.
    
    Аргументы:
        token: Токен авторизации Bitrix24
        name (str): Имя контакта
        last_name (str): Фамилия контакта
        
    Возвращает:
        tuple: (bool, int) - (True если контакт найден, ID контакта) или (False, -1)
    """
    contact = token.call_list_method('crm.contact.list', {
        'filter': {
            '=NAME': name,
            '=LAST_NAME': last_name
        }
        })
    if len(contact):
        return True, contact[0].get('ID')
    return False, -1

def check_company(token, company):
    """
    Проверяет наличие компании в Bitrix24 и создает новую, если она отсутствует.
    
    Эта функция ищет компанию по названию в Bitrix24 CRM. Если компания не найдена,
    создает новую компанию с указанным названием.
    
    Аргументы:
        token: Токен авторизации Bitrix24
        company (str): Название компании для поиска или создания
        
    Возвращает:
        int: ID компании (существующей или вновь созданной)
    """
    company_id = token.call_list_method('crm.company.list', {
        'filter': {
            '=TITLE': company
        }
        })
    
    
    if len(company_id):
        company_id = company_id[0].get('ID')
    else:
        new_company = token.call_list_method('crm.company.add',{
            'fields': {
                'TITLE': company
            }
        })
        
        
        company_id = new_company
    return company_id


def update_contact(token, name, last_name, number, email, company, id):
    """
    Обновляет информацию о контакте в Bitrix24.
    
    Эта функция обновляет данные существующего контакта в Bitrix24 CRM, включая
    имя, фамилию, телефон, email и компанию. Предварительно удаляет все старые
    телефонные номера контакта.
    
    Аргументы:
        token: Токен авторизации Bitrix24
        name (str): Имя контакта
        last_name (str): Фамилия контакта
        number (str): Номер телефона контакта
        email (str): Email контакта
        company (str): Название компании контакта
        id (int): ID контакта в Bitrix24
        
    Возвращает:
        dict: Результат API-запроса на обновление контакта
    """
    contact = token.call_list_method(
    'crm.contact.get', 
    {
        'id': id,
        'select': [ '*',
            'ID', 'NAME', 'LAST_NAME', 
            'EMAIL', 'PHONE', 'COMPANY_ID', 
            # Все необходимые поля
        ] })
    
    
    phones = contact.get('PHONE')
    # Удаляем все старые телефоны. Поддерживаем нормальный стадарт - на одного человека один телефон. Поскольку в таблице только одна колонка
    if phones != None:
        for phone in phones:
            token.call_list_method('crm.contact.update',{
                'ID':id,
                'fields': {
                    
                    'PHONE': [{'ID': phone.get('ID'), 'DELETE': 'Y'}]
                    
                }
            })
    #Проверяем компанию
    company_id = check_company(token, company)
   
    
    


    
    # Обновляем контакт, добавляя новые данные
    print('Почта', email)
    contact = token.call_api_method('crm.contact.update',{
        'ID':id,
        'fields': {
            'NAME': name,
            'LAST_NAME': last_name,
            'PHONE': [{'VALUE': number, 'VALUE_TYPE': 'WORK'}],
            'EMAIL': [{'VALUE': email, 'VALUE_TYPE': 'WORK'}],
            'COMPANY_ID': company_id
        }
    })
    
   
import pandas as pd

def save_to_excel(data, filename='contacts.xlsx'):
    """
    Сохраняет данные в файл формата Excel.
    
    Эта функция принимает список словарей с данными и сохраняет их в файл
    формата Excel (.xlsx) в директории static/data.
    
    Аргументы:
        data (list): Список словарей с данными для сохранения
        filename (str, optional): Имя файла для сохранения. По умолчанию 'contacts.xlsx'
        
    Возвращает:
        None
    """
    current_dir = os.getcwd()
    file_path = os.path.join(current_dir, 'main', 'static','data', filename)
    # Создаем DataFrame
    df = pd.DataFrame(data)
    
    # Сохраняем в Excel
    df.to_excel(file_path, index=False)



def save_to_csv(data, filename='contacts.csv'):
    """
    Сохраняет данные в файл формата CSV.
    
    Эта функция принимает список словарей с данными и сохраняет их в файл
    формата CSV в директории static/data.
    
    Аргументы:
        data (list): Список словарей с данными для сохранения
        filename (str, optional): Имя файла для сохранения. По умолчанию 'contacts.csv'
        
    Возвращает:
        None
    """
    current_dir = os.getcwd()
    file_path = os.path.join(current_dir, 'main', 'static','data', filename)

    headers = data[0].keys()
        
        # Открываем файл для записи
    with open(file_path, mode='w', encoding='utf-8', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=headers)
        
        # Записываем заголовок
        writer.writeheader()
        
        # Записываем данные
        writer.writerows(data)
           

def create_contact(token, name, last_name, number, email, company):
    """
    Создает новый контакт в Bitrix24.
    
    Эта функция создает новый контакт в Bitrix24 CRM с указанными данными.
    Перед созданием проверяет валидность email и наличие компании.
    
    Аргументы:
        token: Токен авторизации Bitrix24
        name (str): Имя контакта
        last_name (str): Фамилия контакта
        number (str): Номер телефона контакта
        email (str): Email контакта
        company (str): Название компании контакта
        
    Возвращает:
        dict: Результат API-запроса на создание контакта
        
    Выбрасывает:
        ValueError: Если email имеет некорректный формат
    """
    company_id = check_company(token, company)
    if not validate_email(email):
        raise ValueError("Некорректный формат email")
    print('Почта', email)
    contact = token.call_api_method('crm.contact.add',{
        'fields': {
            'NAME': name,
            'LAST_NAME': last_name,
            'PHONE': [{'VALUE': number, 'VALUE_TYPE': 'WORK'}],
            'EMAIL': [{'VALUE': email, 'VALUE_TYPE': 'WORK'}],
            'COMPANY_ID': company_id
        }
    })

def validate_email(email):
    """
    Проверяет валидность email адреса.
    
    Эта функция проверяет, соответствует ли переданный email адрес стандартному формату.
    
    Аргументы:
        email (str): Email адрес для проверки
        
    Возвращает:
        bool: True если email валидный, False в противном случае
    """
    print('Почта', email)
    import re
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

@main_auth(on_cookies=True)
def main(request):
    """
    Основная страница приложения.
    
    Эта функция отображает главную страницу приложения ContactFlow.
    Использует декоратор main_auth для аутентификации пользователя
    при запуске приложения и установке cookie.
    
    Аргументы:
        request: HTTP запрос Django
        
    Возвращает:
        HttpResponse: Рендер шаблона 'main/main.html'
    """
    return render(request, 'main/main.html')

@main_auth(on_cookies=True)
def import_page(request):
    """
    Страница импорта контактов.
    
    Эта функция отображает страницу для импорта контактов из файла.
    
    Аргументы:
        request: HTTP запрос Django
        
    Возвращает:
        HttpResponse: Рендер шаблона 'main/import.html'
    """
    return render(request, 'main/import.html')

@main_auth(on_cookies=True)
def export_page(request):
    """
    Страница экспорта контактов.
    
    Эта функция отображает страницу для экспорта контактов в файл.
    
    Аргументы:
        request: HTTP запрос Django
        
    Возвращает:
        HttpResponse: Рендер шаблона 'main/export.html'
    """
    return render(request, 'main/export.html')

@main_auth(on_cookies=True)
@csrf_exempt
def import_file(request):
    """
    Импортирует контакты из CSV файла в Bitrix24.
    
    Эта функция читает данные контактов из CSV файла, проверяет наличие каждого контакта
    в Bitrix24 CRM и либо обновляет существующий контакт, либо создает новый.
    
    Аргументы:
        request: HTTP запрос Django с токеном авторизации Bitrix24
        
    Возвращает:
        JsonResponse: Результат операции импорта
        
    Примечания:
        - Поддерживает автоматическое преобразование файлов .xlsx в .csv
        - Формат данных в CSV: имя, фамилия, номер телефона, email, компания
        - Для каждого контакта проверяется наличие компании и создается при необходимости
        - Email проходит валидацию перед созданием контакта
    """

    token = request.bitrix_user_token
    
    if request.method == 'POST' and request.FILES:
        uploaded_file = request.FILES['file']
        fs = FileSystemStorage()
        file_name = fs.save(uploaded_file.name, uploaded_file)
        file_path = fs.path(file_name)
        
        # Обработка файла
        if file_name.endswith('.xlsx'):
            convert_to_csv(file_name)
            file_path = file_path.replace('.xlsx', '.csv')
        
        processed_count = 0
        created_count = 0
        updated_count = 0
        
        try:
            with open(file_path, encoding='utf-8') as file:
                titles = file.readline()
                for contact in file:
                    try:
                        name, last_name, number, email, company = contact.replace(';', ',').replace('"', "").split(',')
                        is_contact, id = check_is_contact(token, name, last_name)
                        if is_contact:
                            update_contact(token, name, last_name, number, email, company, id)
                            updated_count += 1
                        else:
                            create_contact(token, name, last_name, number, email, company)
                            created_count += 1
                        
                        processed_count += 1
                    except Exception as e:
                        print(f"Ошибка при обработке контакта: {e}")
                        continue
            
            # Удаляем временный файл
            try:
                os.remove(file_path)
                if file_name.endswith('.xlsx'):
                    csv_path = file_path.replace('.xlsx', '.csv')
                    if os.path.exists(csv_path):
                        os.remove(csv_path)
            except:
                pass
            
            return JsonResponse({
                'status': 'success', 
                'message': f'Обработано контактов: {processed_count}',
                'created': created_count,
                'updated': updated_count
            })
            
        except Exception as e:
            return JsonResponse({'status': 'error', 'message': f'Ошибка при обработке файла: {str(e)}'}, status=500)
    
    return JsonResponse({'status': 'error', 'message': 'Неверный запрос'}, status=400)

@main_auth(on_cookies=True)
def export_file(request):
    """
    Экспортирует контакты из Bitrix24 в файл CSV или Excel с возможностью фильтрации.
    
    Эта функция получает список контактов из Bitrix24 CRM, применяет фильтрацию и сортировку,
    форматирует данные и сохраняет их в файл CSV или Excel в зависимости от настроек.
    
    Аргументы:
        request: HTTP запрос Django с токеном авторизации Bitrix24
        
    Возвращает:
        HttpResponse: Файл для скачивания
        
    Примечания:
        - Поддерживаемые типы фильтров: company_asc, company_desc, name_asc, name_desc,
          last_name_asc, last_name_desc, email_asc, email_desc, phone_asc, phone_desc
        - Количество контактов и тип файла получаются с фронтенда
        - Для каждого контакта экспортируются: имя, фамилия, email, телефон, компания
    """

    # Получаем параметры с фронта
    if request.method == 'POST':
        # Если запрос POST - берем параметры из формы
        type_filter = request.POST.get('sort_type', 'phone_asc')
        count = int(request.POST.get('count', 10))
        type_file = request.POST.get('file_type', 'csv')
    else:
        # Если GET - используем параметры по умолчанию или из GET-параметров
        type_filter = request.GET.get('sort_type', 'phone_asc')
        count = int(request.GET.get('count', 10))
        type_file = request.GET.get('file_type', 'csv')

    token = request.bitrix_user_token

    # Применяем фильтрацию в зависимости от типа
    if type_filter == 'email_asc':
        contacts = token.call_list_method('crm.contact.list', {'order': {'EMAIL': 'ASC'},'select': ['*', 'PHONE', 'EMAIL']})
    elif type_filter == 'email_desc':
        contacts = token.call_list_method('crm.contact.list', {'order': {'EMAIL': 'DESC'},'select': ['*', 'PHONE', 'EMAIL']})
    elif type_filter == 'phone_asc':
        contacts = token.call_list_method('crm.contact.list', {'order': {'PHONE': 'ASC'},'select': ['*', 'PHONE', 'EMAIL']})
    elif type_filter == 'phone_desc':
        contacts = token.call_list_method('crm.contact.list', {'order': {'PHONE': 'DESC'},'select': ['*', 'PHONE', 'EMAIL']})
    elif type_filter == 'company_asc':
        contacts = token.call_list_method('crm.contact.list', {'order': {'COMPANY_ID': 'ASC'},'select': ['*', 'PHONE', 'EMAIL']})
    elif type_filter == 'company_desc':
        contacts = token.call_list_method('crm.contact.list', {'order': {'COMPANY_ID': 'DESC'},'select': ['*', 'PHONE', 'EMAIL']})
    elif type_filter == 'name_asc':
        contacts = token.call_list_method('crm.contact.list', {'order': {'NAME': 'ASC'},'select': ['*', 'PHONE', 'EMAIL']})
    elif type_filter == 'name_desc':
        contacts = token.call_list_method('crm.contact.list', {'order': {'NAME': 'DESC'},'select': ['*', 'PHONE', 'EMAIL']})
    elif type_filter == 'last_name_asc':
        contacts = token.call_list_method('crm.contact.list', {'order': {'LAST_NAME': 'ASC'},'select': ['*', 'PHONE', 'EMAIL']})
    elif type_filter == 'last_name_desc':
        contacts = token.call_list_method('crm.contact.list', {'order': {'LAST_NAME': 'DESC'},'select': ['*', 'PHONE', 'EMAIL']})
    else:
        # По умолчанию - без сортировки
        contacts = token.call_list_method('crm.contact.list', {'select': ['*', 'PHONE', 'EMAIL']})
    

    contacts_resalt = []
    i = 0  # Начинаем с 0, так как список контактов начинается с первого элемента
    while i < len(contacts) and i < count:  # Исправлено условие
        print(f"Обработка контакта {i+1}")
        email = contacts[i].get('EMAIL', [])
        if len(email) > 0:
            email = email[0].get('VALUE', '')  # Выгружаю одну почту
        else:
            email = ''
        
        phone = contacts[i].get('PHONE', [])
        if len(phone) > 0:
            phone = phone[0].get('VALUE', '')  # Выгружаю один телефон
        else:
            phone = ''

        # Получаем компанию
        company_id = contacts[i].get('COMPANY_ID')
        company = ''
        if company_id:
            try:
                company_info = token.call_list_method('crm.company.get', {'id': company_id})
                company = company_info.get('TITLE', '')
            except Exception as e:
                print(f"Ошибка при получении компании: {e}")
                company = ''

        contact = {
            'имя': contacts[i].get('NAME', ''),
            'фамилия': contacts[i].get('LAST_NAME', ''),
            'email': email,
            'телефон': phone,
            'компания': company
        }
        contacts_resalt.append(contact)
        i += 1
    
    print(f"Экспортировано контактов: {len(contacts_resalt)}")
    
    # Создаем уникальное имя файла с timestamp
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    if type_file == 'csv':
        filename = f'contacts_export_{timestamp}.csv'
        save_to_csv(contacts_resalt, filename)
        
        # Возвращаем файл для скачивания
        current_dir = os.getcwd()
        file_path = os.path.join(current_dir, 'main', 'static', 'data', filename)
        
        with open(file_path, 'rb') as file:
            response = HttpResponse(file.read(), content_type='text/csv')
            response['Content-Disposition'] = f'attachment; filename="{filename}"'
            return response
    else:
        filename = f'contacts_export_{timestamp}.xlsx'
        save_to_excel(contacts_resalt, filename)
        
        # Возвращаем файл для скачивания
        current_dir = os.getcwd()
        file_path = os.path.join(current_dir, 'main', 'static', 'data', filename)
        
        with open(file_path, 'rb') as file:
            response = HttpResponse(file.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = f'attachment; filename="{filename}"'
            return response