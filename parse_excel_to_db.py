import os
import psycopg2
from psycopg2.extras import execute_values
from datetime import datetime

import pandas as pd

def init_conn():
    creds_df = pd.read_csv('creds/db_cred.csv')
    creds = creds_df.iloc[0]
    db_params = {
        'dbname': creds['DB_NAME'],
        'user': creds['DB_USER'],
        'password': creds['DB_PASSWORD'],
        'host': creds['DB_HOST'],
        'port': creds['DB_PORT']
    }
    conn = psycopg2.connect(**db_params)
    return conn

# --- Параметры ---
EXCEL_DIR = '/Users/andrejstrelcenko/Desktop/Работа/ЕГИС 2.0/ЕГИС 2.0/Целевые системы/ОКП/gen'
# EXCEL_DIR = '/Users/andrejstrelcenko/Desktop/strand.moscow/EGIS20_xlsx_parser/xlsx'

# Колонки для вставки в таблицу functions
FIELDS = [
    'development_year', 'document_source_id', 'tz_item',
    'tz_functionality', 'documented_functions', 'undocumented_functions',
    'critical_technologies', 'analysis_sources', 'included_in_maintenance',
    'demand_rating', 'relevance_rating', 'migrates_to_egis20',
    'implementation_priority', 'comments', 'function_code', 'egis_1_system_id', 'egis_2_system_id',
    'type_of_requirements', 'function_number'
]


def get_or_create_document_source(cur, name, number, date):
    if not name:
        return None
    cur.execute("""
        SELECT id FROM "egis2.0_test".document_sources
        WHERE document_name = %s AND document_number = %s
        """, (name, number))
    row = cur.fetchone()
    if row:
        return row[0]
    cur.execute("""
        INSERT INTO "egis2.0_test".document_sources (document_name, document_date, document_number)
        VALUES (%s, %s, %s) RETURNING id
        """, (name, date.strftime('%Y-%m-%d'), number))
    return cur.fetchone()[0]



def get_or_create_egis_1_system(cur, code, name):
    if not code:
        return None
    cur.execute("""
        SELECT id FROM "egis2.0_test".egis_1_systems
        WHERE system_code = %s
        """, (code,))
    row = cur.fetchone()
    if row:
        return row[0]
    cur.execute("""
        INSERT INTO "egis2.0_test".egis_1_systems (system_code, system_name)
        VALUES (%s, %s) RETURNING id
        """, (code, name))
    return cur.fetchone()[0]

def get_or_create_egis_2_system(cur, code, name):
    if not code:
        return None
    cur.execute("""
        SELECT id FROM "egis2.0_test".egis_2_systems
        WHERE system_code = %s
        """, (code,))
    row = cur.fetchone()
    if row:
        return row[0]
    cur.execute("""
        INSERT INTO "egis2.0_test".egis_2_systems (system_code, system_name)
        VALUES (%s, %s) RETURNING id
        """, (code, name))
    return cur.fetchone()[0]

def get_or_create_egis_2_component(cur, code, sys_id):
    if not code:
        return None
    cur.execute("""
        SELECT id FROM "egis2.0_test".egis_2_system_components
        WHERE component_code = %s AND egis_2_system_id = %s
        """, (code, sys_id))
    row = cur.fetchone()
    if row:
        return row[0]
    cur.execute("""
        INSERT INTO "egis2.0_test".egis_2_system_components (component_code, egis_2_system_id)
        VALUES (%s, %s) RETURNING id
        """, (code, sys_id))
    return cur.fetchone()[0]

def parse_bool(val):
    if pd.isna(val):
        return None
    val = str(val).strip().lower()
    if val in ('да', 'yes', 'true', '1'):
        return True
    if val in ('нет', 'no', 'false', '0'):
        return False
    return None

def process_dataframe(df, cur, egis1_code, egis1_name):
    # Получаем/создаем систему ЕГИС 1.0
    egis1_id = get_or_create_egis_1_system(cur, egis1_code, egis1_name)

    data_to_insert = []

    for idx, row in df.iterrows():
        doc_name = str(row.get('Источник', '')).strip() if not pd.isna(row.get('Источник', None)) else None
        doc_name = doc_name.replace("№ ", "")
        doc_name = doc_name.replace("№", "")
        doc_number = doc_name.split(' ')[1]
        doc_date = datetime.strptime(doc_name.split(' ')[3], "%d.%m.%Y")
        doc_id = get_or_create_document_source(cur, doc_name, doc_number, doc_date) if doc_name else None

        egis2_code = str(row.get('Новое наименование подсистемы', '')).strip() if not pd.isna(row.get('Новое наименование подсистемы', None)) else None
        egis2_name = egis2_code
        egis2_id = get_or_create_egis_2_system(cur, egis2_code, egis2_name) if egis2_code else None

        func_code = str(row.get('Код функции', '')).strip() if not pd.isna(row.get('Код функции', None)) else None
        if func_code:
            ft_or_nft = func_code.split('.')[3]
            func_code = '.'.join([func_code.split('.')[0], func_code.split('.')[1], func_code.split('.')[2],
                                 "ФТ" if func_code.split('.')[3] == "ФТ" else "НФТ", func_code.split('.')[4]])
            egis_2_component = func_code.split('.')[2]
            get_or_create_egis_2_component(cur, egis_2_component, egis2_id)



        # Сопоставляем поля с учетом возможных пропусков
        data_to_insert.append((
            int(row['Год разработки по ГК']) if not pd.isna(row['Год разработки по ГК']) else None,
            doc_id,
            row['Номер п. ТЗ'] if not pd.isna(row['Номер п. ТЗ']) else None,
            row['Функциональность в соответствии с ТЗ (скопировать из ТЗ)'] if not pd.isna(
                row['Функциональность в соответствии с ТЗ (скопировать из ТЗ)']) else None,
            row['Задокументированные функции (скопировать и сопоставить из Рабочей документации, документа "Описание программы" раздел 2 "Функциональное назначение")'] if not pd.isna(
                row['Задокументированные функции (скопировать и сопоставить из Рабочей документации, документа "Описание программы" раздел 2 "Функциональное назначение")']) else None,
            row['Незадокументированные функции (при наличии, на основе экспертных знаний сотрудников ФГУП "ЗащитаИнфоТранс" и заказчиков (ФОИВы - МТ, МВД, ФСБ, и т.д.)'] if not pd.isna(
                row['Незадокументированные функции (при наличии, на основе экспертных знаний сотрудников ФГУП "ЗащитаИнфоТранс" и заказчиков (ФОИВы - МТ, МВД, ФСБ, и т.д.)']) else None,
            row['Перечень используемых технологий требующих импортозамещения и/или улучшения информационной безопасности \n(например: Oracle, IBM mq, FTP и прочее)'] if not pd.isna(
                row['Перечень используемых технологий требующих импортозамещения и/или улучшения информационной безопасности \n(например: Oracle, IBM mq, FTP и прочее)']) else None,
            row['Исходные данные для проведения анализа\n(основное: указать документы с децимальными номерами, из которых копировали функции системы) \n'] if not pd.isna(
                row['Исходные данные для проведения анализа\n(основное: указать документы с децимальными номерами, из которых копировали функции системы) \n']) else None,
            parse_bool(row['Включены в ГК на эксплуатацию']),
            row['Востребованность системы\n(обьективная оценка / экспертная оценка от аналитиков)'] if not pd.isna(
                row['Востребованность системы\n(обьективная оценка / экспертная оценка от аналитиков)']) else None,
            row['Актуальность\n(экспертная оценка о необходимости функции от ДР)'] if not pd.isna(
                row['Актуальность\n(экспертная оценка о необходимости функции от ДР)']) else None,
            parse_bool(row['Уходит в ЕГИС 2.0\nДа/Нет']),
            row['Приоритет реализации (высокий / средний / низкий; влияет на приоритет переноса в ЕГИС 2.0 (внутренняя информация)'].lower().strip() if not pd.isna(
                row['Приоритет реализации (высокий / средний / низкий; влияет на приоритет переноса в ЕГИС 2.0 (внутренняя информация)']) else None,
            row['Комментарии'] if not pd.isna(row['Комментарии']) else None,
            func_code,
            egis1_id,
            egis2_id,
            None if func_code is None else func_code.split('.')[3],
            None if func_code is None else func_code.split('.')[4]
        ))

    if data_to_insert:
        sql = f"""
        INSERT INTO "egis2.0_test".functions
        ({', '.join(FIELDS)})
        VALUES %s
        """
        execute_values(cur, sql, data_to_insert)
        print(f"Прочитано строк: {len(data_to_insert)}")

    return len(data_to_insert)


def parse_all_files():
    conn = init_conn()
    conn.autocommit = False
    cur = conn.cursor()

    try:
        for root, dirs, files in os.walk(EXCEL_DIR):
            for file in files:
                if file.endswith(('.xlsx', '.xls')):
                    file_path = os.path.join(root, file)
                    print(f"Обрабатываю файл: {file_path}")

                    # Пример: из имени файла берем код и имя системы ЕГИС 1.0
                    # Можно доработать под ваши правила именования
                    base_name = os.path.splitext(file)[0]
                    base_name = base_name.split('_')[0]
                    egis1_code = base_name
                    egis1_name = base_name.replace('_', ' ')

                    try:
                        xls = pd.ExcelFile(file_path)
                        for sheet_name in xls.sheet_names:
                            if sheet_name not in ['Легенда', 'К_ПАВП (пример)']:
                                print(f"  Лист: {sheet_name}")
                                try:
                                    # Читаем с заголовком на 3-й строке (header=2)
                                    df = pd.read_excel(xls, sheet_name=sheet_name, header=2)
                                    df = df.reset_index(drop=True)

                                    # Удаляем пустые строки
                                    df = df[df[df.columns[1]].notna()]
                                    len_data_to_insert = process_dataframe(df, cur, egis1_code, egis1_name)
                                    conn.commit()
                                    print(f"Загружено строк: {len_data_to_insert}")
                                except Exception as e:
                                    print(f"    Ошибка чтения листа {sheet_name}: {e}")
                    except Exception as e:
                        print(f"Ошибка открытия файла {file_path}: {e}")

    finally:
        cur.close()
        conn.close()


if __name__ == "__main__":
    parse_all_files()
