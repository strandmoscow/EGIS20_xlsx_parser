import pandas as pd
from sqlalchemy import text

from utils import init_conn, create_xlsx


def export_subsystem_to_excel_egis1(sys_code, output_dir):
    # Подключение к БД
    engine = init_conn()

    # Получаем описание и код подсистемы
    with engine.connect() as conn:
        query = text('SELECT description FROM "egis2.0_test".egis_1_systems WHERE system_code = :sys_code')
        result = conn.execute(query, {'sys_code': sys_code}).fetchone()

    if result is None:
        print(f"Подсистема '{sys_code}' не найдена в справочнике.")
        raise Exception(f"Подсистема '{sys_code}' не найдена в справочнике.")
    sys_desc = result[0]

    # Запрос данных функций через функцию в БД
    query = text('SELECT * FROM "egis2.0_test".get_egis1_subsystem_report(:sys_name)')
    df = pd.read_sql(query, engine, params={'sys_name': sys_code})

    if df.empty:
        print(f"Данные для подсистемы '{sys_code}' не найдены.")
        return

    output_file = f'{output_dir}{sys_code}.xlsx'
    create_xlsx(output_file, sys_code, sys_desc, df, 1)

    print(f"Экспорт данных подсистемы '{sys_code}' завершён. Файл: {output_file}")
