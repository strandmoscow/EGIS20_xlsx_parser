import os
import csv
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import Application, CommandHandler, ContextTypes
import logging
from gen_xlsx_egis1 import export_subsystem_to_excel_egis1
from gen_xlsx_egis2 import export_subsystem_to_excel_egis2
import time

# Настройка логирования
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)


# Открытие файла и получение API ключа
def get_api_key(filename: str):
    try:
        with open(filename, mode='r', newline='') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['Name'] == "tg":
                    return row['Key']
    except Exception as e:
        logger.error(f"Ошибка при чтении API ключа: {e}")
    return None



async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        'Привет! Я телеграм бот формирования отчетов для функций ЕГИС ОТБ. Используйте меню взаимодействия со мной!.'
    )


async def egis1_xlsx(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if len(context.args) != 1:
        await update.message.reply_text(
            "Пожалуйста, укажите подсистему, например:\n"
        )
        await update.message.reply_text(
            "/egis1_xlsx ПКП"
        )
        return

    subsystem_names = context.args[0]

    try:
        for subsystem_name in subsystem_names.split(', '):
            safe_name = subsystem_name.replace(' ', '_').replace('/', '_')
            output_path = f"{os.path.abspath(os.curdir)}/tg_bot_out/egis1"

            export_subsystem_to_excel_egis1(safe_name, output_path)
            await update.message.reply_document(open(output_path, 'rb'))

            # Удаление созданных файлов после отправки
            os.remove(f'{output_path}{subsystem_name}.xlsx')
    except Exception as e:
        logger.error(f"Ошибка при формировании или отправке файла: {e}")
        await update.message.reply_text("Ошибка при отправке файла.")


async def egis2_xlsx(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if len(context.args) != 1:
        await update.message.reply_text(
            "Пожалуйста, укажите подсистему, например:\n"
        )
        await update.message.reply_text(
            "/egis1_xlsx ПКП"
        )
        return

    subsystem_names = context.args[0]

    try:
        for subsystem_name in subsystem_names.split(', '):
            safe_name = subsystem_name.replace(' ', '_').replace('/', '_')
            output_path = f"{os.path.abspath(os.curdir)}/tg_bot_out/egis1"

            export_subsystem_to_excel_egis2(safe_name, output_path)
            await update.message.reply_document(open(output_path, 'rb'))

            # Удаление созданных файлов после отправки
            os.remove(f'{output_path}{subsystem_name}.xlsx')
    except Exception as e:
        logger.error(f"Ошибка при формировании или отправке файла: {e}")
        await update.message.reply_text("Ошибка при отправке файла.")


def main() -> None:
    api_key = get_api_key('creds/api_key.csv')

    if not api_key:
        logger.error("API ключ не найден.")
        return

    application = Application.builder().token(api_key).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("egis1_xlsx", egis1_xlsx))
    application.add_handler(CommandHandler("egis2_xlsx", egis2_xlsx))

    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()