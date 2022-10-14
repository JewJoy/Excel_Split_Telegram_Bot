import telebot
from telebot import types
import logging
import datetime
import os
import openpyxl

# импорт файла с данными о токене и id администратора телеграм бота
import info


token = info.token
bot = telebot.TeleBot(token)
admin = info.admin

# автоматическое логгирование
logger = telebot.logger
telebot.logger.setLevel(logging.INFO)


# функция начала работы с ботом, открывается первичная информация с предоставлением дальнейших действий
@bot.message_handler(commands=['start'])
def start(message):
    try:
        but = types.ReplyKeyboardMarkup(resize_keyboard=True)
        but.add(
            types.KeyboardButton('Отправить файл')
        )
        bot.send_message(message.chat.id, 'Для отправки файла нажмите на кнопку в клавиатуре "Отправить файл"'
                                          ' или воспользуйтесь командой - [ /file ]', reply_markup=but)
    except Exception as e:
        print(f"def - start {e}")


# Команда к началу работы основного функционала системы
# Запрос количества строк
@bot.message_handler(commands=['file'])
def input_file(message):
    try:
        bot.send_message(message.chat.id, 'Перед тем как отправить файл, отправьте некоторые параметры, а именно: ')
        msg = bot.send_photo(message.chat.id,
                             open(f'C:\\Users\\dmi3\\PycharmProjects\\'
                                  f'Python_Excel_Split_Telegram_bot\\photo_ex_spl_bot\\row.png', 'rb'),
                             caption='Укажите количество строк в исходном файле\n'
                                     'Отправьте только число\n'
                                     '(Пример на фото выделен желтым)')

        # Ожидание ответа пользователя и вызов последующей функции
        bot.register_next_step_handler(msg, input_file_2)
    except Exception as e:
        print(f'{e}\ndef input_file(message)\n')


# Запрос конечного столбца
def input_file_2(message):
    try:
        max_line = message.text

        msg = bot.send_photo(message.chat.id,
                             open(f'C:\\Users\\dmi3\\PycharmProjects\\'
                                  f'Python_Excel_Split_Telegram_bot\\photo_ex_spl_bot\\column.png', 'rb'),
                             caption='Укажите конечный столбец\n'
                                     'Отправьте букву столбца\n'
                                     '(Пример на фото выделен желтым)')
        bot.register_next_step_handler(msg, input_file_3, max_line=max_line)
    except Exception as e:
        print(f'{e}\ndef input_file_2(message)\n')


# Запрос числа строк шапки
def input_file_3(message, max_line=None):
    try:
        index_column = message.text

        msg = bot.send_photo(message.chat.id,
                             open(f'C:\\Users\\dmi3\\PycharmProjects\\'
                                  f'Python_Excel_Split_Telegram_bot\\photo_ex_spl_bot\\head.png', 'rb'),
                             caption='Укажите число строк шапки\n(пример шапки на фото выделен желтым)')
        bot.register_next_step_handler(msg, input_file_4, max_line=max_line, index_column=index_column)
    except Exception as e:
        print(f'{e}\ndef input_file_3(message, max_line=None)\n')


# Запрос исходного файла
def input_file_4(message, max_line=None, index_column=None):
    try:
        head_line = message.text

        msg = bot.send_message(message.chat.id, 'Теперь отправьте Ваш файл')
        bot.register_next_step_handler(msg,
                                       input_file_5,
                                       max_line=max_line,
                                       index_column=index_column,
                                       head_line=head_line)
    except Exception as e:
        print(f'{e}\ndef input_file_4(message, max_line=None, max_line=None)\n')


# Функция обработки входных данных, формирование выходных файлов
def input_file_5(message, max_line=None, index_column=None, head_line=None):
    try:
        if message.document is not None:    # ! Добавить проверку на формат файла !
            down_doc_file = bot.download_file(bot.get_file(message.document.file_id).file_path)

            with open(f'C:\\Users\\dmi3\\PycharmProjects\\'
                      f'Python_Excel_Split_Telegram_bot\\excel_file\\{0}.xlsx', 'wb') as new_file:
                new_file.write(down_doc_file)

            # ! Проверить передачу файла напрямую, без сохранения !
            table_all = openpyxl.open(f'C:\\Users\\dmi3\\PycharmProjects\\'
                                      f'Python_Excel_Split_Telegram_bot\\excel_file\\{0}.xlsx')
            sheet = table_all.active

            for row in range(int(head_line) + 1, int(max_line) + 1):

                new_book = openpyxl.Workbook()
                sheet_new_book = new_book.active

                for row_2 in column_list:

                    for row_3 in range(1, (int(head_line)) + 1):  # Head
                        sheet_new_book[f'{row_2 + str(row_3)}'] = sheet[f'{row_2 + str(row_3)}'].value

                    sheet_new_book[f'{row_2 + str(int(head_line) + 1)}'] = sheet[f'{row_2 + str(row)}'].value

                    if str(row_2) == index_column.upper():
                        break

                data = datetime.datetime.now().strftime("%d.%m.%Y")
                new_book.save(f'C:\\Users\\dmi3\\PycharmProjects\\'
                              f'Python_Excel_Split_Telegram_bot\\excel_file\\'
                              f'({str(row - int(head_line))}) {data}.xlsx')
                new_book.close()

                bot.send_document(message.chat.id,
                                  open(f'C:\\Users\\dmi3\\PycharmProjects\\'
                                       f'Python_Excel_Split_Telegram_bot\\excel_file\\'
                                       f'({str(row - int(head_line))}) {data}.xlsx', 'rb'))

                os.remove(f'C:\\Users\\dmi3\\PycharmProjects\\'
                          f'Python_Excel_Split_Telegram_bot\\excel_file\\'
                          f'({str(row - int(head_line))}) {data}.xlsx')

            os.remove(f'C:\\Users\\dmi3\\PycharmProjects\\'
                      f'Python_Excel_Split_Telegram_bot\\excel_file\\{0}.xlsx')

    except Exception as e:
        print(f'{e}\ndef input_file_5(message)\n')

# Функция реакции на текст
@bot.message_handler(content_types=['text'])
def msg_text(message):
    try:
        if message.text == 'Отправить файл':
            input_file(message)
    except Exception as e:
        print(f'{e}\ndef msg_text(message)\n')


# Temporary solution, unfortunately the most effective
column_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
               'V', 'W', 'X', 'Y', 'Z',
               'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ',
               'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
               'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ',
               'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ']

bot.infinity_polling()

