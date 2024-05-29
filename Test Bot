import telebot
from telebot import types
import openpyxl
from io import BytesIO

API_TOKEN = '6892898978:AAFnDZcFGczwed5Dt-vJq3Hpi5lsJu9ZEq4'
bot = telebot.TeleBot(API_TOKEN)

# Словарь для хранения заявок
orders = []

# Функция для создания XLSX файла
def create_xlsx(orders):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Заказ 1С", "Заказчик", "Задача RM", "Дата отгрузки по договору", "Срок завершения по задаче"])

    for order in orders:
        ws.append(order)

    file_stream = BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(message, "Добро пожаловать! Нажмите /new_order чтобы создать новый заказ.")

# Обработчик команды /new_order
@bot.message_handler(commands=['new_order'])
def new_order(message):
    msg = bot.reply_to(message, "Введите данные заказа в формате: Заказ 1С, Заказчик, Задача RM, Дата отгрузки по договору, Срок завершения по задаче")
    bot.register_next_step_handler(msg, process_order)

# Обработчик ввода данных заказа
def process_order(message):
    try:
        order = message.text.split(',')
        if len(order) != 5:
            raise ValueError("Неправильный формат данных")
        orders.append(order)
        bot.reply_to(message, "Заказ добавлен! Нажмите /download чтобы скачать файл с заказами.")
    except Exception as e:
        bot.reply_to(message, f"Ошибка: {e}. Пожалуйста, введите данные в правильном формате.")

# Обработчик команды /download
@bot.message_handler(commands=['download'])
def download_file(message):
    if not orders:
        bot.reply_to(message, "Нет заказов для создания файла.")
        return

    file_stream = create_xlsx(orders)
    bot.send_document(message.chat.id, file_stream, caption="Ваш файл с заказами")

bot.infinity_polling()
