# bot.py
import logging
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import Message, BufferedInputFile
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from parser import parse_wildberries  # Импортируем функцию парсера

# Настройка логирования
logging.basicConfig(level=logging.INFO)

# Токен вашего бота
API_TOKEN = '7869041656:AAHflRmOCs28B_7-y8THO9rI7nbkoV6Fweo'

# Инициализация бота и диспетчера
bot = Bot(token=API_TOKEN)
dp = Dispatcher()

# Состояния для FSM
class Form(StatesGroup):
    waiting_for_query = State()

# Команда /start
@dp.message(Command("start"))
async def send_welcome(message: Message):
    await message.answer(
        "Привет! Я бот для парсинга товаров с Wildberries.\n"
        "Используйте команду /parse, чтобы начать парсинг."
    )

# Команда /help
@dp.message(Command("help"))
async def send_help(message: Message):
    await message.answer(
        "Вот что я умею:\n"
        "/start — начать работу с ботом\n"
        "/parse — запустить парсинг товаров с Wildberries"
    )

# Команда /parse
@dp.message(Command("parse"))
async def parse_command(message: Message, state: FSMContext):
    await message.answer("Введите название товара для поиска:")
    await state.set_state(Form.waiting_for_query)

# Обработка ввода пользователя после команды /parse
@dp.message(Form.waiting_for_query)
async def process_query(message: Message, state: FSMContext):
    query = message.text.strip()
    if not query:
        await message.reply("Пожалуйста, укажите название товара.")
        return

    await message.answer("Выполняю поиск товаров...")

    # Парсим данные через функцию из parser.py
    result = parse_wildberries(query)

    if isinstance(result, tuple) and len(result) == 2:
        # Если функция вернула два значения
        products, excel_filename = result
    else:
        # Если функция вернула только частично собранные данные
        products = result
        excel_filename = None

    if not products:
        await message.reply("Не удалось найти товары по вашему запросу.")
        await state.clear()
        return

    # Если есть файл Excel, отправляем его
    if excel_filename:
        with open(excel_filename, "rb") as file:
            await message.answer_document(BufferedInputFile(file.read(), filename=excel_filename))
    else:
        await message.reply("Не удалось создать файл с результатами.")

    # Сбрасываем состояние
    await state.clear()

# Запуск бота
if __name__ == '__main__':
    dp.run_polling(bot)