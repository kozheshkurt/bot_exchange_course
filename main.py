import asyncio

from aiogram import Bot, types, Dispatcher
from aiogram.filters.command import Command

from config import TOKEN, EXCEL_FILENAME


# підключаємося до бота
bot = Bot(token=TOKEN)
dp = Dispatcher()


# обробка команди start
@dp.message(Command("start"))
async def process_start_command(message: types.Message):
    # бот відправляє вітальне повідомлення та свою робочу команду /get_exchange_rate
    await message.answer("Hello!\nI give you *.xlsx file with hourly USD/UAH exchange rates for today.\nPress /get_exchange_rate")


# обробка команди get_exchange_rate
@dp.message(Command("get_exchange_rate"))
async def process_get_exchange_rate_command(message: types.Message):
    file = types.FSInputFile(EXCEL_FILENAME) # створюємо об'єкт потрібного файлу для відправки
    caption = f'Here is USD/UAH exchange rates for today, {message.date.day}.{message.date.month}' # готуємо підпис до файлу, який містить сьогоднішню дату (отриману з параметрів повідомлення)
    await bot.send_document(message.from_user.id, document=file, caption=caption) # відправляємо користувачу файл


# запуск бота
async def main():
    await dp.start_polling(bot)

# запуск циклу асинхронної програми
if __name__ == "__main__":
    asyncio.run(main())