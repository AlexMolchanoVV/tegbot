import telebot
import openpyxl #модуль для открытия и чтения excel файлов
import datetime #модуль определения даты и дня недели
import urllib.request #модуль для скачивания файла

bot = telebot.TeleBot("1615943763:AAFU12EPuEyNRxPv-iAztAEYvjHN1W9QWQw") #токен бота

book = openpyxl.open("rasp1.xlsx",read_only=True) #открывает таблицу , только для чтения

sheet = book.active #по умолчанию чтение первой страницы excel

day_ned = datetime.datetime.today().weekday()  #день недели 0 - понедельник , 5 - суббота

a = 0

nedely = 3

if day_ned == 0:  #если понедельник
    nedely = nedely + 1  #следующая неделя

@bot.message_handler(commands=['start'])
def start_message(message):
    keyboard = telebot.types.ReplyKeyboardMarkup(True)  #создание клавиатуры
    keyboard.row('Расписание на сегодня', 'Расписание на завтра', 'Расписание на неделю', 'Обновить расписание')  #текст для кнопок
    bot.send_message(message.chat.id, 'Привет, тут ты можешь узнать расписание для группы ИНБО-02-19!', reply_markup=keyboard) #первое сообщение от бота

@bot.message_handler(content_types=['text'])
def send_text(message):
    if message.text.lower() == 'расписание на сегодня':
        if nedely % 2 != 0:   #проверка на четность/нечетность предметов
        # Понедельник
            if day_ned == 0:  #определение дня недели
                for row in range(10, 13, 2):
                    ponedelnik = sheet[row][30].value  #считывает названия предметов
                    time = sheet[row][3].value         #считывает второе время
                    number = sheet[row][2].value       #считывает первое время
                    bot.send_message(message.chat.id, f'Предмет: {ponedelnik}')  #вывод предметов
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')   #вывод времени
            # Вторник
            elif day_ned == 1:
                for row in range(18, 27, 2):
                    vtornik = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {vtornik}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Среда
            elif day_ned == 2:
                for row in range(34, 39, 2):
                    sreda = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {sreda}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Четверг
            elif day_ned == 3:
                bot.send_message(message.chat.id, 'Сегодня пар нет, отдыхай!')
            # Пятница
            elif day_ned == 4:
                for row in range(54, 61, 2):
                    pyt = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {pyt}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Суббота
            elif day_ned == 5:
                for row in range(68, 71, 2):
                    sybota = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {sybota}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            else:
                bot.send_message(message.chat.id, 'Сегодня занятий нет, отдыхай!')
        elif nedely % 2 == 0:
            # Понедельник
            if day_ned == 0:
                for row in range(13, 16, 2):
                    ponedelnik = sheet[row][30].value
                    time = sheet[row-1][3].value
                    number = sheet[row-1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {ponedelnik}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Вторник
            elif day_ned == 1:
                for row in range(19, 28, 2):
                    vtornik = sheet[row][30].value
                    time = sheet[row-1][3].value
                    number = sheet[row-1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {vtornik}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Среда
            elif day_ned == 2:
                for row in range(37, 40, 2):
                    sreda = sheet[row][30].value
                    time = sheet[row-1][3].value
                    number = sheet[row-1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {sreda}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Четверг
            elif day_ned == 3:
                 for row in range(45, 48, 2):
                    chetverg = sheet[row][30].value
                    time = sheet[row-1][3].value
                    number = sheet[row-1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {chetverg}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Пятница
            elif day_ned == 4:
                for row in range(55, 62, 2):
                    pyt = sheet[row][30].value
                    time = sheet[row-1][3].value
                    number = sheet[row-1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {pyt}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Суббота
            elif day_ned == 5:
                for row in range(67, 70, 2):
                    sybota = sheet[row][30].value
                    time = sheet[row-1][3].value
                    number = sheet[row-1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {sybota}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            else:
                bot.send_message(message.chat.id, 'Сегодня занятий нет, отдыхай!')

    elif message.text.lower() == 'расписание на завтра':
        a = day_ned
        a = a + 1
        if a == 7:
            a = 0
        if nedely % 2 != 0:
        # Понедельник
            if a == 0:
                for row in range(10, 13, 2):
                    ponedelnik = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {ponedelnik}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Вторник
            elif a == 1:
                for row in range(18, 27, 2):
                    vtornik = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {vtornik}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Среда
            elif a == 2:
                for row in range(34, 39, 2):
                    sreda = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {sreda}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Четверг
            elif a == 3:
                bot.send_message(message.chat.id, 'Сегодня пар нет, отдыхай!')
            # Пятница
            elif a == 4:
                for row in range(54, 61, 2):
                    pyt = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {pyt}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Суббота
            elif a == 5:
                for row in range(68, 71, 2):
                    sybota = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {sybota}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            else:
                bot.send_message(message.chat.id, 'Сегодня занятий нет, отдыхай!')
        elif nedely % 2 == 0:
            # Понедельник
            if a == 0:
                for row in range(13, 16, 2):
                    ponedelnik = sheet[row][30].value
                    time = sheet[row-1][3].value
                    number = sheet[row-1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {ponedelnik}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Вторник
            elif a == 1:
                for row in range(19, 28, 2):
                    vtornik = sheet[row][30].value
                    time = sheet[row-1][3].value
                    number = sheet[row-1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {vtornik}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Среда
            elif a == 2:
                for row in range(37, 40, 2):
                    sreda = sheet[row][30].value
                    time = sheet[row-1][3].value
                    number = sheet[row-1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {sreda}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Четверг
            elif a == 3:
                 for row in range(45, 48, 2):
                    chetverg = sheet[row][30].value
                    time = sheet[row-1][3].value
                    number = sheet[row-1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {chetverg}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Пятница
            elif a == 4:
                for row in range(55, 62, 2):
                    pyt = sheet[row][30].value
                    time = sheet[row-1][3].value
                    number = sheet[row-1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {pyt}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Суббота
            elif a == 5:
                for row in range(67, 70, 2):
                    sybota = sheet[row][30].value
                    time = sheet[row-1][3].value
                    number = sheet[row-1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {sybota}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            else:
                bot.send_message(message.chat.id, 'Сегодня занятий нет, отдыхай!')

    elif message.text.lower() == 'расписание на неделю':
        a = day_ned
        if nedely % 2 != 0:
                bot.send_message(message.chat.id, 'Понедельник')
                for row in range(10, 13, 2):
                    ponedelnik = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {ponedelnik}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
                # Вторник
                bot.send_message(message.chat.id, '━━━━━')
                bot.send_message(message.chat.id, 'Вторник')
                for row in range(18, 27, 2):
                    vtornik = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {vtornik}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
                # Среда
                bot.send_message(message.chat.id, '━━━━━')
                bot.send_message(message.chat.id, 'Среда')
                for row in range(34, 39, 2):
                    sreda = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {sreda}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
                # Четверг
                bot.send_message(message.chat.id, '━━━━━')
                bot.send_message(message.chat.id, 'Четверг')
                bot.send_message(message.chat.id, 'Сегодня пар нет, отдыхай!')
                # Пятница
                bot.send_message(message.chat.id, '━━━━━')
                bot.send_message(message.chat.id, 'Пятница')
                for row in range(54, 61, 2):
                    pyt = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {pyt}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
                # Суббота
                bot.send_message(message.chat.id, '━━━━━')
                bot.send_message(message.chat.id, 'Суббота')
                for row in range(68, 71, 2):
                    sybota = sheet[row][30].value
                    time = sheet[row][3].value
                    number = sheet[row][2].value
                    bot.send_message(message.chat.id, f'Предмет: {sybota}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
        if nedely % 2 == 0:
                bot.send_message(message.chat.id, 'Понедельник')
                for row in range(13, 16, 2):
                    ponedelnik = sheet[row][30].value
                    time = sheet[row - 1][3].value
                    number = sheet[row - 1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {ponedelnik}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Вторник
                bot.send_message(message.chat.id, '━━━━━')
                bot.send_message(message.chat.id, 'Вторник')
                for row in range(19, 28, 2):
                    vtornik = sheet[row][30].value
                    time = sheet[row - 1][3].value
                    number = sheet[row - 1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {vtornik}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Среда
                bot.send_message(message.chat.id, '━━━━━')
                bot.send_message(message.chat.id, 'Среда')
                for row in range(37, 40, 2):
                    sreda = sheet[row][30].value
                    time = sheet[row - 1][3].value
                    number = sheet[row - 1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {sreda}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Четверг
                bot.send_message(message.chat.id, '━━━━━')
                bot.send_message(message.chat.id, 'Четверг')
                for row in range(45, 48, 2):
                    chetverg = sheet[row][30].value
                    time = sheet[row - 1][3].value
                    number = sheet[row - 1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {chetverg}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Пятница
                bot.send_message(message.chat.id, '━━━━━')
                bot.send_message(message.chat.id, 'Пятница')
                for row in range(55, 62, 2):
                    pyt = sheet[row][30].value
                    time = sheet[row - 1][3].value
                    number = sheet[row - 1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {pyt}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
            # Суббота
                bot.send_message(message.chat.id, '━━━━━')
                bot.send_message(message.chat.id, 'Суббота')
                for row in range(67, 70, 2):
                    sybota = sheet[row][30].value
                    time = sheet[row - 1][3].value
                    number = sheet[row - 1][2].value
                    bot.send_message(message.chat.id, f'Предмет: {sybota}')
                    bot.send_message(message.chat.id, f'Начало: {number} Конец: {time}')
    if message.text.lower() == 'обновить расписание':
        link1 = 'https://webservices.mirea.ru/upload/iblock/56e/%D0%98%D0%98%D0%A2_2%D0%BA_20-21_%D0%B2%D0%B5%D1%81%D0%BD%D0%B0.xlsx'  #скачивание файла
        urllib.request.urlretrieve(link1, "rasp1.xlsx")   #файл получает новое имя
        bot.send_message(message.chat.id, 'Расписание обновлено, наслаждайтесь')  #вывод сообщения

#RUN
bot.polling()  #запуск