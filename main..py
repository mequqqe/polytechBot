from docx import Document
import telebot
import os

# Инициализация
token = "6654160154:AAG0Ylq41itMfeMj5HxLZW699DKpnKQhCKU"
bot = telebot.TeleBot(token)
cabinet_schedule_dict = {}
group_schedule_dict = {}

def extract_data_from_docx(file_info):
    global cabinet_schedule_dict, group_schedule_dict
    file_data = bot.download_file(file_info.file_path)

    # Сохранение данных файла во временный файл
    temp_file = "temp_schedule.docx"
    with open(temp_file, "wb") as f:
        f.write(file_data)
    
    doc = Document(temp_file)
    os.remove(temp_file)  # Удаляем временный файл после обработки

    for table in doc.tables:
        for row in table.rows:
            cabinet = row.cells[0].text.strip().lower()
            if not cabinet:  # Пропускаем строки без кабинета
                continue

            for i, cell in enumerate(row.cells[1:-1:2], start=1):  # Обновлено, чтобы учесть все столбцы и добавить индексацию
                teacher = cell.text.strip().lower()
                group = row.cells[i*2].text.strip().lower()

                # Заполнение словаря для кабинетов
                if cabinet not in cabinet_schedule_dict:
                    cabinet_schedule_dict[cabinet] = []
                cabinet_schedule_dict[cabinet].append((i, teacher, group))  # Добавлен номер пары

                # Заполнение словаря для групп
                if group not in group_schedule_dict:
                    group_schedule_dict[group] = []
                group_schedule_dict[group].append((i, teacher, cabinet))  # Добавлен номер пары

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    try:
        file_info = bot.get_file(message.document.file_id)
        extract_data_from_docx(file_info)
        bot.reply_to(message, "Расписание успешно обновлено!")
    except Exception as e:
        bot.reply_to(message, f"Ошибка при обработке файла: {str(e)}.")

@bot.message_handler(func=lambda message: True)
def send_schedule(message):
    query = message.text.strip().lower()
    response = []

    # Поиск по кабинетам
    if query in cabinet_schedule_dict:
        for pair_num, teacher, group in cabinet_schedule_dict[query]:
            response.append(f"{query} - {pair_num}-я пара: {teacher} ({group})")
    # Поиск по группам
    elif query in group_schedule_dict:
        for pair_num, teacher, cabinet in group_schedule_dict[query]:
            response.append(f"{cabinet} - {pair_num}-я пара: {teacher} ({query})")
    else:
        response.append(f"Расписание для {query} отсутствует.")

    bot.reply_to(message, "\n".join(response))

bot.polling()


