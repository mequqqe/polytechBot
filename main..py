import telebot
from docx import Document
import io

bot = telebot.TeleBot('6654160154:AAG0Ylq41itMfeMj5HxLZW699DKpnKQhCKU')

schedule_dict = {}

def improved_parse_schedule(docx_file):
    doc = Document(docx_file)
    time_intervals = [cell.text.strip() for cell in doc.tables[0].rows[0].cells][1::2]
    schedule = {}

    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            room = cells.pop(0)

            # Process data in pairs (Teacher, Group)
            for i in range(0, len(cells) - 1, 2):  # We subtract 1 to avoid index out of range
                teachers = cells[i].split("\n") if i < len(cells) else []
                groups = cells[i+1].split("\n") if (i+1) < len(cells) else []
                
                for teacher, group in zip(teachers, groups):
                    if not teacher or not group:
                        continue

                    if group not in schedule:
                        schedule[group] = {}
                    
                    interval_index = i // 2
                    current_interval = time_intervals[interval_index]

                    schedule[group][current_interval] = {'room': room, 'teacher': teacher}

    return schedule

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(message, "Привет! Введите название вашей группы, чтобы узнать расписание на сегодня. Или загрузите файл с расписанием для обновления.")

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    global schedule_dict
    if message.document.mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        with io.BytesIO(downloaded_file) as docx_file:
            schedule_dict = improved_parse_schedule(docx_file)
        bot.reply_to(message, "Расписание успешно обновлено!")
    else:
        bot.reply_to(message, "Пожалуйста, загрузите документ в формате .docx.")

@bot.message_handler(func=lambda m: True)
def send_schedule(message):
    group = message.text
    if group in schedule_dict:
        schedule_for_group = schedule_dict[group]
        response = "Расписание для группы {}:".format(group)
        for time, details in schedule_for_group.items():
            teacher_info = details["teacher"] if details["teacher"] else "Преподаватель: Не указан"
            response += "\n{} - Аудитория: {} - {}".format(time, details["room"], teacher_info)
    else:
        response = "Извините, расписание для группы {} не найдено.".format(group)
    bot.reply_to(message, response)

bot.polling()






