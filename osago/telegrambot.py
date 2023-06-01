import telebot
import pandas as pd
import os
import main as main
import threading as th
import time
bot = telebot.TeleBot('5705208354:AAEStTLcM89Y4pcBg1WOCXknxWAfAmkT0W8')


@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.send_message(message.chat.id, "Здравствуйте! Я бот, который поможет вам узнать больше информации об ОСАГО.")
    bot.send_message(message.chat.id, "Не забудьте вызвать команду \"/help\" (без ковычек), чтобы узнать, как правильно мной пользоваться")

@bot.message_handler(commands=['help'])
def send_help(message):
    bot.send_photo(message.chat.id, open('Снимок.PNG', 'rb'), caption='Для того, чтобы начать работу, вам необходимо отправить мне файл в формате .xlsx такого вида.')
    bot.send_message(message.chat.id, "После отправки документа, бот будет не активен, пока не обработает отаправленный файл.")



@bot.message_handler(content_types=['document'])
def get_doc(message):
    file_info = bot.get_file(message.document.file_id)
    if(file_info.file_path.split('.')[-1]!='xlsx'):
        bot.send_message(message.chat.id, "ВНИМАНИЕ! ОТПРАВЛЕН ФАЙЛ С ДРУГИМ РАСШИРЕНИЕМ!")
        bot.send_message(message.chat.id, "Не забудьте вызвать команду \"/help\" (без ковычек), чтобы узнать, как правильно мной пользоваться")
        return
    downloaded_file = bot.download_file(file_info.file_path)
    if not os.path.exists(file_info.file_path.split('/')[0]):
        os.mkdir(file_info.file_path.split('/')[0])
    with open(file_info.file_path, 'wb') as new_file:
        new_file.write(downloaded_file)
    df_new = pd.read_excel(file_info.file_path)
    if df_new.columns.tolist()[0]!='osago.vin':
        os.remove(file_info.file_path)
        bot.send_message(message.chat.id, "ВНИМАНИЕ! ОТПРАВЛЕН ФАЙЛ С НЕ ПРАВИЛЬНЫМИ ДАННЫМИ!")
        bot.send_message(message.chat.id, "Не забудьте вызвать команду \"/help\" (без ковычек), чтобы узнать, как правильно мной пользоваться")
        return
    bot.send_message(message.chat.id, "Файл принят в обработку, пожалуйста, дождитесь конца работы.")
    out_file = file_info.file_path.split('.')[-2]+"_out.xlsx"
    main_thread = th.Thread(target=main.main, args=[file_info.file_path, out_file, message.chat.id, bot])
    main_thread.start()


while True:
    try:    
        bot.polling(none_stop=True)
    except:
        time.sleep(15)