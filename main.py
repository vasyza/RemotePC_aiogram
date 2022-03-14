# -*- coding: UTF-8 -*-
from imports import *
from config import TOKEN


def callback(hwnd, hwnds) -> None:
    """ Обрабатывает активные\неактивные приложения """
    if win32gui.IsWindowVisible(hwnd):
        if not win32gui.GetWindow(hwnd, 4):
            name = win32gui.GetWindowText(hwnd)
            if name:
                line = f"{hwnd}{'  ' * (15 - len(str(hwnd)))}{name}"
                if len(line) > 64:
                    hwnds.append(line[:64])
                    hwnds.append(' ' * 31 + line[64:])
                else:
                    hwnds.append(line)


bot = Bot(token=TOKEN)
dispatcher = Dispatcher(bot)


async def set_default_commands(dp) -> None:
    """ Создает список команд бота """
    await dp.bot.set_my_commands([
        types.BotCommand("start", "Запустить бота"),
        types.BotCommand("cmd", "Выполнить команду в командной строке"),
        types.BotCommand("browser", "Открыть ссылку в браузере"),
        types.BotCommand("screenshot", "Сделать скриншот экрана, отправить"),
        types.BotCommand("create_video_dem", "Записать видео, отправить"),
        types.BotCommand("active_process", "Посмотреть список активных программ"),
        types.BotCommand("set_active_window", "Сделать приложение активным"),
        types.BotCommand("pc_sleep", "Переводит ПК в спящий режим"),
        types.BotCommand("pc_reboot", "Перезагружает ПК"),
        types.BotCommand("pc_shutdown", "Выключает ПК"),
    ])


@dispatcher.message_handler(commands=["start", "help"])
async def command_start(message: types.Message) -> None:
    """ Отправляет FAQ по командам """
    await message.answer(
        "Привет, данный бот даст тебе возможность управлять компьютером с помощью команд\nUsage:\n"
        "<b>/cmd</b> [команда] - Выполняет полученную команду в командной строке\n"
        "<b>/browser</b> [ссылка] - Открывает полученную ссылку в браузере (ссылка должна начинаться на http\https)\n"
        "<b>/screenshot</b> - Скриншотит экран, отправляя фотографию\n"
        "<b>/create_video_dem</b> [длительность] - Записывает демонстрацию экрана, отправляя видео\n"
        "<b>/active_process</b> - Выводит список активных программ (окон)\n"
        "<b>/set_active_window</b> [id] - Делает выбранное приложение активным (переносит на передний план)\n"
        "<b>/pc_sleep</b> - Переводит ПК в спящий режим\n"
        "<b>/pc_reboot</b> - Перезагружает ПК\n"
        "<b>/pc_shutdown</b> - Выключает ПК\n",
        parse_mode=types.ParseMode.HTML)
    await set_default_commands(dispatcher)


@dispatcher.message_handler(commands=["cmd"])
async def command_cmd(message: types.Message) -> None:
    """ Выполняет команду в cmd """
    command = message.text.split("/cmd ")[-1]
    result = "None"
    try:
        output = subprocess.getoutput(command)
        result = str(bytes(output, "cp1251"), "ibm866")
    except subprocess.CalledProcessError as e:
        result = e.output
    await message.answer(f"Вывод:\n{result}" if command != "/cmd" else "Неправильный аргумент")


@dispatcher.message_handler(commands=["browser"])
async def command_browser(message: types.Message) -> None:
    """ Открывает ссылку в браузере """
    regx = r"[htps]{4,5}://[a-zA-Z0-9]{,}\."
    url = message.text.split("/browser ")[-1]
    if re.findall(regx, url):
        if webbrowser.open(url):
            await message.answer("Ссылка успешно открыта")
        else:
            await message.answer("Я не смог открыть ссылку")
    else:
        await message.answer("Введена неправильная ссылка!")


@dispatcher.message_handler(commands=["screenshot"])
async def command_screenshot(message: types.Message) -> None:
    """ Скриншотит экран, отправляя фото в виде документа """
    file_name = "temp.png"
    pyautogui.screenshot(file_name)
    await message.answer_document(open(file_name, "rb"))
    os.remove(file_name)


@dispatcher.message_handler(commands=["create_video_dem"])
async def command_create_video_dem(message: types.Message) -> None:
    """ Записывает демонстрацию экрана, отправляя видео """
    file_name = "temp.avi"
    count_of_seconds = message.text.split("/create_video_dem ")[-1]
    seconds = int(count_of_seconds) if count_of_seconds.isdigit() else 10
    screen_size = pyautogui.size()
    fourcc = cv2.VideoWriter_fourcc(*"XVID")
    out = cv2.VideoWriter(file_name, fourcc, 10.0, screen_size)
    for _ in range(10 * seconds):
        img = pyautogui.screenshot()
        frame = np.array(img)
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        out.write(frame)
    cv2.destroyAllWindows()
    out.release()
    await message.answer_video(video=open(file_name, "rb"))
    os.remove(file_name)


@dispatcher.message_handler(commands=['active_process'])
async def command_active_process(message: types.Message) -> None:
    """ Возвращает список активных приложений """
    result = []
    win32gui.EnumWindows(callback, result)
    await message.answer(f"Активные программы:\nID{' ' * 27}Name\n" + '\n'.join(result))


@dispatcher.message_handler(commands=['set_active_window'])
async def command_set_active_window(message: types.Message) -> None:
    """ Выводит окно на передний план по id """
    window_id = message.text.split('/set_active_window ')[-1]
    if window_id.isdigit():
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.ShowWindow(window_id, 5)
        win32gui.SetForegroundWindow(window_id)
        await message.answer("Успешно!")
    else:
        await message.answer("Введите корректный id. Он должен быть целочисленным.")


@dispatcher.message_handler(commands=['pc_sleep'])
async def command_pc_sleep(message: types.Message) -> None:
    """ Переводит ПК в режим сна """
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")


@dispatcher.message_handler(commands=['pc_reboot'])
async def command_pc_reboot(message: types.Message) -> None:
    """ Перезагружает ПК """
    try:
        subprocess.check_call(['shutdown', '-r' '-t', '0'])
    except Exception as e:
        print(e)


@dispatcher.message_handler(commands=['pc_shutdown'])
async def command_pc_sleep(message: types.Message) -> None:
    """ Выключает ПК """
    os.system("shutdown /s /t 1")


if __name__ == '__main__':
    executor.start_polling(dispatcher, skip_updates=True)
