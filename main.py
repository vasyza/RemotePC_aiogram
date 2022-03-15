# -*- coding: UTF-8 -*-
import aiogram.utils.exceptions

from imports import *


class Form(StatesGroup):
    file = State()


storage = MemoryStorage()
try:
    bot = Bot(token=TOKEN)
except aiogram.utils.exceptions.ValidationError:
    bot = None
    exit("Неправильный токен. Изменить его можно в файле - config.py")
dispatcher = Dispatcher(bot, storage=storage)


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
        "<b>/process_shutdown</b> [id] - Завершает процесс по id\n"
        "<b>/set_active_window</b> [id] - Делает выбранное приложение активным (переносит на передний план)\n"
        "<b>/pc_sleep</b> - Переводит ПК в спящий режим\n"
        "<b>/pc_reboot</b> - Перезагружает ПК\n"
        "<b>/pc_shutdown</b> - Выключает ПК\n"
        "<b>/upload_file</b> - Скачивает отправленный файл\n",
        parse_mode=types.ParseMode.HTML)
    await Botutils.set_default_commands(dispatcher)


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
    await Utils.record_display(message)


@dispatcher.message_handler(commands=["active_process"])
async def command_active_process(message: types.Message) -> None:
    """ Возвращает список активных приложений """
    result = []
    win32gui.EnumWindows(Utils.callback, result)
    await message.answer(f"Активные программы:\nID{' ' * 27}Name\n" + "\n".join(result))


@dispatcher.message_handler(commands=["set_active_window"])
async def command_set_active_window(message: types.Message) -> None:
    """ Выводит окно на передний план по id """
    window_id = message.text.split("/set_active_window ")[-1]
    if window_id.isdigit():
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys("%")
        win32gui.ShowWindow(window_id, 5)
        win32gui.SetForegroundWindow(window_id)
        await message.answer("Успешно!")
    else:
        await message.answer("Введите корректный id. Он должен быть целочисленным.")


@dispatcher.message_handler(commands=["pc_sleep"])
async def command_pc_sleep(message: types.Message) -> None:
    """ Переводит ПК в режим сна """
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")


@dispatcher.message_handler(commands=['pc_reboot'])
async def command_pc_reboot(message: types.Message) -> None:
    """ Перезагружает ПК """
    try:
        subprocess.check_call(['shutdown', "-r" "-t", "0"])
    except Exception as e:
        print(e)


@dispatcher.message_handler(commands=["pc_shutdown"])
async def command_pc_sleep(message: types.Message) -> None:
    """ Выключает ПК """
    os.system("shutdown /s /t 1")


@dispatcher.message_handler(commands=["process_shutdown"])
async def command_process_shutdown(message: types.Message) -> None:
    """ Завершает процесс """
    id_process = message.text.split("/process_shutdown ")[-1]
    if id_process.isdigit():
        await Utils.process_shutdown(message, id_process)
    else:
        await message.answer("Вы ввели некорректный id процесса, он должен быть целочисленным")


@dispatcher.message_handler(commands=['upload_file'])
async def command_upload_file(message: types.Message) -> None:
    """ Загружает полученный файл на компьютер """
    await Form.file.set()
    await message.answer("Отправь файл для загрузки")


@dispatcher.message_handler(state='*', commands='cancel')
async def cancel_handler(message: types.Message, state: FSMContext) -> None:
    """ Отменяет оптравку файла """
    current_state = await state.get_state()
    if current_state is None:
        return

    await state.finish()
    await message.answer("Вы отменили загрузку файла")


@dispatcher.message_handler(lambda message: message.content_type == 'text', state=Form.file,
                            content_types=types.ContentTypes.ANY)
async def check_is_file(message: types.Message) -> None:
    """ Отсеивает текстовые сообщения """
    await message.answer("Текст не является файлом. Отправьте файл или напиши /cancle")


@dispatcher.message_handler(state=Form.file, content_types=types.ContentTypes.ANY)
async def get_file(message: types.Message, state: FSMContext) -> None:
    """ Обработка файлов """
    await Utils.download_file(message)
    await state.finish()


if __name__ == "__main__":
    executor.start_polling(dispatcher, skip_updates=True)
