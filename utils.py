from imports import *


class Utils:
    @staticmethod
    async def download_file(message) -> None:
        message_id = message.message_id
        path = os.path.expanduser(r'~/Desktop/RemoteDesktop_Files')
        if message.photo:
            await message.photo[-1].download(destination_dir=os.path.expanduser(f"~/Desktop/RemoteDesktop_Files"))
            await message.answer(
                f"Файл успешно скачан\nРасположение: {os.path.expanduser('~/Desktop/RemoteDesktop_Files/photos')}")
        elif message.document:
            file_name = f"_{message_id}.".join(message.document.file_name.split("."))
            await message.document.download(destination_file=path + f"/documents/{file_name}")
            await message.answer(
                f"Файл успешно скачан\nРасположение: {path}/documents\nИмя: {file_name}")
        elif message.audio:
            file_name = f"_{message_id}.".join(message.audio.file_name.split("."))
            await message.audio.download(
                destination_file=path + f"/audio/{file_name}")
            await message.answer(
                f"Файл успешно скачан\nРасположение: {path}/audio\nИмя: {file_name}")
        elif message.video:
            file_name = f"_{message_id}.".join(message.video.file_name.split("."))
            await message.video.download(
                destination_file=path + f"/videos/{file_name}")
            await message.answer(
                f"Файл успешно скачан\nРасположение: {path}/videos\nИмя: {file_name}")
        else:
            await message.answer("Я не умею скачивать данный файл, напишите разработчику в тг @vasyza")

    @staticmethod
    async def process_shutdown(message, id_process):
        try:
            win32gui.PostMessage(id_process, win32con.WM_CLOSE, 0, 0)
            await message.answer("Успешно!")
        except Exception as e:
            if str(e).find("Недопустимый дескриптор окна."):
                await message.answer("Вы ввели некорректный id процесса, такого id не существует")

    @staticmethod
    async def record_display(message):
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

    @staticmethod
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


class Botutils:
    @staticmethod
    async def set_default_commands(dp) -> None:
        """ Создает список команд бота """
        await dp.bot.set_my_commands([
            types.BotCommand("start", "Запустить бота"),
            types.BotCommand("cmd", "Выполнить команду в командной строке"),
            types.BotCommand("browser", "Открыть ссылку в браузере"),
            types.BotCommand("screenshot", "Сделать скриншот экрана, отправить"),
            types.BotCommand("create_video_dem", "Записать видео, отправить"),
            types.BotCommand("active_process", "Посмотреть список активных программ"),
            types.BotCommand("process_shutdown", "Завершаает процесс по id"),
            types.BotCommand("set_active_window", "Сделать приложение активным"),
            types.BotCommand("pc_sleep", "Переводит ПК в спящий режим"),
            types.BotCommand("pc_reboot", "Перезагружает ПК"),
            types.BotCommand("pc_shutdown", "Выключает ПК"),
            types.BotCommand("upload_file", "Скачивает отправленный файл"),
        ])
