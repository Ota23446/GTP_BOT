import os
import logging
import openpyxl
from datetime import datetime, timedelta
from aiogram import Router, F
from aiogram.filters import Command, CommandStart
from aiogram.types import Message, CallbackQuery
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.utils.keyboard import InlineKeyboardBuilder

from config import ADMIN_USERS
from utils import (
    UserDataManager,
    NOTIFICATION_TYPES,
    format_notification_status,
    is_valid_username
)
from services import get_next_shift, calculate_worked_time

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

router = Router()
user_manager = UserDataManager()

class NotificationStates(StatesGroup):
    waiting_for_time = State()
    waiting_for_username = State()


# Базовые команды
@router.message(CommandStart())
async def cmd_start(message: Message):
    username = user_manager.get_user_by_telegram_id(str(message.from_user.id))
    if username:
        await message.answer(
            f"Добро пожаловать обратно!\n"
            f"Ваш username: {username}\n"
            f"Используйте /help для просмотра доступных команд."
        )
    else:
        await message.answer(
            "Добро пожаловать!\n"
            "Для начала работы необходимо привязать ваш username.\n"
            "Используйте команду /register sm_username"
        )


@router.message(Command("help"))
async def cmd_help(message: Message):
    help_text = (
        "Доступные команды:\n"
        "/start - Начать работу с ботом\n"
        "/register sm_username - Привязать username\n"
        "/settings - Настройки уведомлений\n"
        "/time - Изменить время уведомлений\n"
        "/status - Просмотр текущих настроек\n"
        "/shift - Информация о следующей смене\n"
        "/worked_time - Подсчет отработанного времени"
    )
    await message.answer(help_text)

# Регистрация пользователя
@router.message(Command("register"))
async def cmd_register(message: Message):
    args = message.text.split()
    if len(args) != 2:
        await message.answer("Использование: /register sm_username")
        return

    username = args[1].lower()
    if not await is_valid_username(username):
        await message.answer("Неверный формат username. Должен начинаться с 'sm_' и быть в нижнем регистре.")
        return

    data = user_manager.load_user_data()
    if username not in data:
        await message.answer("Такой username не найден в системе.")
        return

    # Проверяем, не привязан ли уже этот username к другому пользователю
    if data[username]["user_id"] != "0" and data[username]["user_id"] != str(message.from_user.id):
        await message.answer("Этот username уже привязан к другому пользователю.")
        return

    data[username]["user_id"] = str(message.from_user.id)
    user_manager.save_user_data(data)
    await message.answer("Username успешно привязан!")


# Настройки уведомлений
@router.message(Command("settings"))
async def cmd_settings(message: Message):
    username = user_manager.get_user_by_telegram_id(str(message.from_user.id))
    if not username:
        await message.answer("Сначала необходимо привязать username через /register")
        return

    builder = InlineKeyboardBuilder()
    user_data = user_manager.get_user_settings(username)

    for notif_key, notif_name in NOTIFICATION_TYPES.items():
        status = user_data["notifications"][notif_key]
        builder.button(
            text=f"{notif_name}: {format_notification_status(status)}",
            callback_data=f"toggle_{notif_key}"
        )
    builder.adjust(1)

    await message.answer(
        "Настройки уведомлений:",
        reply_markup=builder.as_markup()
    )


# Обработка нажатий на кнопки настроек
@router.callback_query(F.data.startswith("toggle_"))
async def process_notification_toggle(callback: CallbackQuery):
    username = user_manager.get_user_by_telegram_id(str(callback.from_user.id))
    if not username:
        await callback.answer("Ошибка: пользователь не найден", show_alert=True)
        return

    notif_type = callback.data.split("_")[1]
    user_data = user_manager.get_user_settings(username)
    current_status = user_data["notifications"][notif_type]

    # Инвертируем статус уведомления
    user_manager.update_user_notifications(username, notif_type, not current_status)

    # Обновляем клавиатуру
    builder = InlineKeyboardBuilder()
    user_data = user_manager.get_user_settings(username)

    for nk, nn in NOTIFICATION_TYPES.items():
        status = user_data["notifications"][nk]
        builder.button(
            text=f"{nn}: {format_notification_status(status)}",
            callback_data=f"toggle_{nk}"
        )
    builder.adjust(1)

    await callback.message.edit_reply_markup(reply_markup=builder.as_markup())
    await callback.answer()


# Изменение времени уведомлений
@router.message(Command("time"))
async def cmd_time(message: Message, state: FSMContext):
    username = user_manager.get_user_by_telegram_id(str(message.from_user.id))
    if not username:
        await message.answer("Сначала необходимо привязать username через /register")
        return

    await state.set_state(NotificationStates.waiting_for_time)
    await message.answer(
        "Введите желаемое время уведомлений в формате ЧЧ:ММ (например, 19:00)"
    )


@router.message(NotificationStates.waiting_for_time)
async def process_notification_time(message: Message, state: FSMContext):
    username = user_manager.get_user_by_telegram_id(str(message.from_user.id))

    if user_manager.update_notification_time(username, message.text):
        await message.answer(f"Время уведомлений успешно обновлено на {message.text}")
    else:
        await message.answer("Неверный формат времени. Используйте формат ЧЧ:ММ (например, 19:00)")

    await state.clear()

@router.message(Command("status"))
async def cmd_status(message: Message):
    username = user_manager.get_user_by_telegram_id(str(message.from_user.id))
    if not username:
        await message.answer("Сначала необходимо привязать username через /register")
        return

    user_data = user_manager.get_user_settings(username)
    if not user_data:
        await message.answer("Ошибка: данные пользователя не найдены")
        return

    status_text = [
        f"📱 Username: {username}",
        f"⏰ Время уведомлений: {user_data['notification_time']}",
        "\n📋 Статус уведомлений:"
    ]

    for notif_key, notif_name in NOTIFICATION_TYPES.items():
        status = user_data["notifications"][notif_key]
        status_text.append(f"- {notif_name}: {format_notification_status(status)}")

    await message.answer("\n".join(status_text))

@router.message(Command("shift"))
async def cmd_shift(message: Message):
    username = user_manager.get_user_by_telegram_id(str(message.from_user.id))
    if not username:
        await message.answer("Сначала необходимо привязать username через /register")
        return

    try:
        next_shift_info = await get_next_shift(username)
        await message.answer(next_shift_info)
    except Exception as e:
        logging.error(f"Error in shift command: {e}")
        await message.answer("Произошла ошибка при получении информации о смене")

@router.message(Command("worked_time"))
async def cmd_worked_time(message: Message):
    username = user_manager.get_user_by_telegram_id(str(message.from_user.id))
    if not username:
        await message.answer("Сначала необходимо привязать username через /register")
        return

    try:
        worked_time = await calculate_worked_time(username)
        await message.answer(worked_time)
    except Exception as e:
        logging.error(f"Error in worked_time command: {e}")
        await message.answer("Произошла ошибка при подсчете отработанного времени")


@router.message(Command("debug_schedule"))
async def cmd_debug_schedule(message: Message):
    if str(message.from_user.id) not in ADMIN_USERS:
        return

    try:
        wb = openpyxl.load_workbook('schedule.xlsx')
        ws = wb.active

        debug_info = ["📋 Отладочная информация о расписании:"]

        # Получаем значение ячейки для проверки
        username = "sm_kirillts"
        current_date = datetime.now()
        next_date = current_date + timedelta(days=1)
        next_day_col = next_date.day + 1

        debug_info.append(f"\n📅 Текущая дата: {current_date.strftime('%d.%m.%Y')}")
        debug_info.append(f"📅 Следующая дата: {next_date.strftime('%d.%m.%Y')}")
        debug_info.append(f"📊 Номер колонки для следующей даты: {next_day_col}")

        found = False
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and str(cell.value).lower() == username.lower():
                    shift = ws.cell(row=cell.row, column=next_day_col).value
                    debug_info.append(f"\n👤 Найден пользователь: {username}")
                    debug_info.append(f"📍 Строка в таблице: {cell.row}")
                    debug_info.append(f"🔄 Значение смены: {shift}")
                    found = True
                    break
            if found:
                break

        if not found:
            debug_info.append(f"\n❌ Пользователь {username} не найден в расписании")

        # Добавим информацию о файле
        debug_info.append(f"\n📁 Путь к файлу: {os.path.abspath('schedule.xlsx')}")
        debug_info.append(f"📊 Размер таблицы: {ws.max_row}x{ws.max_column}")

        await message.answer("\n".join(debug_info))
        wb.close()
    except Exception as e:
        error_msg = f"🚫 Ошибка при отладке:\n{str(e)}"
        if 'schedule.xlsx' not in os.listdir():
            error_msg += "\n\nФайл schedule.xlsx не найден в директории!"
        await message.answer(error_msg)


@router.message(Command("check_cell"))
async def cmd_check_cell(message: Message):
    if str(message.from_user.id) not in ADMIN_USERS:
        return

    try:
        wb = openpyxl.load_workbook('schedule.xlsx')
        ws = wb.active

        # Проверяем конкретную ячейку
        next_date = datetime.now() + timedelta(days=1)
        next_day_col = next_date.day + 3  # +3 так как смены начинаются с 4-й колонки

        debug_info = [
            "🔍 Проверка ячеек:",
            f"📅 Завтрашняя дата: {next_date.strftime('%d.%m.%Y')}",
            f"📊 Номер колонки: {next_day_col}\n"
        ]

        # Проверяем все ячейки в колонке с логинами (3-я колонка)
        for row in range(1, ws.max_row + 1):
            login_cell = ws.cell(row=row, column=3)  # 3-я колонка для логинов
            shift_cell = ws.cell(row=row, column=next_day_col)

            if login_cell.value:
                debug_info.append(f"Строка {row}: {login_cell.value} -> {shift_cell.value}")
                if str(login_cell.value).lower() == 'sm_kirillts':
                    debug_info.append(f"\n🎯 Найдена целевая строка!")
                    debug_info.append(f"Координаты ячейки смены: {shift_cell.coordinate}")
                    debug_info.append(f"Тип данных: {type(shift_cell.value)}")
                    debug_info.append(f"Значение: '{shift_cell.value}'")

        await message.answer("\n".join(debug_info))
        wb.close()
    except Exception as e:
        await message.answer(f"🚫 Ошибка при проверке: {str(e)}")