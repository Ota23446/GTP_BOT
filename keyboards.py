from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.types import InlineKeyboardButton
from utils import NOTIFICATION_TYPES
from aiogram.types import CallbackQuery
from datetime import datetime, timedelta
from aiogram.utils.keyboard import InlineKeyboardBuilder
from config import MONTHS_RU, WEEKDAYS



def get_main_keyboard():
    """Создание основной клавиатуры"""
    builder = InlineKeyboardBuilder()

    buttons = [
        ("🗓 Ближайшие смены", "shifts"),
        ("⏱ Учёт рабочего времени", "worked_time"),
        ("📅 Проверить смену", "check_shift"),  # <-  Изменено
        ("⚙️ Настройки", "settings"),
        ("ℹ️ Помощь", "help")
    ]

    for text, callback_data in buttons:
        builder.button(text=text, callback_data=callback_data)

    builder.adjust(1)
    return builder.as_markup()


def get_settings_keyboard():
    """Создание клавиатуры настроек"""
    builder = InlineKeyboardBuilder()
    builder.button(text="🔔 Настройка уведомлений", callback_data="notifications")
    builder.button(text="⏰ Изменить время уведомлений", callback_data="set_time")
    builder.button(text="📊 Текущие настройки", callback_data="status")
    builder.button(text="◀️ Главное меню", callback_data="main_menu")
    builder.adjust(1)
    return builder.as_markup()


def get_notification_settings_keyboard():
    """Создание клавиатуры настроек уведомлений"""
    builder = InlineKeyboardBuilder()

    # Добавляем кнопки для каждого типа уведомлений
    for notif_key, notif_name in NOTIFICATION_TYPES.items():
        builder.button(
            text=notif_name,
            callback_data=f"toggle_{notif_key}"
        )

    # Добавляем кнопку возврата
    builder.button(
        text="◀️ Назад к настройкам",
        callback_data="back_to_settings"
    )

    # Устанавливаем расположение кнопок (по одной в ряд)
    builder.adjust(1)

    return builder.as_markup()

async def show_schedule_choice(callback: CallbackQuery):
    """Показывает выбор месяца для просмотра расписания"""
    builder = InlineKeyboardBuilder()
    current_date = datetime.now()
    current_month = current_date.month
    next_month = (current_month % 12) + 1
    builder.button(text=MONTHS_RU[current_month], callback_data=f"schedule_month:{current_month}")
    builder.button(text=MONTHS_RU[next_month], callback_data=f"schedule_month:{next_month}")
    await callback.message.edit_text("Выберите месяц:", reply_markup=builder.as_markup())


async def show_dates_for_month(callback: CallbackQuery, month: int):
    """Показывает кнопки с датами для выбранного месяца, разбивая по неделям, в столбец"""
    current_date = datetime.now()
    year = current_date.year
    if month > 12:
        year += 1

    first_day = datetime(year, month, 1)
    days_in_month = (datetime(year, (month % 12) + 1, 1) - timedelta(days=1)).day

    builder = InlineKeyboardBuilder()

    start_day = 1
    for day in range(1, days_in_month + 1):
        date_obj = datetime(year, month, day)
        if date_obj.weekday() == 6 or day == days_in_month:  # Воскресенье или последний день месяца
            end_day = day
            days_text = f"{start_day} - {end_day}"
            callback_data = f"schedule_dates:{month}:{start_day}-{end_day}"
            builder.button(text=days_text, callback_data=callback_data)
            start_day = day + 1

    builder.adjust(1)  # Вывод кнопок в столбец
    await callback.message.edit_text(f"Выберите неделю ({MONTHS_RU[month]}):", reply_markup=builder.as_markup())


async def show_specific_date_buttons(callback: CallbackQuery, month: int, days_range: str):
    """Отображает кнопки с датами в виде колонки с днями недели"""

    start_day, end_day = map(int, days_range.split('-'))
    current_date = datetime.now()
    year = current_date.year
    if month > 12:
        year += 1

    builder = InlineKeyboardBuilder()
    for day in range(start_day, end_day + 1):
        date_obj = datetime(year, month, day)
        weekday = WEEKDAYS[date_obj.strftime('%A').lower()]
        button_text = f"{day} ({weekday})"
        callback_data = f"schedule_day:{month}:{day}"
        builder.button(text=button_text, callback_data=callback_data)

    builder.adjust(1)  # Располагаем кнопки в одну колонку
    await callback.message.edit_text("Выберите дату:", reply_markup=builder.as_markup())