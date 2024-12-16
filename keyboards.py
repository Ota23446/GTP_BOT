from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.types import InlineKeyboardButton
from utils import NOTIFICATION_TYPES


def get_main_keyboard():
    """Создание основной клавиатуры"""
    builder = InlineKeyboardBuilder()

    buttons = [
        ("🗓 Ближайшие смены", "shifts"),
        ("⏱ Учёт рабочего времени", "worked_time"),
        ("⚙️ Настройки", "settings"),
        ("ℹ️ Помощь", "help")
    ]

    for text, callback_data in buttons:
        builder.button(
            text=text,
            callback_data=callback_data
        )

    builder.adjust(1)
    return builder.as_markup()


def get_settings_keyboard():
    """Создание клавиатуры настроек"""
    builder = InlineKeyboardBuilder()
    builder.button(text="🔔 Настройка уведомлений", callback_data="notifications")
    builder.button(text="⏰ Изменить время уведомлений", callback_data="set_time")
    builder.button(text="👤 Изменить username", callback_data="change_username")
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