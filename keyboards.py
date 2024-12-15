from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton

def get_main_keyboard() -> ReplyKeyboardMarkup:
    """Главное меню с основными функциями"""
    keyboard = ReplyKeyboardMarkup(resize_keyboard=True)
    keyboard.row(
        KeyboardButton("🗓 Ближайшие смены"),
        KeyboardButton("⏱ Учёт рабочего времени")
    )
    keyboard.row(
        KeyboardButton("⚙️ Настройки"),
        KeyboardButton("ℹ️ Помощь")
    )
    return keyboard

def get_settings_keyboard() -> InlineKeyboardMarkup:
    """Меню настроек"""
    keyboard = InlineKeyboardMarkup(row_width=1)
    keyboard.add(
        InlineKeyboardButton("⏰ Время уведомлений", callback_data="set_time"),
        InlineKeyboardButton("🔔 Настройка уведомлений", callback_data="notifications"),
        InlineKeyboardButton("👤 Изменить логин", callback_data="change_username"),
        InlineKeyboardButton("📊 Мои настройки", callback_data="status"),
        InlineKeyboardButton("◀️ Вернуться в меню", callback_data="main_menu")
    )
    return keyboard

def get_notification_settings_keyboard() -> InlineKeyboardMarkup:
    """Меню настройки оповещений"""
    keyboard = InlineKeyboardMarkup(row_width=1)
    keyboard.add(
        InlineKeyboardButton("🌅 Первая смена (8:00-16:30)", callback_data="toggle_shift1"),
        InlineKeyboardButton("🌞 Вторая смена (9:30-18:00)", callback_data="toggle_shift2"),
        InlineKeyboardButton("🌆 Третья смена (11:30-20:00)", callback_data="toggle_shift3"),
        InlineKeyboardButton("📆 Дежурство в выходной", callback_data="toggle_weekend_duty"),
        InlineKeyboardButton("↩️ Вернуться в настройки", callback_data="back_to_settings")
    )
    return keyboard