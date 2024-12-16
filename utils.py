from typing import Dict, Any, Optional
import json
from datetime import datetime, time

# Константы для работы с уведомлениями
NOTIFICATION_TYPES = {
    'shift1': 'Первая смена',
    'shift2': 'Вторая смена',
    'shift3': 'Третья смена',
    'weekend': 'Дежурство в выходной',
    'dayoff': 'Выходной день'
}

DEFAULT_NOTIFICATION_TIME = "19:00"


class UserDataManager:
    """Класс для работы с данными пользователей"""

    def __init__(self, file_path: str = "user_data.json"):
        self.file_path = file_path

    def load_user_data(self) -> Dict[str, Any]:
        """Загрузка данных пользователей из JSON файла"""
        try:
            with open(self.file_path, 'r', encoding='utf-8') as file:
                return json.load(file)
        except FileNotFoundError:
            return {}

    def save_user_data(self, data: Dict[str, Any]) -> None:
        """Сохранение данных пользователей в JSON файл"""
        with open(self.file_path, 'w', encoding='utf-8') as file:
            json.dump(data, file, indent=4)

    def get_user_by_telegram_id(self, user_id: str) -> Optional[str]:
        """Получение username пользователя по его Telegram ID"""
        data = self.load_user_data()
        for username, user_data in data.items():
            if user_data["user_id"] == user_id:
                return username
        return None

    def get_user_settings(self, username: str) -> Optional[Dict[str, Any]]:
        """Получение настроек пользователя по username"""
        data = self.load_user_data()
        return data.get(username)

    def update_user_notifications(self, username: str, notification_type: str, status: bool) -> bool:
        """Обновление настроек уведомлений пользователя"""
        data = self.load_user_data()
        if username in data and notification_type in data[username]["notifications"]:
            data[username]["notifications"][notification_type] = status
            self.save_user_data(data)
            return True
        return False

    def update_notification_time(self, username: str, new_time: str) -> bool:
        """Обновление времени уведомлений пользователя"""
        try:
            # Проверка корректности формата времени
            datetime.strptime(new_time, "%H:%M")

            data = self.load_user_data()
            if username in data:
                data[username]["notification_time"] = new_time
                self.save_user_data(data)
                return True
            return False
        except ValueError:
            return False


def parse_time(time_str: str) -> time:
    """Преобразование строки времени в объект time"""
    try:
        hour, minute = map(int, time_str.split(':'))
        return time(hour=hour, minute=minute)
    except (ValueError, TypeError):
        return time(hour=19, minute=0)  # возвращаем время по умолчанию


def format_notification_status(status: bool) -> str:
    """Форматирование статуса уведомления для вывода пользователю"""
    return "✅ Включено" if status else "❌ Выключено"


def get_active_users_for_notification(notification_type: str) -> Dict[str, Any]:
    """Получение списка пользователей с активными уведомлениями определенного типа"""
    user_manager = UserDataManager()
    data = user_manager.load_user_data()

    active_users = {}
    for username, user_data in data.items():
        if user_data["notifications"].get(notification_type, False):
            active_users[username] = user_data

    return active_users


async def is_valid_username(username: str) -> bool:
    """Проверка валидности username"""
    return username.startswith("sm_") and username.islower()