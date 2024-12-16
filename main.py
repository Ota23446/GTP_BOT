import asyncio
import logging
from datetime import datetime

from aiogram import Bot, Dispatcher
from aiogram.fsm.storage.memory import MemoryStorage
from apscheduler.schedulers.asyncio import AsyncIOScheduler

from config import BOT_TOKEN
from handlers import router as handlers_router
from services import send_notifications
from utils import UserDataManager
from test import router as test_router
from services import send_monday_notification, send_tuesday_notification

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Инициализация бота и диспетчера
bot = Bot(token=BOT_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)

# Регистрация роутеров
dp.include_router(handlers_router)
dp.include_router(test_router)  # Добавляем тестовый роутер

# Инициализация планировщика
scheduler = AsyncIOScheduler(timezone="Europe/Moscow")



async def on_startup():
    """Действия при запуске бота"""
    logger.info("Bot starting up...")

    # Инициализация менеджера пользовательских данных
    user_manager = UserDataManager()

    # Настройка задач планировщика для отправки уведомлений
    scheduler.add_job(
        send_notifications,
        'cron',
        minute='*',
        kwargs={'bot': bot}
    )
    scheduler.add_job(send_monday_notification, 'cron', day_of_week='mon', hour=9, minute=30, kwargs={'bot': bot})
    scheduler.add_job(send_tuesday_notification, 'cron', day_of_week='tue', hour=9, minute=30, kwargs={'bot': bot})
    # Запуск планировщика
    scheduler.start()

    logger.info("Bot started successfully!")


async def on_shutdown():
    """Действия при остановке бота"""
    logger.info("Bot shutting down...")

    # Останавливаем планировщик
    scheduler.shutdown()

    # Закрываем сессию бота
    await bot.session.close()

    logger.info("Bot stopped successfully!")


async def main():
    """Основная функция запуска бота"""
    try:
        # Регистрируем хэндлеры startup и shutdown
        dp.startup.register(on_startup)
        dp.shutdown.register(on_shutdown)

        # Запуск бота в режиме поллинга
        logger.info("Starting polling...")
        await dp.start_polling(bot)

    except Exception as e:
        logger.error(f"Critical error: {e}")
        raise
    finally:
        await dp.storage.close()


if __name__ == '__main__':
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        logger.info("Bot stopped!")
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        raise