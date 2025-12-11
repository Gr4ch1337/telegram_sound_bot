import os
import asyncio

from flask import Flask, request
from aiogram.types import Update

from bot_core import bot, dp, WEBHOOK_PATH, WEBHOOK_URL

app = Flask(__name__)

# Глобальный event loop для всего приложения
loop = asyncio.new_event_loop()
asyncio.set_event_loop(loop)

# Флаг, чтобы не дёргать set_webhook лишний раз
webhook_set = False


async def ensure_webhook():
    """
    Гарантирует, что webhook один раз установлен.
    """
    global webhook_set
    if webhook_set:
        return
    await bot.set_webhook(WEBHOOK_URL)
    webhook_set = True


async def process_update(update: Update):
    """
    Обработка апдейта от Telegram.
    """
    await ensure_webhook()
    await dp.feed_update(bot, update)


@app.route("/", methods=["GET"])
def index():
    # При первом заходе на корень выставляем webhook
    loop.run_until_complete(ensure_webhook())
    return "Telegram bot is running."


@app.route(WEBHOOK_PATH, methods=["POST"])
def telegram_webhook():
    json_data = request.get_json(force=True)
    update = Update.model_validate(json_data)
    loop.run_until_complete(process_update(update))
    return "OK"


if __name__ == "__main__":
    # Локальный запуск (для отладки)
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8000")))