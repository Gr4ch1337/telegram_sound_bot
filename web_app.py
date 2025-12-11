import os
import asyncio

from flask import Flask, request
from aiogram.types import Update

from bot_core import bot, dp, WEBHOOK_PATH, WEBHOOK_URL

app = Flask(__name__)


@app.route("/", methods=["GET"])
def index():
    return "Telegram bot is running."


@app.before_first_request
def on_startup():
    """
    При первом запросе (например, когда ты откроешь сайт в браузере)
    выставляем webhook для Telegram.
    """
    asyncio.run(bot.set_webhook(WEBHOOK_URL))


@app.route(WEBHOOK_PATH, methods=["POST"])
def telegram_webhook():
    """
    Сюда Telegram будет слать обновления.
    """
    json_data = request.get_json(force=True)
    update = Update.model_validate(json_data)
    asyncio.run(dp.feed_update(bot, update))
    return "OK"


if __name__ == "__main__":
    # Локальный запуск (для отладки, не обязателен)
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8000")))