import os
import asyncio

from flask import Flask, request
from aiogram.types import Update

from bot_core import bot, dp, WEBHOOK_PATH, WEBHOOK_URL

app = Flask(__name__)

# Флаг, чтобы не дергать set_webhook слишком часто
webhook_set = False


async def ensure_webhook():
    """
    Гарантирует, что webhook один раз установлен.
    Если что — повторный вызов не страшен, Telegram перезапишет URL.
    """
    global webhook_set
    if webhook_set:
        return
    await bot.set_webhook(WEBHOOK_URL)
    webhook_set = True


@app.route("/", methods=["GET"])
def index():
    # При первом заходе на корень выставляем webhook
    asyncio.run(ensure_webhook())
    return "Telegram bot is running."


@app.route(WEBHOOK_PATH, methods=["POST"])
def telegram_webhook():
    # На всякий случай гарантируем, что webhook установлен
    asyncio.run(ensure_webhook())

    json_data = request.get_json(force=True)
    update = Update.model_validate(json_data)
    asyncio.run(dp.feed_update(bot, update))
    return "OK"


if __name__ == "__main__":
    # Локальный запуск (для отладки, не обязателен на Render)
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8000")))