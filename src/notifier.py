import telegram
import asyncio
from pathlib import Path
from config import TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID


async def _send(image_path, image_name):
    bot = telegram.Bot(token=TELEGRAM_BOT_TOKEN)
    with open(image_path, "rb") as img:
        await bot.send_photo(
            chat_id=TELEGRAM_CHAT_ID,
            photo=img,
            caption=image_name
        )


def send_timetable(image_path):
    asyncio.run(_send(image_path, Path(image_path).name))
    print("Timetable is sent to Telegram bot: Timetable Extractor")