"""
Telegram-бот: логирует в Excel на SharePoint все сообщения, где его упомянули через @username.
Без ИИ — чистый парсинг entities. Запись через Microsoft Graph (workbook API, app-only auth).

Зависимости:  pip install -r requirements.txt

Переменные окружения (все обязательные):
    BOT_TOKEN            токен от @BotFather
    AZURE_TENANT_ID      Directory (tenant) ID из Entra ID
    AZURE_CLIENT_ID      Application (client) ID
    AZURE_CLIENT_SECRET  значение (Value) client secret
    SP_DRIVE_ID          parentReference.driveId  (вид  b!xxxx...)
    SP_ITEM_ID           id файла                 (вид  01ABCDEF...)
    SP_TABLE             имя таблицы Excel внутри файла, например  Log

Идентификаторы получаются одним запросом в Graph Explorer:
    GET /shares/{sharing-token}/driveItem?$select=id,name,webUrl,parentReference
Они не зависят от имени файла и его расположения — файл можно переименовать
или перенести в другую папку, бот продолжит писать в него.
"""

import logging
import os
import time
from datetime import timedelta, timezone

import msal
import requests
from telegram import MessageEntity, Update
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters

logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("mention-logger")

ATYRAU = timezone(timedelta(hours=5))
GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["https://graph.microsoft.com/.default"]


# ============================================================ Microsoft Graph
class GraphExcel:
    """Минимальный клиент: разрешает site/файл и добавляет строки в таблицу Excel."""

    def __init__(self) -> None:
        self._app = msal.ConfidentialClientApplication(
            client_id=os.environ["AZURE_CLIENT_ID"],
            client_credential=os.environ["AZURE_CLIENT_SECRET"],
            authority=f"https://login.microsoftonline.com/{os.environ['AZURE_TENANT_ID']}",
        )
        self.drive_id = os.environ["SP_DRIVE_ID"]
        self.item_id = os.environ["SP_ITEM_ID"]
        self.table = os.environ["SP_TABLE"]
        self.file_name = "?"

    # ---------------------------------------------------------------- низкий уровень
    def _token(self) -> str:
        # MSAL кэширует токен сам и обновляет его только когда истёк
        result = self._app.acquire_token_for_client(scopes=SCOPES)
        if "access_token" not in result:
            raise RuntimeError(
                f"Авторизация не прошла: {result.get('error')} — "
                f"{result.get('error_description', '')[:300]}"
            )
        return result["access_token"]

    def _call(self, method: str, url: str, **kwargs) -> dict:
        last = ""
        for attempt in range(4):
            resp = requests.request(
                method,
                url,
                headers={
                    "Authorization": f"Bearer {self._token()}",
                    "Content-Type": "application/json",
                },
                timeout=30,
                **kwargs,
            )
            if resp.status_code in (429, 500, 503, 504):
                pause = int(resp.headers.get("Retry-After", 2**attempt))
                log.warning("Graph %s, повтор через %s с", resp.status_code, pause)
                time.sleep(pause)
                last = f"{resp.status_code}: {resp.text[:300]}"
                continue
            if not resp.ok:
                raise RuntimeError(f"Graph {resp.status_code}: {resp.text[:400]}")
            return resp.json() if resp.content else {}
        raise RuntimeError(f"Graph недоступен после 4 попыток. Последний ответ — {last}")

    # ---------------------------------------------------------------- высокий уровень
    @property
    def _base(self) -> str:
        return f"{GRAPH}/drives/{self.drive_id}/items/{self.item_id}"

    def resolve(self) -> None:
        """Проверка при старте: доступен ли файл и есть ли в нём нужная таблица."""
        item = self._call("GET", f"{self._base}?$select=id,name,size")
        self.file_name = item.get("name", "?")
        log.info("Файл найден: %s (%s байт)", self.file_name, item.get("size"))

        tables = self._call("GET", f"{self._base}/workbook/tables?$select=name")
        names = [t["name"] for t in tables.get("value", [])]
        if self.table not in names:
            raise RuntimeError(
                f"В файле нет таблицы '{self.table}'. Найденные таблицы: "
                f"{names or 'ни одной — сделайте Ctrl+T на шапке и задайте имя'}"
            )
        log.info("Таблица '%s' на месте", self.table)

    def append_row(self, values: list) -> None:
        url = f"{self._base}/workbook/tables/{self.table}/rows/add"
        self._call("POST", url, json={"index": None, "values": [values]})

    def table_info(self) -> dict:
        return self._call("GET", f"{self._base}/workbook/tables/{self.table}/range")


# ============================================================ разбор сообщения
def _text_and_entities(message):
    return (message.text or message.caption or ""), (
        message.entities or message.caption_entities or []
    )


def is_mentioned(message, bot_username: str) -> bool:
    text, entities = _text_and_entities(message)
    target = f"@{bot_username}".lower()
    return any(
        e.type == MessageEntity.MENTION
        and text[e.offset : e.offset + e.length].lower() == target
        for e in entities
    )


def strip_mention(text: str, entities, bot_username: str) -> str:
    """Вырезает @botname, остальной текст не трогает."""
    if not text:
        return ""
    target = f"@{bot_username}".lower()
    cuts = [
        (e.offset, e.offset + e.length)
        for e in entities
        if e.type == MessageEntity.MENTION
        and text[e.offset : e.offset + e.length].lower() == target
    ]
    for start, end in sorted(cuts, reverse=True):
        text = text[:start] + text[end:]
    return " ".join(text.split())


def message_link(chat, message_id: int) -> str:
    if chat.username:
        return f"https://t.me/{chat.username}/{message_id}"
    if str(chat.id).startswith("-100"):
        return f"https://t.me/c/{str(chat.id)[4:]}/{message_id}"
    return ""


# ============================================================ хендлеры
async def record(update: Update, context: ContextTypes.DEFAULT_TYPE, body: str) -> None:
    """Единая запись строки. body — уже очищенный текст."""
    msg = update.effective_message
    user = msg.from_user
    if not body:
        await msg.reply_text(
            "Текст пустой — напишите описание после команды.\n"
            "Например: /inc Не закреплён шланг ЛВД на кусте 42"
        )
        return
    row = [
        msg.date.astimezone(ATYRAU).strftime("%Y-%m-%d %H:%M:%S"),
        msg.chat.title or (user.full_name if user else ""),
        str(msg.chat.id),
        user.full_name if user else "",
        f"@{user.username}" if user and user.username else "",
        body,
        message_link(msg.chat, msg.message_id),
    ]

    graph: GraphExcel = context.bot_data["graph"]
    try:
        graph.append_row(row)
        log.info("Записано: %s | %.60s", row[3], row[5])
        await msg.set_reaction("👌")  # тихое подтверждение вместо ответа в чат
    except Exception as exc:
        log.exception("Запись не удалась")
        await msg.reply_text(f"⚠️ Не смог записать в таблицу.\n{exc}")


async def on_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/inc описание — основной способ, работает при любом privacy mode."""
    msg = update.effective_message
    text = msg.text or msg.caption or ""
    body = " ".join(text.split(maxsplit=1)[1:]).strip()
    await record(update, context, body)


async def on_mention(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """@бот описание — альтернативный способ."""
    msg = update.effective_message
    if not msg or not is_mentioned(msg, context.bot.username):
        return
    raw, entities = _text_and_entities(msg)
    await record(update, context, strip_mention(raw, entities, context.bot.username))


async def cmd_check(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Диагностика: /check покажет, видит ли бот файл и таблицу."""
    graph: GraphExcel = context.bot_data["graph"]
    try:
        rng = graph.table_info()
        await update.effective_message.reply_text(
            "✅ Связь с SharePoint есть\n"
            f"Файл: {graph.file_name}\n"
            f"Таблица: {graph.table}\n"
            f"Диапазон: {rng.get('address', '?')}\n"
            f"Строк с данными: {max(rng.get('rowCount', 1) - 1, 0)}"
        )
    except Exception as exc:
        await update.effective_message.reply_text(f"❌ Проблема с доступом:\n{exc}")


DEBUG_ALL = os.getenv("DEBUG_ALL") == "1"


async def debug_all(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Диагностика: пишет в лог всё входящее. Включается DEBUG_ALL=1."""
    msg = update.effective_message
    if not msg:
        log.info("DEBUG: апдейт без сообщения: %s", update.to_dict())
        return
    text, entities = _text_and_entities(msg)
    log.info(
        "DEBUG: чат=%s(%s) от=%s текст=%r разметка=%s ожидаю=@%s",
        msg.chat.type,
        msg.chat.id,
        msg.from_user.username if msg.from_user else "?",
        text[:120],
        [(e.type, text[e.offset : e.offset + e.length]) for e in entities],
        context.bot.username,
    )


# ============================================================ main
def main() -> None:
    graph = GraphExcel()
    graph.resolve()  # падаем сразу при старте, если доступа нет — видно в логах Railway

    app = Application.builder().token(os.environ["BOT_TOKEN"]).build()
    app.bot_data["graph"] = graph
    if DEBUG_ALL:
        app.add_handler(MessageHandler(filters.ALL, debug_all), group=-1)
        log.info("Диагностика DEBUG_ALL включена")
    app.add_handler(CommandHandler("check", cmd_check))
    app.add_handler(CommandHandler("inc", on_command))
    app.add_handler(
        MessageHandler(
            filters.Entity(MessageEntity.MENTION)
            | filters.CaptionEntity(MessageEntity.MENTION),
            on_mention,
        )
    )

    async def show_me(application) -> None:
        me = await application.bot.get_me()
        log.info("Бот запущен как @%s (id %s)", me.username, me.id)

    app.post_init = show_me
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
