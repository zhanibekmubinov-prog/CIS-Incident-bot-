"""
Telegram-бот: логирует в Excel на SharePoint сообщения с командой /inc или упоминанием @бота.
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

ВАЖНО про таблицу:
    Бот раскладывает данные ПО ИМЕНАМ КОЛОНОК, а не по их порядку. Шапку читает
    перед каждой записью. Поэтому в таблицу можно добавлять свои колонки
    (например «Скважина») и переставлять их — бот оставит чужие колонки пустыми.
    Переименовывать колонки из FIELDS нельзя: данные этой колонки перестанут
    записываться (бот предупредит в логах и в /check).
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

# Колонки, которые заполняет бот. Ключ = точное имя колонки в шапке таблицы.
F_DATE = "Дата/время (Атырау)"
F_CHAT = "Чат"
F_CHAT_ID = "Chat ID"
F_AUTHOR = "Автор"
F_USERNAME = "Username"
F_TEXT = "Текст сообщения"
F_LINK = "Ссылка"
FIELDS = [F_DATE, F_CHAT, F_CHAT_ID, F_AUTHOR, F_USERNAME, F_TEXT, F_LINK]

# Значения, начинающиеся с этих символов, Excel воспринимает как формулу
# (так появились «=@Eraybek» и #ИМЯ?). Апостроф заставляет Excel хранить текст.
_FORMULA_START = ("=", "+", "-", "@")


def _norm(name) -> str:
    """Сравниваем имена колонок без учёта регистра и лишних пробелов."""
    return " ".join(str(name or "").split()).casefold()


def _as_text(value) -> str:
    s = "" if value is None else str(value)
    return "'" + s if s.startswith(_FORMULA_START) else s


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

    def headers(self) -> list:
        """Текущая шапка таблицы — читается каждый раз, т.к. её могут менять руками."""
        rng = self._call(
            "GET", f"{self._base}/workbook/tables/{self.table}/headerRowRange?$select=values"
        )
        rows = rng.get("values") or [[]]
        return [str(h) for h in rows[0]]

    @staticmethod
    def missing_fields(headers: list) -> list:
        present = {_norm(h) for h in headers}
        return [f for f in FIELDS if _norm(f) not in present]

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
        cols = self.headers()
        log.info("Таблица '%s' на месте, колонки: %s", self.table, cols)
        missing = self.missing_fields(cols)
        if missing:
            log.warning("В таблице нет колонок %s — эти данные записываться не будут", missing)

    def append_record(self, record: dict) -> list:
        """Пишет строку, раскладывая record по именам колонок. Возвращает список
        полей бота, для которых не нашлось колонки (для предупреждения)."""
        cols = self.headers()
        by_name = {_norm(k): v for k, v in record.items()}
        missing = self.missing_fields(cols)
        if len(missing) == len(FIELDS):
            raise RuntimeError(
                f"В таблице '{self.table}' не найдено ни одной колонки бота. "
                f"Шапка сейчас: {cols}. Ожидаются: {FIELDS}"
            )
        row = [_as_text(by_name.get(_norm(h), "")) for h in cols]
        url = f"{self._base}/workbook/tables/{self.table}/rows/add"
        self._call("POST", url, json={"index": None, "values": [row]})
        if missing:
            log.warning("Строка записана, но нет колонок %s", missing)
        return missing

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
    data = {
        F_DATE: msg.date.astimezone(ATYRAU).strftime("%Y-%m-%d %H:%M:%S"),
        F_CHAT: msg.chat.title or (user.full_name if user else ""),
        F_CHAT_ID: str(msg.chat.id),
        F_AUTHOR: user.full_name if user else "",
        F_USERNAME: f"@{user.username}" if user and user.username else "",
        F_TEXT: body,
        F_LINK: message_link(msg.chat, msg.message_id),
    }

    graph: GraphExcel = context.bot_data["graph"]
    try:
        graph.append_record(data)
        log.info("Записано: %s | %.60s", data[F_AUTHOR], data[F_TEXT])
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
    """Диагностика: /check покажет, видит ли бот файл, таблицу и свои колонки."""
    graph: GraphExcel = context.bot_data["graph"]
    try:
        rng = graph.table_info()
        cols = graph.headers()
        missing = graph.missing_fields(cols)
        status = (
            "все колонки бота на месте"
            if not missing
            else "⚠️ не найдены колонки: " + ", ".join(missing)
        )
        await update.effective_message.reply_text(
            "✅ Связь с SharePoint есть\n"
            f"Файл: {graph.file_name}\n"
            f"Таблица: {graph.table}\n"
            f"Диапазон: {rng.get('address', '?')}\n"
            f"Строк с данными: {max(rng.get('rowCount', 1) - 1, 0)}\n"
            f"Колонок в таблице: {len(cols)} — {status}"
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

    # Только новые сообщения. Без этого правка уже отправленного /inc в Telegram
    # записывала вторую строку-дубль (так появились две строки по сообщению 3871).
    new_only = filters.UpdateType.MESSAGE | filters.UpdateType.CHANNEL_POST

    if DEBUG_ALL:
        app.add_handler(MessageHandler(filters.ALL, debug_all), group=-1)
        log.info("Диагностика DEBUG_ALL включена")
    app.add_handler(CommandHandler("check", cmd_check, filters=new_only))
    app.add_handler(CommandHandler("inc", on_command, filters=new_only))
    app.add_handler(
        MessageHandler(
            new_only
            & (
                filters.Entity(MessageEntity.MENTION)
                | filters.CaptionEntity(MessageEntity.MENTION)
            ),
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
