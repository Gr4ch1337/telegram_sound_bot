import os
import calendar
import io
import sqlite3
from datetime import datetime, date

from aiogram import Bot, Dispatcher, F
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import StatesGroup, State
from aiogram.types import (
    Message,
    CallbackQuery,
    ReplyKeyboardMarkup,
    KeyboardButton,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    BufferedInputFile,
)

from openpyxl import Workbook

# =============== НАСТРОЙКИ ===============

# Токен читаем из переменной окружения BOT_TOKEN (на Render мы её зададим)
API_TOKEN = os.getenv("BOT_TOKEN")
if not API_TOKEN:
    raise RuntimeError("Не задана переменная окружения BOT_TOKEN")

DB_PATH = "tickets.db"

# Ограничение доступа к отчётам (если нужно — впиши сюда ID)
ADMIN_IDS: list[int] = []  # пример: [123456789]

# WEBHOOK_PATH: уникальный путь, зависящий от токена
WEBHOOK_PATH = f"/webhook/{API_TOKEN}"

# На Render будет переменная RENDER_EXTERNAL_URL с полным URL сервиса
BASE_URL = os.getenv("RENDER_EXTERNAL_URL", "http://localhost:8000")
WEBHOOK_URL = BASE_URL.rstrip("/") + WEBHOOK_PATH


# =============== КОНСТАНТЫ ===============

EMPLOYEES = [
    "Казаченкова",
    "Гвоздева",
    "Богданов",
    "Петрова",
    "Кожин",
    "Курланов",
    "Салакаев",
    "Климентьев",
    "Иванов",
    "Трембицкий",
]

VENUES = ["Бронная", "Мельников"]

PLAYS_BRONNAYA = [
    "12-я ночь",
    "Бесы",
    "Бэтмен",
    "Благо",
    "Вероника",
    "Гамлет",
    "Гордая",
    "Дачники",
    "Дядя Лёва",
    "Змея",
    "Калина Красная",
    "Капитанская дочка",
    "Молодожёны",
    "Невесты",
    "Незнайка",
    "Одна и Один",
    "Пигмалион",
    "Привидение",
    "Слава",
    "Таня",
    "Тузенбах",
    "Чайка",
    "Шкаф",
]

PLAYS_MELNIKOV = [
    "Баня",
    "Гора",
    "Дети солнца",
    "Зори",
    "Лукич",
    "Москва",
    "Снегурочка",
    "Туника",
    "Путаны",
    "Царь-девица",
]

ALL_PLAYS = PLAYS_BRONNAYA + PLAYS_MELNIKOV


# =============== СОСТОЯНИЯ FSM ===============

class Form(StatesGroup):
    employees = State()
    date = State()
    venue = State()
    play = State()
    problem = State()
    cause = State()


class Report(StatesGroup):
    date = State()   # выбор даты
    month = State()  # выбор месяца (год + месяц)


# =============== БАЗА ДАННЫХ ===============

def init_db() -> None:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS tickets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            user_id INTEGER,
            username TEXT,
            employees TEXT,
            date TEXT,
            venue TEXT,
            play TEXT,
            problem TEXT,
            cause TEXT
        )
        """
    )
    conn.commit()
    conn.close()


def insert_ticket(ticket: dict) -> None:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO tickets (
            created_at, user_id, username,
            employees, date, venue, play,
            problem, cause
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            ticket.get("created_at"),
            ticket.get("user_id"),
            ticket.get("username"),
            ", ".join(ticket.get("employees", [])),
            ticket.get("date"),
            ticket.get("venue"),
            ticket.get("play"),
            ticket.get("problem"),
            ticket.get("cause"),
        ),
    )
    conn.commit()
    conn.close()


def get_tickets(filter_date: str | None = None, filter_play: str | None = None):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    query = """
        SELECT
            id,
            created_at,
            user_id,
            username,
            employees,
            date,
            venue,
            play,
            problem,
            cause
        FROM tickets
    """
    conditions: list[str] = []
    params: list = []

    if filter_date:
        conditions.append("date = ?")
        params.append(filter_date)

    if filter_play:
        conditions.append("play = ?")
        params.append(filter_play)

    if conditions:
        query += " WHERE " + " AND ".join(conditions)

    query += " ORDER BY id"

    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    return rows


def get_tickets_by_month(year_month: str):
    """
    year_month: строка вида 'YYYY-MM'
    """
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    query = """
        SELECT
            id,
            created_at,
            user_id,
            username,
            employees,
            date,
            venue,
            play,
            problem,
            cause
        FROM tickets
        WHERE date LIKE ?
        ORDER BY id
    """
    like_pattern = f"{year_month}-%"
    cur.execute(query, (like_pattern,))
    rows = cur.fetchall()
    conn.close()
    return rows


# =============== КЛАВИАТУРЫ ===============

def build_employees_keyboard(selected: list[int]) -> InlineKeyboardMarkup:
    """
    Мультивыбор сотрудников: отмеченные помечаются ✅.
    """
    buttons: list[list[InlineKeyboardButton]] = []

    for i, name in enumerate(EMPLOYEES):
        prefix = "✅ " if i in selected else ""
        buttons.append(
            [InlineKeyboardButton(text=prefix + name, callback_data=f"EMP:{i}")]
        )

    # Кнопка "Готово" — визуально "зелёная"
    buttons.append(
        [InlineKeyboardButton(text="🟢 Готово", callback_data="EMP_DONE")]
    )
    return InlineKeyboardMarkup(inline_keyboard=buttons)


def build_venue_keyboard() -> InlineKeyboardMarkup:
    """
    Инлайн-клавиатура для выбора площадки.
    """
    rows: list[list[InlineKeyboardButton]] = []
    for v in VENUES:
        rows.append(
            [InlineKeyboardButton(text=v, callback_data=f"VENUE:{v}")]
        )
    return InlineKeyboardMarkup(inline_keyboard=rows)


def build_plays_keyboard(venue: str) -> InlineKeyboardMarkup:
    if venue == "Бронная":
        plays = PLAYS_BRONNAYA
        prefix = "BRN"
    else:
        plays = PLAYS_MELNIKOV
        prefix = "MLN"

    rows: list[list[InlineKeyboardButton]] = []
    for i, name in enumerate(plays):
        rows.append(
            [InlineKeyboardButton(text=name, callback_data=f"PLAY:{prefix}:{i}")]
        )

    return InlineKeyboardMarkup(inline_keyboard=rows)


def build_report_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="Все обращения", callback_data="RPT:ALL")],
            [InlineKeyboardButton(text="Отчёт по дате", callback_data="RPT:DATE")],
            [InlineKeyboardButton(text="Отчёт по спектаклю", callback_data="RPT:PLAY")],
            [InlineKeyboardButton(text="Отчёт по месяцу", callback_data="RPT:MONTH")],
        ]
    )


def build_report_plays_keyboard() -> InlineKeyboardMarkup:
    rows: list[list[InlineKeyboardButton]] = []
    for i, name in enumerate(ALL_PLAYS):
        rows.append(
            [InlineKeyboardButton(text=name, callback_data=f"RPLAY:{i}")]
        )
    return InlineKeyboardMarkup(inline_keyboard=rows)


def build_main_keyboard() -> ReplyKeyboardMarkup:
    """
    Reply-клавиатура, которая всегда снизу.
    """
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🚨 Хьюстон, у нас проблемы")],
            [KeyboardButton(text="📊 Отчёт")],
            [KeyboardButton(text="🏠 Главное меню")],
        ],
        resize_keyboard=True,
        one_time_keyboard=False,
    )


def build_calendar(year: int | None = None, month: int | None = None) -> InlineKeyboardMarkup:
    """
    Инлайн-календарь для выбора даты.
    """
    if year is None or month is None:
        today = date.today()
        year, month = today.year, today.month

    kb: list[list[InlineKeyboardButton]] = []

    month_name = calendar.month_name[month]
    kb.append([
        InlineKeyboardButton(text=f"{month_name} {year}", callback_data="CAL:IGNORE")
    ])

    week_days = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
    kb.append([InlineKeyboardButton(text=d, callback_data="CAL:IGNORE") for d in week_days])

    month_calendar = calendar.monthcalendar(year, month)
    for week in month_calendar:
        row: list[InlineKeyboardButton] = []
        for day_ in week:
            if day_ == 0:
                row.append(InlineKeyboardButton(text=" ", callback_data="CAL:IGNORE"))
            else:
                day_str = f"{day_:02d}"
                month_str = f"{month:02d}"
                callback = f"CAL:DAY:{year}-{month_str}-{day_str}"
                row.append(InlineKeyboardButton(text=day_str, callback_data=callback))
        kb.append(row)

    # Навигация
    if month == 1:
        prev_month = 12
        prev_year = year - 1
    else:
        prev_month = month - 1
        prev_year = year

    if month == 12:
        next_month = 1
        next_year = year + 1
    else:
        next_month = month + 1
        next_year = year

    kb.append([
        InlineKeyboardButton(
            text="<<",
            callback_data=f"CAL:PREV:{prev_year}-{prev_month:02d}"
        ),
        InlineKeyboardButton(
            text=">>",
            callback_data=f"CAL:NEXT:{next_year}-{next_month:02d}"
        ),
    ])

    return InlineKeyboardMarkup(inline_keyboard=kb)


def build_month_keyboard(year: int) -> InlineKeyboardMarkup:
    """
    Клавиатура выбора месяца для отчёта.
    """
    months = [
        ("01", "Янв"),
        ("02", "Фев"),
        ("03", "Мар"),
        ("04", "Апр"),
        ("05", "Май"),
        ("06", "Июн"),
        ("07", "Июл"),
        ("08", "Авг"),
        ("09", "Сен"),
        ("10", "Окт"),
        ("11", "Ноя"),
        ("12", "Дек"),
    ]

    rows: list[list[InlineKeyboardButton]] = []

    rows.append([
        InlineKeyboardButton(text="<<", callback_data=f"MON:PREV:{year}"),
        InlineKeyboardButton(text=str(year), callback_data="MON:IGNORE"),
        InlineKeyboardButton(text=">>", callback_data=f"MON:NEXT:{year}"),
    ])

    row: list[InlineKeyboardButton] = []
    for idx, (m_num, m_name) in enumerate(months, start=1):
        callback = f"MON:SEL:{year}-{m_num}"
        row.append(InlineKeyboardButton(text=m_name, callback_data=callback))
        if idx % 4 == 0:
            rows.append(row)
            row = []
    if row:
        rows.append(row)

    return InlineKeyboardMarkup(inline_keyboard=rows)


# =============== EXCEL ОТЧЁТЫ ===============

def tickets_to_excel(rows) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Обращения"

    headers = [
        "id",
        "created_at",
        "user_id",
        "username",
        "employees",
        "date",
        "venue",
        "play",
        "problem",
        "cause",
    ]
    ws.append(headers)

    for row in rows:
        ws.append(row)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


async def send_report_excel(message: Message, rows, description: str):
    if not rows:
        await message.answer(f"Нет обращений {description}.")
        return

    data = tickets_to_excel(rows)
    file = BufferedInputFile(data, filename="tickets_report.xlsx")
    await message.answer_document(file, caption=f"Отчёт {description}")


# =============== ХЕНДЛЕРЫ ===============

# --- Главное меню и кнопки ---

async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    kb = build_main_keyboard()
    await message.answer(
        "Привет! Я бот заявок звукового цеха.\n\n"
        "Нажми «🚨 Хьюстон, у нас проблемы», чтобы создать новое обращение.",
        reply_markup=kb,
    )


async def new_ticket_message(message: Message, state: FSMContext):
    await state.clear()
    await state.set_state(Form.employees)
    await state.update_data(selected_employees_idx=[])

    kb = build_employees_keyboard(selected=[])
    await message.answer(
        "Начинаем новое обращение.\n\n"
        "1. Выберите сотрудника/ов (можно несколько):",
        reply_markup=kb,
    )


async def main_menu_message(message: Message, state: FSMContext):
    await cmd_start(message, state)


async def report_button_message(message: Message, state: FSMContext):
    await cmd_menu(message)


# --- Сотрудники ---

async def employees_callback(call: CallbackQuery, state: FSMContext):
    await call.answer()
    data = await state.get_data()
    selected: list[int] = data.get("selected_employees_idx", [])

    if call.data == "EMP_DONE":
        if not selected:
            await call.message.answer("Пожалуйста, выберите хотя бы одного сотрудника.")
            return

        employees = [EMPLOYEES[i] for i in selected]
        await state.update_data(employees=employees)

        await state.set_state(Form.date)
        cal = build_calendar()
        await call.message.answer(
            "2. Выберите дату из календаря:",
            reply_markup=cal
        )
        return

    _, idx_str = call.data.split(":")
    idx = int(idx_str)
    if idx in selected:
        selected.remove(idx)
    else:
        selected.append(idx)

    await state.update_data(selected_employees_idx=selected)
    kb = build_employees_keyboard(selected)
    await call.message.edit_reply_markup(reply_markup=kb)


# --- Календарь при заполнении формы ---

async def calendar_form_callback(call: CallbackQuery, state: FSMContext):
    parts = call.data.split(":")
    if len(parts) < 2:
        await call.answer()
        return

    action = parts[1]

    if action == "IGNORE":
        await call.answer()
        return

    if action == "DAY":
        date_str = parts[2]
        await state.update_data(date=date_str)
        await state.set_state(Form.venue)
        await call.message.answer(
            f"Вы выбрали дату: {date_str}\n\n"
            "3. Выберите площадку:",
            reply_markup=build_venue_keyboard(),
        )
        await call.answer()
        return

    if action in ("PREV", "NEXT"):
        ym = parts[2]
        year, month = map(int, ym.split("-"))
        cal = build_calendar(year, month)
        await call.message.edit_reply_markup(reply_markup=cal)
        await call.answer()
        return


# --- Площадка ---

async def venue_callback(call: CallbackQuery, state: FSMContext):
    await call.answer()
    if not call.data.startswith("VENUE:"):
        return

    venue = call.data.split(":", 1)[1]
    if venue not in VENUES:
        return

    await state.update_data(venue=venue)
    await state.set_state(Form.play)

    kb = build_plays_keyboard(venue)
    await call.message.answer(
        "4. Выберите спектакль:",
        reply_markup=kb
    )


# --- Спектакль ---

async def play_callback(call: CallbackQuery, state: FSMContext):
    await call.answer()
    parts = call.data.split(":")
    if len(parts) != 3:
        return

    _, prefix, idx_str = parts
    idx = int(idx_str)

    if prefix == "BRN":
        play_list = PLAYS_BRONNAYA
    else:
        play_list = PLAYS_MELNIKOV

    if idx < 0 or idx >= len(play_list):
        return

    play_name = play_list[idx]
    await state.update_data(play=play_name)
    await state.set_state(Form.problem)

    await call.message.answer(
        f"Вы выбрали спектакль: {play_name}\n\n"
        "5. Опишите проблему (текстом):",
        reply_markup=None
    )


# --- Проблема ---

async def problem_entered(message: Message, state: FSMContext):
    problem_text = message.text.strip()
    await state.update_data(
        problem=problem_text,
        problem_msg_id=message.message_id,
    )
    await state.set_state(Form.cause)

    await message.answer("6. Предполагаемая причина проблемы (текстом):")


# --- Причина + сохранение тикета ---

async def cause_entered(message: Message, state: FSMContext):
    cause_text = message.text.strip()
    data = await state.get_data()

    ticket = {
        "created_at": datetime.utcnow().isoformat(),
        "user_id": message.from_user.id,
        "username": message.from_user.username,
        "employees": data.get("employees", []),
        "date": data.get("date", ""),
        "venue": data.get("venue", ""),
        "play": data.get("play", ""),
        "problem": data.get("problem", ""),
        "cause": cause_text,
    }

    insert_ticket(ticket)

    # Удаляем сообщения с проблемой и причиной, чтобы не висели простыни
    bot_obj = message.bot
    problem_msg_id = data.get("problem_msg_id")
    try:
        await bot_obj.delete_message(chat_id=message.chat.id, message_id=message.message_id)
    except Exception:
        pass
    if problem_msg_id:
        try:
            await bot_obj.delete_message(chat_id=message.chat.id, message_id=problem_msg_id)
        except Exception:
            pass

    await state.clear()

    employees_str = ", ".join(ticket["employees"])
    text = (
        "Обращение сохранено ✅\n\n"
        f"Сотрудники: {employees_str}\n"
        f"Дата: {ticket['date']}\n"
        f"Площадка: {ticket['venue']}\n"
        f"Спектакль: {ticket['play']}\n"
        f"Проблема: {ticket['problem']}\n"
        f"Причина: {ticket['cause']}\n"
    )

    kb = build_main_keyboard()
    await message.answer(text, reply_markup=kb)


# --- Команды отчётов ---

async def cmd_report_all(message: Message):
    if ADMIN_IDS and message.from_user.id not in ADMIN_IDS:
        await message.answer("У вас нет прав для просмотра отчёта.")
        return

    rows = get_tickets()
    await send_report_excel(message, rows, "по всем обращениям")


async def cmd_report_date(message: Message):
    if ADMIN_IDS and message.from_user.id not in ADMIN_IDS:
        await message.answer("У вас нет прав для просмотра отчёта.")
        return

    parts = message.text.strip().split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Укажи дату в формате YYYY-MM-DD, например:\n/report_date 2025-12-10")
        return

    filter_date = parts[1].strip()
    rows = get_tickets(filter_date=filter_date)
    await send_report_excel(message, rows, f"по дате {filter_date}")


async def cmd_report_play(message: Message):
    if ADMIN_IDS and message.from_user.id not in ADMIN_IDS:
        await message.answer("У вас нет прав для просмотра отчёта.")
        return

    parts = message.text.strip().split(maxsplit=1)
    if len(parts) < 2:
        await message.answer("Укажи название спектакля, например:\n/report_play Гамлет")
        return

    filter_play = parts[1].strip()
    rows = get_tickets(filter_play=filter_play)
    await send_report_excel(message, rows, f"по спектаклю «{filter_play}»")


async def cmd_menu(message: Message):
    if ADMIN_IDS and message.from_user.id not in ADMIN_IDS:
        await message.answer("У вас нет прав для просмотра отчётов.")
        return

    kb = build_report_menu_keyboard()
    await message.answer(
        "Меню отчётов:\n"
        "— Все обращения\n"
        "— По дате\n"
        "— По спектаклю\n"
        "— По месяцу",
        reply_markup=kb,
    )


async def report_menu_callback(call: CallbackQuery, state: FSMContext):
    if ADMIN_IDS and call.from_user.id not in ADMIN_IDS:
        await call.answer("Нет прав", show_alert=True)
        return

    _, action = call.data.split(":")

    if action == "ALL":
        rows = get_tickets()
        await send_report_excel(call.message, rows, "по всем обращениям")
        await call.answer()
        return

    if action == "DATE":
        await state.set_state(Report.date)
        cal = build_calendar()
        await call.message.answer(
            "Выберите дату для отчёта:",
            reply_markup=cal
        )
        await call.answer()
        return

    if action == "PLAY":
        kb = build_report_plays_keyboard()
        await call.message.answer(
            "Выберите спектакль для отчёта:",
            reply_markup=kb
        )
        await call.answer()
        return

    if action == "MONTH":
        await state.set_state(Report.month)
        this_year = date.today().year
        kb = build_month_keyboard(this_year)
        await call.message.answer(
            "Выберите год и месяц для отчёта:",
            reply_markup=kb
        )
        await call.answer()
        return


async def calendar_report_callback(call: CallbackQuery, state: FSMContext):
    parts = call.data.split(":")
    if len(parts) < 2:
        await call.answer()
        return

    action = parts[1]

    if action == "IGNORE":
        await call.answer()
        return

    if action == "DAY":
        filter_date = parts[2]
        rows = get_tickets(filter_date=filter_date)
        await send_report_excel(call.message, rows, f"по дате {filter_date}")
        await state.clear()
        await call.answer()
        return

    if action in ("PREV", "NEXT"):
        ym = parts[2]
        year, month = map(int, ym.split("-"))
        cal = build_calendar(year, month)
        await call.message.edit_reply_markup(reply_markup=cal)
        await call.answer()
        return


async def report_play_callback(call: CallbackQuery):
    if ADMIN_IDS and call.from_user.id not in ADMIN_IDS:
        await call.answer("Нет прав", show_alert=True)
        return

    _, idx_str = call.data.split(":")
    idx = int(idx_str)
    if idx < 0 or idx >= len(ALL_PLAYS):
        await call.answer()
        return

    play_name = ALL_PLAYS[idx]
    rows = get_tickets(filter_play=play_name)
    await send_report_excel(call.message, rows, f"по спектаклю «{play_name}»")
    await call.answer()


async def month_report_callback(call: CallbackQuery, state: FSMContext):
    parts = call.data.split(":")
    if len(parts) < 2:
        await call.answer()
        return

    action = parts[1]

    if action == "IGNORE":
        await call.answer()
        return

    if action == "SEL":
        year_month = parts[2]  # YYYY-MM
        rows = get_tickets_by_month(year_month)
        await send_report_excel(call.message, rows, f"за {year_month}")
        await state.clear()
        await call.answer()
        return

    if action in ("PREV", "NEXT"):
        year = int(parts[2])
        if action == "PREV":
            year -= 1
        else:
            year += 1
        kb = build_month_keyboard(year)
        await call.message.edit_reply_markup(reply_markup=kb)
        await call.answer()
        return


# =============== СОЗДАНИЕ BOT И DISPATCHER ===============

bot = Bot(
    token=API_TOKEN,
    default=DefaultBotProperties(parse_mode=ParseMode.HTML),
)
dp = Dispatcher()

# Инициализируем базу при старте
init_db()

# РЕГИСТРАЦИЯ ХЕНДЛЕРОВ

# Команды
dp.message.register(cmd_start, Command("start", "new"))
dp.message.register(cmd_report_all, Command("report"))
dp.message.register(cmd_report_date, Command("report_date"))
dp.message.register(cmd_report_play, Command("report_play"))
dp.message.register(cmd_menu, Command("menu", "reports_menu", "reports"))

# Reply-кнопки
dp.message.register(new_ticket_message, F.text == "🚨 Хьюстон, у нас проблемы")
dp.message.register(report_button_message, F.text == "📊 Отчёт")
dp.message.register(main_menu_message, F.text == "🏠 Главное меню")

# Опрос
dp.callback_query.register(employees_callback, Form.employees, F.data.startswith("EMP"))
dp.callback_query.register(calendar_form_callback, Form.date, F.data.startswith("CAL"))
dp.callback_query.register(venue_callback, Form.venue, F.data.startswith("VENUE:"))
dp.callback_query.register(play_callback, Form.play, F.data.startswith("PLAY"))
dp.message.register(problem_entered, Form.problem)
dp.message.register(cause_entered, Form.cause)

# Меню отчётов
dp.callback_query.register(report_menu_callback, F.data.startswith("RPT"))
dp.callback_query.register(calendar_report_callback, Report.date, F.data.startswith("CAL"))
dp.callback_query.register(report_play_callback, F.data.startswith("RPLAY"))
dp.callback_query.register(month_report_callback, Report.month, F.data.startswith("MON"))