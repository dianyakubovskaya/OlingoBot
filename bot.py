"""
OlingoBot — Telegram quiz bot.

Quiz with 7 themed blocks (57 questions total).
Admin sends questions to all registered participants via /send_next_question.
Supports both multiple-choice (inline buttons) and open-ended (text input) questions.
Answers and timestamps are saved to an Excel file.
"""

import asyncio
import html
import json
import logging
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook
from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

from config import BOT_TOKEN, ADMIN_IDS, EXCEL_FILE, QUESTIONS_FILE

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

# Loaded block data from JSON
blocks: list[dict] = []

# Flat list of all questions with added runtime fields:
#   global_idx, block_name, block_idx (1-based)
all_questions: list[dict] = []

# Registered participants: {user_id: {"name": str, "username": str}}
participants: dict[int, dict] = {}

# Collected answers — rows for Excel
answers: list[dict] = []

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------

# Pointer for /send_next_question (index into all_questions)
next_question_idx: int = 0

# Timestamp when each question was sent: global_idx -> str
question_sent_times: dict[int, str] = {}

# Users who already answered a given question: global_idx -> set of user_ids
answered_users: dict[int, set[int]] = {}

# For open-ended questions: which question a user should answer next
# user_id -> global_idx
pending_open: dict[int, int] = {}

OPTION_LABELS = ["A", "B", "C", "D", "E", "F"]


# ---------------------------------------------------------------------------
# Loading questions
# ---------------------------------------------------------------------------


def load_questions() -> None:
    """Load question blocks from JSON and build a flat list."""
    global blocks, all_questions
    path = Path(QUESTIONS_FILE)
    if not path.exists():
        logger.error("Questions file %s not found!", QUESTIONS_FILE)
        return
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    blocks = data.get("blocks", [])

    idx = 0
    for bi, block in enumerate(blocks):
        for q in block["questions"]:
            q["global_idx"] = idx
            q["block_name"] = block["name"]
            q["block_idx"] = bi + 1
            all_questions.append(q)
            idx += 1

    logger.info("Loaded %d blocks, %d questions.", len(blocks), len(all_questions))


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

EXCEL_HEADERS = [
    "Имя",
    "Username",
    "User ID",
    "Блок",
    "Вопрос №",
    "Текст вопроса",
    "Ответ",
    "Вопрос отправлен (UTC)",
    "Ответ получен (UTC)",
]


def save_to_excel() -> None:
    """Write all collected answers to an Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Ответы"

    for ci, header in enumerate(EXCEL_HEADERS, 1):
        cell = ws.cell(row=1, column=ci, value=header)
        cell.font = cell.font.copy(bold=True)

    for ri, a in enumerate(answers, 2):
        ws.cell(row=ri, column=1, value=a["user_name"])
        ws.cell(row=ri, column=2, value=a["username"])
        ws.cell(row=ri, column=3, value=a["user_id"])
        ws.cell(row=ri, column=4, value=a["block_name"])
        ws.cell(row=ri, column=5, value=a["question_number"])
        ws.cell(row=ri, column=6, value=a["question_text"])
        ws.cell(row=ri, column=7, value=a["answer_text"])
        ws.cell(row=ri, column=8, value=a["sent_at"])
        ws.cell(row=ri, column=9, value=a["answered_at"])

    # Auto-width
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 60)

    wb.save(EXCEL_FILE)
    logger.info("Saved %d answers to %s.", len(answers), EXCEL_FILE)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS


def now_utc() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def ensure_participant(user) -> None:
    """Register user as participant if not yet known."""
    if user.id not in participants:
        participants[user.id] = {
            "name": user.full_name,
            "username": user.username or "",
        }


def build_question_message(q: dict) -> tuple[str, InlineKeyboardMarkup | None]:
    """Build HTML message text and optional inline keyboard for a question."""
    block = html.escape(q["block_name"])
    number = html.escape(str(q["number"]))
    text = html.escape(q["text"])

    lines = [
        f"🦉 <b>{block}</b>",
        f"Вопрос {number}:\n",
        text,
    ]

    keyboard = None
    if q["type"] == "choice":
        lines.append("")
        buttons = []
        for i, opt in enumerate(q.get("options", [])):
            lines.append(f"  {OPTION_LABELS[i]}. {html.escape(opt)}")
            buttons.append(
                InlineKeyboardButton(
                    text=OPTION_LABELS[i],
                    callback_data=f"a_{q['global_idx']}_{i}",
                )
            )
        keyboard = InlineKeyboardMarkup([buttons])
    else:
        lines.append("\n🦉 Жду твой ответ текстом. Не заставляй меня ждать.")

    return "\n".join(lines), keyboard


# ---------------------------------------------------------------------------
# Sending questions
# ---------------------------------------------------------------------------


async def send_question_to_all(
    q: dict, context: ContextTypes.DEFAULT_TYPE
) -> int:
    """Send a question to every registered participant. Returns count sent."""
    gidx = q["global_idx"]
    sent_time = now_utc()
    question_sent_times[gidx] = sent_time
    answered_users[gidx] = set()

    # For open questions, set pending state for all participants
    if q["type"] == "open":
        for uid in participants:
            pending_open[uid] = gidx

    text, keyboard = build_question_message(q)

    # Send to all participants concurrently so everyone gets the question
    # at the same time (important for fair timing in a live quiz).
    async def _send(uid: int) -> bool:
        try:
            await context.bot.send_message(
                chat_id=uid,
                text=text,
                reply_markup=keyboard,
                parse_mode="HTML",
            )
            return True
        except Exception as e:
            logger.warning("Failed to send to %d: %s", uid, e)
            return False

    results = await asyncio.gather(*[_send(uid) for uid in participants])
    return sum(results)


# ---------------------------------------------------------------------------
# Command handlers
# ---------------------------------------------------------------------------


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Register participant or greet admin."""
    user = update.effective_user
    if is_admin(user.id):
        block_list = "\n".join(
            f"  {i + 1}. {b['name']} ({len(b['questions'])} вопр.)"
            for i, b in enumerate(blocks)
        )
        await update.message.reply_text(
            f"🦉 Ух-ух! Рада тебя видеть, хозяин. "
            f"Я загрузила {len(all_questions)} вопросов и готова терроризировать участников.\n\n"
            f"Блоки:\n{block_list}\n\n"
            "Команды:\n"
            "/send_next_question — выпустить следующий вопрос на волю\n"
            "/send_question <блок> <вопрос> — отправить конкретный вопрос\n"
            "/results — скачать Excel с уликами\n"
            "/participants — посмотреть на жертв\n"
            "/reset — стереть всем память и начать сначала"
        )
        return

    ensure_participant(user)
    await update.message.reply_text(
        f"🦉 О, {user.first_name}! Ты пришёл. Я уже начала волноваться.\n\n"
        "Я — сова-викторина. Я буду присылать тебе вопросы, "
        "а ты будешь на них отвечать. Быстро.\n\n"
        "Не вздумай игнорировать меня. Я знаю, где ты живёшь. "
        "Ну, в смысле... твой Telegram ID точно знаю. 👀\n\n"
        "Жди вопросов!"
    )
    logger.info("Participant registered: %s (id=%d)", user.full_name, user.id)


async def cmd_send_next_question(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    """Send the next question in sequence to all participants (admin only)."""
    global next_question_idx

    if not is_admin(update.effective_user.id):
        return

    if not participants:
        await update.message.reply_text(
            "🦉 Тут пусто... Ни одной жертвы. "
            "Участники должны сначала написать /start, чтобы я могла за ними следить."
        )
        return

    if next_question_idx >= len(all_questions):
        await update.message.reply_text(
            "🦉 Все вопросы уже отправлены! Я выжата как лимон. "
            "Используй /reset если хочешь повторить этот кошмар."
        )
        return

    q = all_questions[next_question_idx]
    sent = await send_question_to_all(q, context)
    next_question_idx += 1

    remaining = len(all_questions) - next_question_idx
    q_type = "с вариантами" if q["type"] == "choice" else "открытый"

    await update.message.reply_text(
        f"🦉 Улетел! {q['block_name']}, вопрос {q['number']} ({q_type})\n"
        f"Доставлено жертвам: {sent}\n"
        f"Ещё в гнезде: {remaining} вопросов"
    )


async def cmd_send_question(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    """Send a specific question: /send_question <block_num> <question_num>"""
    if not is_admin(update.effective_user.id):
        return

    if len(context.args) < 2:
        await update.message.reply_text(
            "🦉 Эй, ты забыл указать что отправлять!\n\n"
            "/send_question <блок> <вопрос>\n"
            "Например: /send_question 1 3\n"
            "Или: /send_question 7 11a"
        )
        return

    try:
        block_num = int(context.args[0])
    except ValueError:
        await update.message.reply_text("🦉 Номер блока должен быть числом. Ты ведь это знаешь, правда?")
        return

    q_number = context.args[1]  # string to support "11a", "11b" etc.

    if block_num < 1 or block_num > len(blocks):
        await update.message.reply_text(
            f"🦉 Блок {block_num}? Такого не существует. У меня всего {len(blocks)}. Считать умеем?"
        )
        return

    block = blocks[block_num - 1]
    q = None
    for question in block["questions"]:
        if question["number"] == q_number:
            q = question
            break

    if not q:
        numbers = ", ".join(qq["number"] for qq in block["questions"])
        await update.message.reply_text(
            f"🦉 Вопрос «{q_number}» в блоке {block_num} ({block['name']})? Нет такого.\n"
            f"Вот что есть: {numbers}"
        )
        return

    if not participants:
        await update.message.reply_text("🦉 Некому отправлять. Ни одной живой души.")
        return

    sent = await send_question_to_all(q, context)
    q_type = "с вариантами" if q["type"] == "choice" else "открытый"
    await update.message.reply_text(
        f"🦉 Улетел! {q['block_name']}, вопрос {q['number']} ({q_type})\n"
        f"Доставлено жертвам: {sent}"
    )


async def cmd_participants(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Show the list of registered participants (admin only)."""
    if not is_admin(update.effective_user.id):
        return

    if not participants:
        await update.message.reply_text("🦉 Пока никто не пришёл. Грустно. Одиноко. Как всегда.")
        return

    lines = []
    for uid, info in participants.items():
        uname = f"@{info['username']}" if info["username"] else "—"
        lines.append(f"🐣 {info['name']} ({uname})")

    await update.message.reply_text(
        f"🦉 Мои подопечные ({len(participants)}):\n\n" + "\n".join(lines)
    )


async def cmd_reset(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Reset the quiz: clear answers and rewind the question pointer (admin only)."""
    global next_question_idx

    if not is_admin(update.effective_user.id):
        return

    old_answers = len(answers)
    answers.clear()
    question_sent_times.clear()
    answered_users.clear()
    pending_open.clear()
    next_question_idx = 0

    await update.message.reply_text(
        f"🦉 *щёлк* Память стёрта. Как будто ничего не было.\n\n"
        f"Уничтожено ответов: {old_answers}\n"
        f"Вопросов снова в гнезде: {len(all_questions)}\n"
        f"Участники ({len(participants)}) — никуда не денутся."
    )
    logger.info("Quiz reset by admin. %d answers cleared.", old_answers)


async def cmd_results(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send the Excel results file to the admin."""
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("🦉 Это не для тебя. Отойди от моего гнезда.")
        return

    if not answers:
        await update.message.reply_text("🦉 Пока пусто. Никто ещё не ответил. Может, они меня боятся?")
        return

    save_to_excel()
    path = Path(EXCEL_FILE)
    if not path.exists():
        await update.message.reply_text("🦉 Файл пропал... Кто-то украл мои улики!")
        return

    with open(path, "rb") as f:
        await update.message.reply_document(
            document=f,
            filename=EXCEL_FILE,
            caption=f"🦉 Досье на участников. {len(answers)} ответов. Я всё записала.",
        )


# ---------------------------------------------------------------------------
# Answer handlers
# ---------------------------------------------------------------------------


async def handle_choice_answer(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    """Process an inline button press (multiple-choice answer)."""
    query = update.callback_query
    await query.answer()

    user = query.from_user
    parts = query.data.split("_")  # a_{global_idx}_{option_idx}
    if len(parts) != 3:
        return

    try:
        gidx = int(parts[1])
        oidx = int(parts[2])
    except ValueError:
        return

    if gidx < 0 or gidx >= len(all_questions):
        return

    # Prevent double answers
    if user.id in answered_users.get(gidx, set()):
        await query.message.reply_text("🦉 Э нет, один ответ — один шанс. Я же не благотворительность.")
        return

    answered_users.setdefault(gidx, set()).add(user.id)

    q = all_questions[gidx]
    options = q.get("options", [])
    option_text = options[oidx] if oidx < len(options) else "?"
    answer_text = f"{OPTION_LABELS[oidx]}. {option_text}"

    ensure_participant(user)

    answers.append(
        {
            "user_id": user.id,
            "user_name": user.full_name,
            "username": user.username or "",
            "block_name": q["block_name"],
            "question_number": q["number"],
            "question_text": q["text"],
            "answer_text": answer_text,
            "sent_at": question_sent_times.get(gidx, ""),
            "answered_at": now_utc(),
        }
    )

    # Remove keyboard and confirm
    await query.edit_message_reply_markup(reply_markup=None)
    await query.message.reply_text(f"🦉 Записала: {answer_text}\nНадеюсь, ты уверен. Назад пути нет.")

    logger.info(
        "Choice answer from %s (id=%d): Q[%d] -> %s",
        user.full_name,
        user.id,
        gidx,
        OPTION_LABELS[oidx],
    )


async def handle_text_answer(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    """Process a text message as an answer to the current open-ended question."""
    user = update.effective_user

    # Ignore admin messages (they use commands)
    if is_admin(user.id):
        return

    gidx = pending_open.get(user.id)
    if gidx is None:
        # No pending question — might be a random message
        return

    # Prevent double answers
    if user.id in answered_users.get(gidx, set()):
        await update.message.reply_text("🦉 Ты уже отвечал. Я помню ВСЁ.")
        return

    answered_users.setdefault(gidx, set()).add(user.id)
    del pending_open[user.id]

    q = all_questions[gidx]
    answer_text = update.message.text.strip()

    ensure_participant(user)

    answers.append(
        {
            "user_id": user.id,
            "user_name": user.full_name,
            "username": user.username or "",
            "block_name": q["block_name"],
            "question_number": q["number"],
            "question_text": q["text"],
            "answer_text": answer_text,
            "sent_at": question_sent_times.get(gidx, ""),
            "answered_at": now_utc(),
        }
    )

    await update.message.reply_text("🦉 Ответ принят! Хороший человек. Пока что.")

    logger.info(
        "Text answer from %s (id=%d): Q[%d] -> %s",
        user.full_name,
        user.id,
        gidx,
        answer_text[:50],
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


async def main() -> None:
    load_questions()

    if BOT_TOKEN == "YOUR_BOT_TOKEN_HERE":
        logger.error(
            "Bot token not configured! Set BOT_TOKEN env variable or edit config.py."
        )
        return

    if not ADMIN_IDS:
        logger.warning(
            "No admin IDs configured. Set ADMIN_IDS env variable "
            "(comma-separated Telegram user IDs)."
        )

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("send_next_question", cmd_send_next_question))
    app.add_handler(CommandHandler("send_question", cmd_send_question))
    app.add_handler(CommandHandler("results", cmd_results))
    app.add_handler(CommandHandler("participants", cmd_participants))
    app.add_handler(CommandHandler("reset", cmd_reset))
    app.add_handler(CallbackQueryHandler(handle_choice_answer, pattern=r"^a_"))
    app.add_handler(
        MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text_answer)
    )

    logger.info("Bot started. Admin IDs: %s", ADMIN_IDS)

    async with app:
        await app.updater.start_polling(allowed_updates=Update.ALL_TYPES)
        await app.start()
        logger.info("Bot is running. Press Ctrl+C to stop.")
        try:
            await asyncio.Event().wait()
        except (KeyboardInterrupt, SystemExit):
            pass
        finally:
            await app.updater.stop()
            await app.stop()


if __name__ == "__main__":
    asyncio.run(main())
