import asyncio
import logging
import os
from contextlib import asynccontextmanager

import httpx
from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import CommandStart
from aiogram.types import (
    BufferedInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    MenuButtonWebApp,
    Message,
    Update,
    WebAppInfo,
)
from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from document_builder import (
    build_certificate_pdf_pillow,
    build_certificate_pptx,
    build_dealer_contract,
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN", "YOUR_BOT_TOKEN")
WEBHOOK_URL = os.getenv("WEBHOOK_URL", "https://your-service.onrender.com").rstrip("/")
WEBHOOK_PATH = f"/webhook/{BOT_TOKEN}"

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
router = Router()
dp.include_router(router)


async def keep_alive_pinger():
    await asyncio.sleep(15)
    async with httpx.AsyncClient(timeout=10.0) as client:
        while True:
            try:
                if "your-service" not in WEBHOOK_URL:
                    resp = await client.get(f"{WEBHOOK_URL}/health")
                    logger.info(f"Auto-ping status: {resp.status_code}")
            except Exception as e:
                logger.warning(f"Auto-ping error: {e}")
            await asyncio.sleep(600)


@asynccontextmanager
async def lifespan(app: FastAPI):
    if BOT_TOKEN != "YOUR_BOT_TOKEN" and "your-service" not in WEBHOOK_URL:
        await bot.set_webhook(
            url=f"{WEBHOOK_URL}{WEBHOOK_PATH}",
            drop_pending_updates=True,
        )
        await bot.set_chat_menu_button(
            menu_button=MenuButtonWebApp(
                text="МАГИЯ ДОКУМЕНТОВ",
                web_app=WebAppInfo(url=f"{WEBHOOK_URL}/webapp"),
            )
        )
        logger.info(f"Webhook установлен на {WEBHOOK_URL}{WEBHOOK_PATH}")

    ping_task = asyncio.create_task(keep_alive_pinger())
    yield
    ping_task.cancel()
    await bot.session.close()


app = FastAPI(lifespan=lifespan, title="ПК МИШКА Document Generator")
app.mount("/static", StaticFiles(directory="static"), name="static")


@router.message(CommandStart())
async def cmd_start(message: Message):
    webapp_url = f"{WEBHOOK_URL}/webapp"
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="✨ МАГИЯ ДОКУМЕНТОВ ✨",
                    web_app=WebAppInfo(url=webapp_url),
                )
            ]
        ]
    )
    await message.answer(
        "🐻 <b>Добро пожаловать в генератор документов ООО «ПК МИШКА»!</b>\n\n"
        "Здесь вы можете за несколько секунд сформировать:\n"
        "• 📄 <b>Дилерский договор</b> со всеми приложениями\n"
        "• 🏆 <b>Сертификат официального дилера</b>\n\n"
        "Нажмите кнопку ниже, чтобы открыть интерактивную форму:",
        reply_markup=kb,
        parse_mode="HTML",
    )


async def send_generated_docs(chat_id: int, payload: dict):
    doc_type = payload.get("doc_type", "both")

    await bot.send_message(
        chat_id=chat_id,
        text="⏳ <b>Генерирую ваши документы...</b> Пару секунд!",
        parse_mode="HTML",
    )

    if doc_type in ("contract", "both"):
        docx_stream = build_dealer_contract(payload)
        safe_num = payload.get("contract_number", "777").replace("/", "_")
        filename = f"Дилерский_договор_МИШКА_№{safe_num}.docx"
        await bot.send_document(
            chat_id=chat_id,
            document=BufferedInputFile(docx_stream.read(), filename=filename),
            caption=f"✅ <b>Дилерский договор № {payload.get('contract_number')}</b>\nДилер: {payload.get('dealer_name')}",
            parse_mode="HTML",
        )

    if doc_type in ("certificate", "both"):
        safe_cert = payload.get("cert_number", "2708").replace("/", "_")
        pptx_stream = build_certificate_pptx(payload)
        if pptx_stream:
            pptx_filename = f"Сертификат_МИШКА_{safe_cert}.pptx"
            await bot.send_document(
                chat_id=chat_id,
                document=BufferedInputFile(pptx_stream.read(), filename=pptx_filename),
                caption=f"🏆 <b>Сертификат дилера ({payload.get('cert_company')})</b> — исходник PPTX",
                parse_mode="HTML",
            )

        pdf_stream = build_certificate_pdf_pillow(payload)
        if pdf_stream:
            pdf_filename = f"Сертификат_МИШКА_{safe_cert}.pdf"
            await bot.send_document(
                chat_id=chat_id,
                document=BufferedInputFile(pdf_stream.read(), filename=pdf_filename),
                caption="🖨 <b>Сертификат дилера в формате PDF (готов к печати)</b>",
                parse_mode="HTML",
            )


@router.message(F.web_app_data)
async def handle_web_app_data(message: Message):
    import json
    data = json.loads(message.web_app_data.data)
    await send_generated_docs(message.chat.id, data)


@app.get("/health")
async def health_check():
    return {"status": "ok", "service": "mishka-doc-bot"}


@app.get("/webapp")
async def serve_webapp():
    return FileResponse("static/index.html")


class GenerateRequest(BaseModel):
    user_id: int
    doc_type: str
    contract_number: str = ""
    contract_date: str = ""
    dealer_name: str = ""
    dealer_signatory_gen: str = ""
    dealer_short_sign: str = ""
    dealer_address: str = ""
    dealer_inn_kpp: str = ""
    dealer_rs: str = ""
    dealer_bank: str = ""
    dealer_ks: str = ""
    dealer_bik: str = ""
    dealer_phone: str = ""
    dealer_email: str = ""
    cert_number: str = ""
    cert_company: str = ""
    cert_inn: str = ""
    cert_valid_until: str = ""


@app.post("/api/generate")
async def api_generate(req: GenerateRequest):
    try:
        await send_generated_docs(req.user_id, req.model_dump())
        return JSONResponse({"ok": True, "message": "Документы успешно отправлены в чат!"})
    except Exception as e:
        logger.exception("Ошибка при генерации документов:")
        return JSONResponse({"ok": False, "error": str(e)}, status_code=500)


@app.post(WEBHOOK_PATH)
async def telegram_webhook(request: Request):
    update_data = await request.json()
    update = Update.model_validate(update_data, context={"bot": bot})
    await dp.feed_update(bot, update)
    return {"ok": True}
