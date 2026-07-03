import os
import json
import logging
import asyncio
from contextlib import asynccontextmanager
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from fastapi import FastAPI, Request, Response, status
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters

# 1. Cấu hình log
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

# Lấy các cấu hình từ Biến môi trường (Environment Variables) để bảo mật
BOT_TOKEN = os.getenv("BOT_TOKEN")
RENDER_EXTERNAL_URL = os.getenv("RENDER_EXTERNAL_URL") # Render tự động cung cấp biến này (e.g., https://your-app.onrender.com)
PORT = int(os.getenv("PORT", 8000)) # Render tự động cấp cổng qua biến PORT

# Khởi tạo Application của Telegram
tg_application = Application.builder().token(BOT_TOKEN).build()

# --- Hàm xử lý logic chuyển đổi JSON -> Excel (Giữ nguyên) ---
def convert_json_to_styled_excel(json_filepath, excel_filepath):
    with open(json_filepath, 'r', encoding='utf-8') as f:
        data = json.load(f)
    cols = data['cols']
    students_data = data['students']
    rows = []
    for sbd, scores in students_data.items():
        row = {'SBD': sbd}
        for col_name, score in zip(cols, scores): row[col_name] = score
        rows.append(row)
    df = pd.DataFrame(rows)
    sheet_name = "DiemThi"
    with pd.ExcelWriter(excel_filepath, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    wb = openpyxl.load_workbook(excel_filepath)
    ws = wb[sheet_name]
    ws.views.sheetView[0].showGridLines = True
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_font = Font(name="Segoe UI", size=10)
    sbd_alignment = Alignment(horizontal="center", vertical="center")
    score_alignment = Alignment(horizontal="right", vertical="center")
    thin_side = Side(border_style="thin", color="D9D9D9")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    zebra_fill = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    ws.row_dimensions[1].height = 28
    for cell in ws[1]:
        cell.font = header_font; cell.fill = header_fill; cell.alignment = header_alignment; cell.border = border_all
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column), start=2):
        ws.row_dimensions[row_idx].height = 20
        row_fill = zebra_fill if row_idx % 2 == 0 else white_fill
        for col_idx, cell in enumerate(row, start=1):
            cell.font = data_font; cell.border = border_all; cell.fill = row_fill
            if col_idx == 1: cell.alignment = sbd_alignment
            else:
                cell.alignment = score_alignment
                if cell.value is not None: cell.number_format = '0.00'
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    ws.freeze_panes = "B2"
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(max_len + 3, 10)
    wb.save(excel_filepath)

# --- Telegram Handlers ---
async def start_command(update: Update, context):
    await update.message.reply_text("👋 Bot Web Service đã sẵn sàng! Hãy gửi file `.json` cấu trúc điểm thi cho tôi.")

async def handle_document(update: Update, context):
    document = update.message.document
    file_name = document.file_name
    if not file_name.endswith('.json'):
        await update.message.reply_text("❌ Vui lòng chỉ gửi file có định dạng đuôi `.json`.")
        return
    status_message = await update.message.reply_text("📥 Đã nhận file. Đang xử lý qua Web Service...")
    try:
        base_name = os.path.splitext(file_name)[0]
        json_path = f"temp_{file_name}"
        excel_path = f"{base_name}_converted.xlsx"
        tg_file = await context.bot.get_file(document.file_id)
        await tg_file.download_to_drive(json_path)
        
        # Chạy hàm đồng bộ convert bằng run_in_executor để tránh block bot
        loop = asyncio.get_running_loop()
        await loop.run_in_executor(None, convert_json_to_styled_excel, json_path, excel_path)

        await update.message.reply_document(document=open(excel_path, 'rb'), filename=f"{base_name}.xlsx", caption="🎉 Đã chuyển đổi thành công từ Web Service!")
        if os.path.exists(json_path): os.remove(json_path)
        if os.path.exists(excel_path): os.remove(excel_path)
        await status_message.delete()
    except Exception as e:
        logger.error(f"Lỗi: {str(e)}")
        await update.message.reply_text("💥 Có lỗi xảy ra khi xử lý file.")

# Đăng ký handler với telegram app
tg_application.add_handler(CommandHandler("start", start_command))
tg_application.add_handler(MessageHandler(filters.Document.ALL, handle_document))

# --- Cấu hình FastAPI & Webhook Lifespan ---
@asynccontextmanager
async def lifespan(app: FastAPI):
    # Thiết lập Webhook khi Web Service khởi động
    webhook_url = f"{RENDER_EXTERNAL_URL}/telegram-webhook"
    logger.info(f"Đang thiết lập Webhook tại: {webhook_url}")
    await tg_application.initialize()
    await tg_application.bot.set_webhook(url=webhook_url)
    await tg_application.start()
    yield
    # Hủy Webhook khi Web Service tắt
    logger.info("Đang tắt Web Service...")
    await tg_application.bot.delete_webhook()
    await tg_application.stop()
    await tg_application.shutdown()

app = FastAPI(lifespan=lifespan)

@app.get("/")
def index():
    return {"status": "Bot is running perfectly on Render!"}

@app.post("/telegram-webhook")
async def telegram_webhook(request: Request):
    """Đường dẫn tiếp nhận dữ liệu từ Telegram gửi về"""
    try:
        req_body = await request.json()
        update = Update.de_json(req_body, tg_application.bot)
        await tg_application.process_update(update)
        return Response(status_code=status.HTTP_200_OK)
    except Exception as e:
        logger.error(f"Lỗi xử lý webhook: {str(e)}")
        return Response(status_code=status.HTTP_400_BAD_REQUEST)
