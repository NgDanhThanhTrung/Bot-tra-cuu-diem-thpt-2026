import os
import json
import logging
import asyncio
from contextlib import asynccontextmanager

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from fastapi import FastAPI, Request, Response, status
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters

# Cấu hình log hệ thống
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")
RENDER_EXTERNAL_URL = os.getenv("RENDER_EXTERNAL_URL") 
PORT = int(os.getenv("PORT", 8000))

tg_application = Application.builder().token(BOT_TOKEN).build()

# --- HÀM XỬ LÝ CHUYỂN ĐỔI KÈM HIỂN THỊ TIẾN TRÌNH ---
def convert_json_to_styled_excel_with_progress(json_filepath, excel_filepath, update_progress_sync_func):
    """
    update_progress_sync_func: Hàm callback đồng bộ dùng để gửi phần trăm (%) về Telegram
    """
    with open(json_filepath, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    cols = data.get('cols', [])
    students_data = data.get('students', {})
    total_students = len(students_data)
    
    # Sử dụng chế độ write_only=True giúp openpyxl ghi file siêu nhanh và tốn cực ít RAM
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet(title="DiemThi")
    ws.views.sheetView[0].showGridLines = True
    
    # Định nghĩa Styles sẵn để tái sử dụng (Tăng tốc độ xử lý cấu trúc ô)
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
    
    # 1. Tạo dòng tiêu đề (Header)
    header_row = ["SBD"] + cols
    # Trong chế độ write_only, chúng ta cần bọc từng ô vào đối tượng Cell nếu muốn định dạng trực tiếp luôn khi ghi
    header_cells = []
    for val in header_row:
        cell = openpyxl.cell.WriteOnlyCell(ws, value=val)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = border_all
        header_cells.append(cell)
    ws.append(header_cells)
    
    # Thiết lập chiều cao dòng tiêu đề (Vì write_only không hỗ trợ trực tiếp từ đầu, sẽ cấu hình sau hoặc bỏ qua để tối ưu tốc độ)
    
    # 2. Tạo các dòng dữ liệu thí sinh
    last_reported_percentage = 0
    
    for idx, (sbd, scores) in enumerate(students_data.items(), start=1):
        row_fill = zebra_fill if idx % 2 == 0 else white_fill
        row_cells = []
        
        # Ô Số báo danh (Cột 1)
        sbd_cell = openpyxl.cell.WriteOnlyCell(ws, value=sbd)
        sbd_cell.font = data_font
        sbd_cell.alignment = sbd_alignment
        sbd_cell.fill = row_fill
        sbd_cell.border = border_all
        row_cells.append(sbd_cell)
        
        # Các ô điểm số
        for i in range(len(cols)):
            score_val = scores[i] if i < len(scores) else None
            score_cell = openpyxl.cell.WriteOnlyCell(ws, value=score_val)
            score_cell.font = data_font
            score_cell.alignment = score_alignment
            score_cell.fill = row_fill
            score_cell.border = border_all
            if score_val is not None and isinstance(score_val, (int, float)):
                score_cell.number_format = '0.00'
            row_cells.append(score_cell)
            
        ws.append(row_cells)
        
        # Tính toán phần trăm và cập nhật sau mỗi 20% tiến trình (Tránh spam API Telegram quá nhiều gây nghẽn mạch)
        if total_students > 0:
            percentage = int((idx / total_students) * 100)
            if percentage >= last_reported_percentage + 20 or percentage == 100:
                last_reported_percentage = percentage
                update_progress_sync_func(percentage)

    # Chế độ write_only yêu cầu phải save để đóng luồng ghi dữ liệu
    wb.save(excel_filepath)
    wb.close()
    
    # Lưu ý: Chế độ write_only tối ưu tốc độ ghi bằng cách loại bỏ bộ nhớ đệm nên nó không hỗ trợ tính năng Tự động co độ rộng cột hay AutoFilter trực tiếp dựa trên dữ liệu động. Tuy nhiên, đổi lại tốc độ xuất file sẽ "vọt" lên gấp 5-10 lần với các file lớn!

# --- TELEGRAM HANDLERS ---
async def start_command(update: Update, context):
    await update.message.reply_text("👋 Bot chuyển đổi siêu tốc độ đã sẵn sàng! Gửi file `.json` để trải nghiệm.")

async def handle_document(update: Update, context):
    document = update.message.document
    file_name = document.file_name
    
    if not file_name.lower().endswith('.json'):
        await update.message.reply_text("❌ Vui lòng chỉ gửi file có đuôi `.json`.")
        return
        
    status_message = await update.message.reply_text("📥 Đã nhận file. Chuẩn bị khởi tạo tiến trình... [0%]")
    
    # Tạo một hàm callback lồng bên trong để cầu nối hàm đồng bộ (Excel) với hàm bất đồng bộ (Telegram)
    def update_progress_callback(percent):
        # Trực tiếp ra lệnh cho Event Loop của hệ thống chạy ngầm việc sửa tin nhắn Telegram
        asyncio.run_coroutine_threadsafe(
            status_message.edit_text(f"⏳ Đang xử lý cấu trúc dữ liệu... [{percent}%]"),
            loop
        )

    try:
        base_name = os.path.splitext(file_name)[0]
        json_path = f"temp_{file_name}"
        excel_path = f"{base_name}_converted.xlsx"
        
        tg_file = await context.bot.get_file(document.file_id)
        await tg_file.download_to_drive(json_path)
        
        loop = asyncio.get_running_loop()
        # Chạy tác vụ ghi file trong Executor thread
        await loop.run_in_executor(None, convert_json_to_styled_excel_with_progress, json_path, excel_path, update_progress_callback)
        
        # Cập nhật bước cuối
        await status_message.edit_text("📤 Đang gửi lại file Excel hoàn chỉnh...")
        
        with open(excel_path, 'rb') as excel_file:
            await update.message.reply_document(
                document=excel_file,
                filename=f"{base_name}.xlsx",
                caption="🎉 Đã xử lý xong 100%! File Excel siêu tốc của bạn đây."
            )
            
        if os.path.exists(json_path): os.remove(json_path)
        if os.path.exists(excel_path): os.remove(excel_path)
        await status_message.delete()
        
    except Exception as e:
        logger.error(f"Error: {str(e)}")
        await update.message.reply_text("💥 Đã xảy ra sự cố trong quá trình xử lý thời gian thực.")

# Đăng ký handler cho bot
tg_application.add_handler(CommandHandler("start", start_command))
tg_application.add_handler(MessageHandler(filters.Document.ALL, handle_document))

# Khởi tạo FastAPI
app = FastAPI(lifespan=asynccontextmanager(lambda app: lifespan_handler(app)))

async def lifespan_handler(app):
    webhook_url = f"{RENDER_EXTERNAL_URL}/telegram-webhook"
    logger.info(f"Setting webhook to: {webhook_url}")
    await tg_application.initialize()
    await tg_application.bot.set_webhook(url=webhook_url)
    await tg_application.start()
    yield
    await tg_application.bot.delete_webhook()
    await tg_application.stop()
    await tg_application.shutdown()

@app.get("/")
def health_check():
    return {"status": "ok"}

@app.post("/telegram-webhook")
async def telegram_webhook(request: Request):
    try:
        req_body = await request.json()
        update = Update.de_json(req_body, tg_application.bot)
        await tg_application.process_update(update)
        return Response(status_code=status.HTTP_200_OK)
    except Exception as e:
        logger.error(f"Webhook error: {str(e)}")
        return Response(status_code=status.HTTP_400_BAD_REQUEST)
