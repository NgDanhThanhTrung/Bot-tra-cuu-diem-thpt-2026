import os
import json
import logging
import asyncio
import glob
from contextlib import asynccontextmanager

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from fastapi import FastAPI, Request, Response, status
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters

# Cấu hình log để theo dõi trên Render Dashboard
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', 
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Đọc cấu hình từ Biến môi trường (Environment Variables) của Render
BOT_TOKEN = os.getenv("BOT_TOKEN")
RENDER_EXTERNAL_URL = os.getenv("RENDER_EXTERNAL_URL") 
PORT = int(os.getenv("PORT", 8000))

# Khởi tạo Bot Telegram
tg_application = Application.builder().token(BOT_TOKEN).build()

# --- HÀM KHẨN CẤP: XÓA SẠCH DỮ LIỆU TẠM TRÊN HỆ THỐNG ---
def clear_system_cached_files():
    logger.info("🧹 Đang kích hoạt lệnh quét dọn hệ thống...")
    count = 0
    patterns = ["temp_*", "*_converted.xlsx"]
    for pattern in patterns:
        for filepath in glob.glob(pattern):
            try:
                if os.path.exists(filepath):
                    os.remove(filepath)
                    logger.info(f"🗑️ Đã xóa file rác: {filepath}")
                    count += 1
            except Exception as e:
                logger.error(f"❌ Không thể xóa file {filepath}: {str(e)}")
    return count

# --- HÀM XỬ LÝ CHUYỂN ĐỔI EXCEL SIÊU TỐC ---
def convert_json_to_styled_excel_heavy(json_filepath, excel_filepath, queue, loop):
    logger.info(f"🔄 Bắt đầu xử lý file: {json_filepath}")
    
    with open(json_filepath, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    cols = data.get('cols', [])
    students_data = data.get('students', {})
    total_students = len(students_data)
    
    # Ghi thẳng xuống đĩa (write_only) giúp RAM trống hoàn toàn, chạy siêu mượt
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet(title="DiemThi")
    ws.views.sheetView[0].showGridLines = True
    
    # Định nghĩa bảng Styles cao cấp
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
    
    # 1. Ghi Header
    header_row = ["SBD"] + cols
    header_cells = []
    for val in header_row:
        cell = openpyxl.cell.WriteOnlyCell(ws, value=val)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = border_all
        header_cells.append(cell)
    ws.append(header_cells)
    
    # 2. Ghi Rows & Tính toán tiến trình theo %
    last_reported_percentage = 0
    
    for idx, (sbd, scores) in enumerate(students_data.items(), start=1):
        row_fill = zebra_fill if idx % 2 == 0 else white_fill
        row_cells = []
        
        # Ô Số báo danh
        sbd_cell = openpyxl.cell.WriteOnlyCell(ws, value=sbd)
        sbd_cell.font = data_font
        sbd_cell.alignment = sbd_alignment
        sbd_cell.fill = row_fill
        sbd_cell.border = border_all
        row_cells.append(sbd_cell)
        
        # Ô điểm số
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
        
        # Đẩy tiến độ về Telegram mỗi khi tăng thêm 10%
        if total_students > 0:
            percentage = int((idx / total_students) * 100)
            if percentage >= last_reported_percentage + 10 or percentage == 100:
                last_reported_percentage = percentage
                loop.call_soon_threadsafe(queue.put_nowait, percentage)

    wb.save(excel_filepath)
    wb.close()
    
    loop.call_soon_threadsafe(queue.put_nowait, None)
    logger.info("✨ Ghi file Excel thành công!")

# --- XỬ LÝ LỆNH TELEGRAM ---
async def start_command(update: Update, context):
    """Lệnh /start: Quét dọn toàn bộ dữ liệu tạm cũ để làm mới hoàn toàn hệ thống"""
    loop = asyncio.get_running_loop()
    deleted_files_count = await loop.run_in_executor(None, clear_system_cached_files)
    
    msg = "👋 Hệ thống Web Service đã được làm mới hoàn toàn!\n"
    if deleted_files_count > 0:
        msg += f"🧹 Đã giải phóng tận gốc {deleted_files_count} tệp tin tạm cũ.\n"
    else:
        msg += "✨ Hệ thống sạch 100%, RAM trống hoàn toàn.\n"
        
    msg += "\nHãy gửi file `.json` để bắt đầu chuyển hóa siêu tốc và theo dõi tiến độ thời gian thực!"
    await update.message.reply_text(msg)

async def handle_document(update: Update, context):
    document = update.message.document
    file_name = document.file_name
    
    if not file_name.lower().endswith('.json'):
        await update.message.reply_text("❌ Vui lòng chỉ gửi file có đuôi `.json`.")
        return
        
    status_message = await update.message.reply_text("📥 Đã nhận file. Đang chuẩn bị phân tích... [0%]")
    
    base_name = os.path.splitext(file_name)[0]
    json_path = f"temp_{file_name}"
    excel_path = f"{base_name}_converted.xlsx"
    
    try:
        tg_file = await context.bot.get_file(document.file_id)
        await tg_file.download_to_drive(json_path)
        
        queue = asyncio.Queue()
        loop = asyncio.get_running_loop()
        
        # Chạy ngầm trong Thread riêng để tránh đơ FastAPI
        convert_task = loop.run_in_executor(
            None, convert_json_to_styled_excel_heavy, json_path, excel_path, queue, loop
        )
        
        # Luồng lắng nghe hàng đợi cập nhật % liên tục
        while True:
            percentage = await queue.get()
            if percentage is None:
                break
            try:
                await status_message.edit_text(f"⏳ Đang chuyển hóa dữ liệu... [{percentage}%]")
            except Exception:
                pass
        
        await convert_task
        await status_message.edit_text("📤 Đã hoàn thành 100%! Đang gửi lại file Excel...")
        
        with open(excel_path, 'rb') as excel_file:
            await update.message.reply_document(
                document=excel_file,
                filename=f"{base_name}.xlsx",
                caption="🎉 Chuyển đổi siêu tốc thành công bằng Web Service!"
            )
            
    except Exception as e:
        logger.error(f"🚨 Lỗi: {str(e)}")
        await update.message.reply_text("💥 Đã xảy ra lỗi trong quá trình xử lý file.")
        
    finally:
        # Tự động dọn dẹp sau khi gửi xong file hiện tại
        if os.path.exists(json_path): os.remove(json_path)
        if os.path.exists(excel_path): os.remove(excel_path)
        try:
            await status_message.delete()
        except Exception:
            pass

tg_application.add_handler(CommandHandler("start", start_command))
tg_application.add_handler(filters.Document.ALL & MessageHandler(filters.Document.ALL, handle_document))

# --- CẤU HÌNH WEBHOOK LIFESPAN CHO FASTAPI ---
@asynccontextmanager
async def lifespan(app: FastAPI):
    webhook_url = f"{RENDER_EXTERNAL_URL}/telegram-webhook"
    logger.info(f"🚀 Đang kết nối Webhook: {webhook_url}")
    await tg_application.initialize()
    await tg_application.bot.set_webhook(url=webhook_url)
    await tg_application.start()
    yield
    logger.info("🛑 Đang đóng Web Service...")
    await tg_application.bot.delete_webhook()
    await tg_application.stop()
    await tg_application.shutdown()

app = FastAPI(lifespan=lifespan)

@app.get("/")
def health_check():
    return {"status": "online"}

@app.post("/telegram-webhook")
async def telegram_webhook(request: Request):
    try:
        req_body = await request.json()
        update = Update.de_json(req_body, tg_application.bot)
        await tg_application.process_update(update)
        return Response(status_code=status.HTTP_200_OK)
    except Exception as e:
        logger.error(f"🚨 Webhook Error: {str(e)}")
        return Response(status_code=status.HTTP_400_BAD_REQUEST)
