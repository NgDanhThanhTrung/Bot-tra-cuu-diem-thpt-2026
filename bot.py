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

# Cấu hình log hệ thống
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', 
    level=logging.INFO
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")
RENDER_EXTERNAL_URL = os.getenv("RENDER_EXTERNAL_URL") 
PORT = int(os.getenv("PORT", 8000))

tg_application = Application.builder().token(BOT_TOKEN).build()

# --- HÀM LÀM SẠCH HỆ THỐNG ---
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

# --- HÀM CHUYỂN ĐỔI FILE CORE (CHẠY TRONG THREAD RIÊNG) ---
def convert_json_to_styled_excel_heavy(json_filepath, excel_filepath, queue, loop):
    try:
        logger.info(f"🔄 Bắt đầu phân tích file: {json_filepath}")
        with open(json_filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        cols = data.get('cols', [])
        students_data = data.get('students', {})
        total_students = len(students_data)
        
        wb = openpyxl.Workbook(write_only=True)
        ws = wb.create_sheet(title="DiemThi")
        ws.views.sheetView[0].showGridLines = True
        
        # Styles tiêu chuẩn cao
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
        
        last_reported_percentage = 0
        for idx, (sbd, scores) in enumerate(students_data.items(), start=1):
            row_fill = zebra_fill if idx % 2 == 0 else white_fill
            row_cells = []
            
            sbd_cell = openpyxl.cell.WriteOnlyCell(ws, value=sbd)
            sbd_cell.font = data_font
            sbd_cell.alignment = sbd_alignment
            sbd_cell.fill = row_fill
            sbd_cell.border = border_all
            row_cells.append(sbd_cell)
            
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
            
            if total_students > 0:
                percentage = int((idx / total_students) * 100)
                if percentage >= last_reported_percentage + 10 or percentage == 100:
                    last_reported_percentage = percentage
                    loop.call_soon_threadsafe(queue.put_nowait, percentage)

        wb.save(excel_filepath)
        wb.close()
        logger.info(f"✨ Ghi xong file xuất ra: {excel_filepath}")
    except Exception as e:
        logger.error(f"❌ Lỗi xuất Excel: {str(e)}")
    finally:
        loop.call_soon_threadsafe(queue.put_nowait, None)

# --- TIẾN TRÌNH NỀN: XỬ LÝ BACKGROUND KHÔNG LÀM ĐƠ WEBHOOK ---
async def process_file_background(chat_id, file_id, file_name):
    base_name = os.path.splitext(file_name)[0]
    json_path = f"temp_{file_id}_{file_name}"
    excel_path = f"{base_name}_{file_id}_converted.xlsx"
    
    # Tạo tin nhắn trạng thái ban đầu
    status_message = await tg_application.bot.send_message(
        chat_id=chat_id, 
        text="📥 Đã nhận file vào hàng đợi thành công! Đang tải xuống... [0%]"
    )
    
    try:
        # 1. Tải file về đĩa cứng
        tg_file = await tg_application.bot.get_file(file_id)
        await tg_file.download_to_drive(json_path)
        
        queue = asyncio.Queue()
        loop = asyncio.get_running_loop()
        
        # 2. Đẩy vào luồng chạy ngầm độc lập
        convert_task = loop.run_in_executor(
            None, convert_json_to_styled_excel_heavy, json_path, excel_path, queue, loop
        )
        
        # 3. Lắng nghe % tiến độ để cập nhật lên đoạn chat
        while True:
            percentage = await queue.get()
            if percentage is None:
                break
            try:
                await status_message.edit_text(f"⏳ Đang chuyển hóa dữ liệu... [{percentage}%]")
            except Exception:
                pass
                
        await convert_task
        await status_message.edit_text("📤 Đã xử lý xong 100%! Đang chuẩn bị chuyển file trả bạn...")
        
        # 4. Gửi trả file thành phẩm Excel
        with open(excel_path, 'rb') as excel_file:
            await tg_application.bot.send_document(
                chat_id=chat_id,
                document=excel_file,
                filename=f"{base_name}.xlsx",
                caption="🎉 Đã chuyển đổi thành công sang định dạng .xlsx chuẩn!"
            )
            
    except Exception as e:
        logger.error(f"🚨 Lỗi tiến trình ngầm: {str(e)}")
        try:
            await tg_application.bot.send_message(chat_id=chat_id, text="💥 Đã xảy ra lỗi không mong muốn trong lúc chuyển file.")
        except Exception:
            pass
    finally:
        # Tự dọn dẹp file tạm ngay lập tức sau khi hoàn thành/lỗi
        if os.path.exists(json_path): os.remove(json_path)
        if os.path.exists(excel_path): os.remove(excel_path)
        try:
            await status_message.delete()
        except Exception:
            pass

# --- ĐIỀU HƯỚNG LỆNH TELEGRAM ---
async def start_command(update: Update, context):
    loop = asyncio.get_running_loop()
    deleted_files_count = await loop.run_in_executor(None, clear_system_cached_files)
    
    msg = "👋 Hệ thống Web Service đã được làm mới hoàn toàn!\n"
    if deleted_files_count > 0:
        msg += f"🧹 Đã giải phóng sạch sẽ {deleted_files_count} file tạm bị kẹt cũ.\n"
    else:
        msg += "✨ Không có tệp rác nào, hệ thống sạch 100%.\n"
    msg += "\nHãy gửi file `.json` để tôi bắt đầu chuyển hóa siêu tốc ở chế độ nền (Background)!"
    await update.message.reply_text(msg)

async def handle_document(update: Update, context):
    document = update.message.document
    file_name = document.file_name
    
    if not file_name.lower().endswith('.json'):
        await update.message.reply_text("❌ Vui lòng chỉ gửi file có đuôi `.json`.")
        return
    
    # 💥 ĐIỂM CỐT LÕI: Tạo Task chạy độc lập ở nền và nhả luồng ngay để phản hồi Webhook lập tức!
    asyncio.create_task(process_file_background(update.effective_chat.id, document.file_id, file_name))

tg_application.add_handler(CommandHandler("start", start_command))
tg_application.add_handler(MessageHandler(filters.Document.ALL, handle_document))

# --- LIFESPAN WEBHOOK FASTAPI ---
@asynccontextmanager
async def lifespan(app: FastAPI):
    webhook_url = f"{RENDER_EXTERNAL_URL}/telegram-webhook"
    logger.info(f"🚀 Đang kết nối Webhook bảo mật: {webhook_url}")
    await tg_application.initialize()
    await tg_application.bot.set_webhook(url=webhook_url, drop_pending_updates=True) # Xóa hết tin nhắn kẹt cũ
    await tg_application.start()
    yield
    logger.info("🛑 Đang đóng kết nối Web Service...")
    await tg_application.bot.delete_webhook()
    await tg_application.stop()
    await tg_application.shutdown()

app = FastAPI(lifespan=lifespan)

@app.get("/")
def health_check():
    return {"status": "online", "mode": "async_background_task"}

@app.post("/telegram-webhook")
async def telegram_webhook(request: Request):
    try:
        req_body = await request.json()
        update = Update.de_json(req_body, tg_application.bot)
        # Không dùng await tg_application.process_update để Webhook phản hồi 200 ngay lập tức
        asyncio.create_task(tg_application.process_update(update))
        return Response(status_code=status.HTTP_200_OK)
    except Exception as e:
        logger.error(f"🚨 Webhook Error: {str(e)}")
        return Response(status_code=status.HTTP_400_BAD_REQUEST)
