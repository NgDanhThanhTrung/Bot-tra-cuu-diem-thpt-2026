import os
import logging
import asyncio
import glob
import ijson
import requests
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# Cấu hình log hệ thống
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', 
    level=logging.INFO
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")

# --- HÀM LÀM SẠCH HỆ THỐNG ---
def clear_system_cached_files():
    logger.info("🧹 Đang dọn dẹp file tạm...")
    count = 0
    patterns = ["temp_*", "*_converted.xlsx"]
    for pattern in patterns:
        for filepath in glob.glob(pattern):
            try:
                if os.path.exists(filepath):
                    os.remove(filepath)
                    count += 1
            except Exception as e:
                logger.error(f"❌ Lỗi xóa {filepath}: {str(e)}")
    return count

# --- TẢI FILE DUNG LƯỢNG LỚN (STREAM CHUNK) ---
def download_file_low_ram(url, dest_path):
    # Sử dụng requests stream để tải file theo từng block 64KB, tránh đầy RAM
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(dest_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=65536):
                f.write(chunk)

# --- XỬ LÝ JSON LỚN ĐÚNG 1 LƯỢT DUY NHẤT (SINGLE-PASS) ---
def convert_heavy_json_to_excel(json_filepath, excel_filepath, queue, loop):
    try:
        logger.info(f"⚡ Bắt đầu phân tích Single-Pass: {json_filepath}")
        
        wb = openpyxl.Workbook(write_only=True)
        ws = wb.create_sheet(title="DiemThi")
        ws.views.sheetView[0].showGridLines = True
        
        # Styles định dạng cấu trúc Excel đẹp
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

        cols = []
        header_written = False
        student_count = 0

        with open(json_filepath, 'rb') as f:
            parser = ijson.parse(f)
            current_sbd = None
            current_scores = []
            in_cols, in_students, in_student_scores = False, False, False

            for prefix, event, value in parser:
                # 1. Đọc danh sách cột
                if prefix == 'cols' and event == 'start_array':
                    in_cols = True
                    continue
                if in_cols:
                    if event == 'string':
                        cols.append(value)
                    elif event == 'end_array':
                        in_cols = False
                    continue

                # 2. Đọc và ghi dữ liệu học sinh tuần tự
                if prefix == 'students' and event == 'start_map':
                    in_students = True
                    if not header_written:
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
                        header_written = True
                    continue

                if in_students:
                    if prefix == 'students' and event == 'map_key':
                        current_sbd = value
                    elif prefix == f'students.{current_sbd}' and event == 'start_array':
                        in_student_scores = True
                        current_scores = []
                    elif in_student_scores:
                        if event in ('number', 'integer', 'double'):
                            current_scores.append(value)
                        elif event == 'null':
                            current_scores.append(None)
                        elif event == 'end_array':
                            in_student_scores = False
                            student_count += 1
                            
                            row_fill = zebra_fill if student_count % 2 == 0 else white_fill
                            row_cells = []
                            
                            sbd_cell = openpyxl.cell.WriteOnlyCell(ws, value=current_sbd)
                            sbd_cell.font = data_font
                            sbd_cell.alignment = sbd_alignment
                            sbd_cell.fill = row_fill
                            sbd_cell.border = border_all
                            row_cells.append(sbd_cell)
                            
                            for i in range(len(cols)):
                                score_val = current_scores[i] if i < len(current_scores) else None
                                score_cell = openpyxl.cell.WriteOnlyCell(ws, value=score_val)
                                score_cell.font = data_font
                                score_cell.alignment = score_alignment
                                score_cell.fill = row_fill
                                score_cell.border = border_all
                                if score_val is not None and isinstance(score_val, (int, float)):
                                    score_cell.number_format = '0.00'
                                row_cells.append(score_cell)
                            
                            ws.append(row_cells)
                            
                            # Cứ 5000 dòng báo cáo 1 lần để không bị lạm phát CPU
                            if student_count % 5000 == 0:
                                loop.call_soon_threadsafe(queue.put_nowait, student_count)
                                
                    elif prefix == 'students' and event == 'end_map':
                        in_students = False

        wb.save(excel_filepath)
        wb.close()
        loop.call_soon_threadsafe(queue.put_nowait, f"DONE_{student_count}")
        
    except Exception as e:
        logger.error(f"❌ Lỗi ghi Excel: {str(e)}")
        loop.call_soon_threadsafe(queue.put_nowait, None)

# --- TIẾN TRÌNH XỬ LÝ NỀN ---
async def process_file_background(chat_id, file_id, file_name, context: ContextTypes.DEFAULT_TYPE):
    base_name = os.path.splitext(file_name)[0]
    json_path = f"temp_{file_id}_{file_name}"
    excel_path = f"{base_name}_{file_id}_converted.xlsx"
    
    status_message = await context.bot.send_message(
        chat_id=chat_id, 
        text="📥 Đã tiếp nhận file! Đang streaming tải xuống đĩa cứng..."
    )
    
    try:
        # Lấy thông tin URL tải file trực tiếp từ Telegram
        tg_file = await context.bot.get_file(file_id)
        loop = asyncio.get_running_loop()
        
        # Chạy tải file ngầm tiết kiệm bộ nhớ
        await loop.run_in_executor(None, download_file_low_ram, tg_file.file_path, json_path)
        await status_message.edit_text("⏳ Đang bóc tách dữ liệu Single-Pass... Đã ghi 0 dòng.")
        
        queue = asyncio.Queue()
        convert_task = loop.run_in_executor(
            None, convert_heavy_json_to_excel, json_path, excel_path, queue, loop
        )
        
        while True:
            res = await queue.get()
            if res is None:
                raise Exception("Lỗi cấu trúc dữ liệu JSON.")
            if isinstance(res, str) and res.startswith("DONE_"):
                total_row = res.split("_")[1]
                break
            try:
                await status_message.edit_text(f"⏳ Đang chuyển đổi... Đã xử lý {res} dòng thí sinh.")
            except Exception:
                pass
                
        await convert_task
        await status_message.edit_text("📤 Đang stream upload file Excel thành phẩm...")
        
        # Đẩy file trực tiếp từ đĩa cứng trả lại cho người dùng
        with open(excel_path, 'rb') as excel_file:
            await context.bot.send_document(
                chat_id=chat_id,
                document=excel_file,
                filename=f"{base_name}.xlsx",
                caption=f"🎉 Chuyển đổi thành công!\n📊 Tổng cộng: {total_row} dòng thí sinh.\n🤖 Chạy mượt mà trên giới hạn RAM 512MB."
            )
            
    except Exception as e:
        logger.error(f"🚨 Lỗi luồng nền: {str(e)}")
        try:
            await context.bot.send_message(
                chat_id=chat_id, 
                text="💥 Lỗi: File quá lớn vượt giới hạn Bot API (Tối đa 50MB cho Bot thường) hoặc sai cấu trúc JSON."
            )
        except Exception:
            pass
    finally:
        if os.path.exists(json_path): os.remove(json_path)
        if os.path.exists(excel_path): os.remove(excel_path)
        try:
            await status_message.delete()
        except Exception:
            pass

# --- ĐIỀU HƯỚNG SỰ KIỆN BOT ---
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    loop = asyncio.get_running_loop()
    deleted_files_count = await loop.run_in_executor(None, clear_system_cached_files)
    msg = f"👋 Chào mừng bạn! Bot chạy chế độ Long Polling tối ưu RAM.\n🧹 Đã dọn sạch {deleted_files_count} file tạm.\n\nHãy gửi file `.json` để bắt đầu xử lý!"
    await update.message.reply_text(msg)

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    file_name = document.file_name
    
    if not file_name.lower().endswith('.json'):
        await update.message.reply_text("❌ Vui lòng chỉ gửi file có đuôi định dạng `.json`.")
        return
    
    # Kích hoạt hàm xử lý ngầm, không block luồng nhận tin nhắn tiếp theo
    asyncio.create_task(process_file_background(update.effective_chat.id, document.file_id, file_name, context))

def main():
    logger.info("🚀 Bot thuần đang khởi chạy vòng lặp Long Polling...")
    application = Application.builder().token(BOT_TOKEN).build()
    
    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    
    # Khởi chạy long polling, tự động xóa webhook cũ nếu có để tránh xung đột
    application.run_polling(drop_pending_updates=True)

if __name__ == '__main__':
    main()
