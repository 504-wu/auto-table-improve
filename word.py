import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from PIL import Image, ImageDraw, ImageFont
from PIL.ExifTags import TAGS
import pillow_heif
import io
import os
from datetime import datetime

# 註冊 HEIC 格式支援
pillow_heif.register_heif_opener()

# 初始化 session_state
if "uploaded_photos_list" not in st.session_state:
    st.session_state.uploaded_photos_list = []

# -----------------------------------------
# 照片處理與日期讀取
# -----------------------------------------
def get_photo_date(image_bytes, file_name):
    """
    從照片的 EXIF 資訊讀取拍攝日期。
    傳回值: (date_str, is_from_exif)
    is_from_exif 為 True 代表是原生相機時間；False 代表是從檔名解析或預設今天。
    """
    try:
        image = Image.open(io.BytesIO(image_bytes))
        exif = image._getexif()
        if exif:
            for tag, value in exif.items():
                decoded = TAGS.get(tag, tag)
                if decoded == "DateTimeOriginal" or decoded == "DateTime":
                    # EXIF 標準格式 YYYY:MM:DD HH:MM:SS
                    date_str = value.split()[0].replace(":", "/")
                    return date_str, True
    except Exception:
        pass
    
    # 解析 YYYYMMDD 或 YYYY-MM-DD
    for fmt in ["%Y%m%d", "%Y-%m-%d", "%Y/%m/%d"]:
        try:
            clean_name = "".join(filter(lambda ch: ch.isdigit() or ch in "-/", file_name))
            if len(clean_name) >= 8:
                dt = datetime.strptime(clean_name[:10], fmt)
                return dt.strftime("%Y/%m/%d"), False
        except Exception:
            continue
            
    return datetime.today().strftime("%Y/%m/%d"), False

def resize_and_compress_image(image_bytes, date_str, is_from_exif):
    """
    處理照片尺寸與浮水印：
    - 橫式照片：維持置中裁切，填滿格子。
    - 直式照片：等比例縮放（不旋轉也不硬切成橫式）。
    - 無時間的照片：在右下角增加日期文字。
    """
    img = Image.open(io.BytesIO(image_bytes))
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
        
    img_w, img_h = img.size
    target_ratio = 8.0 / 6.15
    
    # 判斷直式（寬 < 高）或橫式
    if img_w < img_h:
        # 直式照片：等比例縮放至寬度 8cm，並裁切高度以符合 8:6.15 比例
        new_w = img_w
        new_h = int(img_w / target_ratio)
        # 縮放後的高度會高於原照片，則以高度為準調整
        if new_h > img_h:
            new_h = img_h
            new_w = int(img_h * target_ratio)
        
        # 進行縮放裁切
        offset_x = (img_w - new_w) // 2
        offset_y = (img_h - new_h) // 2
        img = img.crop((offset_x, offset_y, img_w - offset_x, img_h - offset_y))
    else:
        # 橫式照片：裁切至 8:6.15 比例
        current_ratio = img_w / img_h
        if current_ratio > target_ratio:
            new_w = int(target_ratio * img_h)
            offset = (img_w - new_w) // 2
            img = img.crop((offset, 0, img_w - offset, img_h))
        elif current_ratio < target_ratio:
            new_h = int(img_w / target_ratio)
            offset = (img_h - new_h) // 2
            img = img.crop((0, offset, img_w, img_h - offset))

    # 無時間照片，在右下角加上時間
    if not is_from_exif:
        draw = ImageDraw.Draw(img)
        w, h = img.size
        
        # 根據照片寬度決定字體大小
        font_size = max(20, int(w * 0.04))
        
        # 載入系統預設字型，若找不到則用 PIL 內建
        try:
            
            font = ImageFont.truetype("msjh.ttc", font_size)
        except IOError:
            try:
                
                font = ImageFont.truetype("Arial.ttf", font_size)
            except IOError:
                font = ImageFont.load_default()

        # 計算文字寬高以放置在右下角
        # 使用相容的 textbbox  或 textsize 
        try:
            text_w = draw.textbbox((0, 0), date_str, font=font)[2]
            text_h = draw.textbbox((0, 0), date_str, font=font)[3]
        except AttributeError:
            text_w, text_h = draw.textsize(date_str, font=font)
            
        # 設定右下角座標
        margin_x = int(w * 0.05)
        margin_y = int(h * 0.05)
        x = w - text_w - margin_x
        y = h - text_h - margin_y
        
        # 繪製文字黑邊
        border_thickness = max(1, int(font_size * 0.08))
        for dx in range(-border_thickness, border_thickness + 1):
            for dy in range(-border_thickness, border_thickness + 1):
                if dx != 0 or dy != 0:
                    draw.text((x + dx, y + dy), date_str, font=font, fill="black")
                    
        # 繪製白字
        draw.text((x, y), date_str, font=font, fill="white")

    out_io = io.BytesIO()
    img.save(out_io, format="JPEG", quality=95, dpi=(300, 300))
    out_io.seek(0)
    return out_io

# -----------------------------------------
# 固定欄寬與不拆頁
# -----------------------------------------
def set_cell_width(cell, width_cm):
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(width_cm * 567)))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def set_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

def set_row_height(row, height_cm):
    trPr = row._tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(int(height_cm * 567)))
    trHeight.set(qn('w:hRule'), 'exact')
    trPr.append(trHeight)

# -----------------------------------------
# 建立 Word 文件
# -----------------------------------------
def create_report(grouped_photos):
    doc = docx.Document()
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(1.2)
        section.bottom_margin = Cm(1.2)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)
    
    style = doc.styles['Normal']
    style.font.name = '標楷體'
    style.element.rPr.get_or_add_rFonts().set(qn('w:hint'), 'eastAsia')
    style.font.size = Pt(12)

    sorted_dates = sorted(grouped_photos.keys())
    
    for date_idx, date_str in enumerate(sorted_dates):
        photos = grouped_photos[date_str]
        total_photos = len(photos)
        
        for page_idx in range(0, max(1, total_photos), 6):
            if date_idx > 0 or page_idx > 0:
                doc.add_page_break()
                
            title = doc.add_paragraph()
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title.paragraph_format.space_before = Pt(0)
            title.paragraph_format.space_after = Pt(6)
            run = title.add_run(f"監造報表 ({date_str})")
            run.bold = True
            run.font.size = Pt(16)
            
            table = doc.add_table(rows=6, cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.style = 'Table Grid'
            
            page_photos = photos[page_idx:page_idx + 6]
            
            for block_idx in range(3):
                r_big = block_idx * 2
                r_small = block_idx * 2 + 1
                
                row_big = table.rows[r_big]
                row_small = table.rows[r_small]
                
                set_row_height(row_big, 6.6)
                set_row_height(row_small, 0.8)
                set_row_cant_split(row_big)
                set_row_cant_split(row_small)
                
                for c_idx in range(2):
                    cell_big = row_big.cells[c_idx]
                    cell_small = row_small.cells[c_idx]
                    
                    set_cell_width(cell_big, 8.0)
                    set_cell_width(cell_small, 8.0)
                    
                    cell_big.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    cell_small.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    
                    photo_pos = block_idx * 2 + c_idx
                    
                    p_para = cell_big.paragraphs[0]
                    p_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_para.paragraph_format.space_before = Pt(0)
                    p_para.paragraph_format.space_after = Pt(0)
                    
                    s_para = cell_small.paragraphs[0]
                    s_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    s_para.paragraph_format.space_before = Pt(0)
                    s_para.paragraph_format.space_after = Pt(0)
                    
                    if photo_pos < len(page_photos):
                        photo_data = page_photos[photo_pos]
                        # 輸入日期與 EXIF 來源進行縮放
                        compressed_io = resize_and_compress_image(
                            photo_data['bytes'], 
                            photo_data['date'], 
                            photo_data['is_from_exif']
                        )
                        run_img = p_para.add_run()
                        run_img.add_picture(compressed_io, width=Cm(8.0))
                        
                        s_para.add_run(photo_data.get('note', ''))
                    else:
                        p_para.add_run("")
                        s_para.add_run("") 

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# -----------------------------------------
# Streamlit 介面設計
# -----------------------------------------
st.set_page_config(page_title="監造報表", page_icon="📝", layout="wide")
st.title("📝 監造報表")
st.caption("照片自動依日期分類，會自己旋轉跟增加時間喔~~")

# 初始化
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "uploaded_photos_list" not in st.session_state:
    st.session_state.uploaded_photos_list = []
if "delete_idx" not in st.session_state:
    st.session_state.delete_idx = None

# 處理單張刪除邏輯
if st.session_state.delete_idx is not None:
    idx_to_del = st.session_state.delete_idx
    if 0 <= idx_to_del < len(st.session_state.uploaded_photos_list):
        st.session_state.uploaded_photos_list.pop(idx_to_del)
    st.session_state.delete_idx = None  # 重置刪除狀態
    st.rerun()

# 照片上傳 
uploaded_files = st.file_uploader(
    "上傳照片", 
    type=["jpg", "jpeg", "png", "heic"], 
    accept_multiple_files=True, 
    key=f"uploader_{st.session_state.uploader_key}"
)

if uploaded_files:
    for f in uploaded_files:
        file_bytes = f.read()
        # 檢查是否重複上傳
        if not any(d['name'] == f.name for d in st.session_state.uploaded_photos_list):
            # 呼叫日期解析函式
            photo_date, is_from_exif = get_photo_date(file_bytes, f.name)
            st.session_state.uploaded_photos_list.append({
                "name": f.name,
                "bytes": file_bytes,
                "date": photo_date,
                "is_from_exif": is_from_exif, 
                "note": ""
            })
    st.session_state.uploader_key += 1
    st.rerun()

# 照片管理與顯示
if st.session_state.uploaded_photos_list:
    st.success(f"目前有 {len(st.session_state.uploaded_photos_list)} 張照片")
    
    if st.button("🗑️ 一鍵刪除", type="primary"):
        st.session_state.uploaded_photos_list = []
        st.rerun()
        
    st.write("---")
    st.subheader("📸 已上傳的照片，會有預覽喔）")
    
    # 建立網格
    cols = st.columns(4)
    
    # 負責畫面渲染與刪除標記
    for idx, photo in enumerate(st.session_state.uploaded_photos_list):
        with cols[idx % 4]:
            source_tag = " [相機時間]" if photo['is_from_exif'] else " [手動/檔名解析-已蓋時間章]"
            st.image(photo['bytes'], caption=f"{photo['name']}\n{photo['date']}{source_tag}", use_container_width=True)
            
            # 點擊刪除時，將索引存入 session_state 並重新渲染
            if st.button(f"❌ 刪除", key=f"del_{idx}"):
                st.session_state.delete_idx = idx
                st.rerun()

    st.write("---")
    
    # 資料分類與報表生成
    if st.button("📝 要確定是這些照片喔，按下後圖片才會下一步喔"):
        with st.spinner("努力產生 Word 檔，敢吵我直接當機"):
            # 按按鈕才整理最新的照片進行分組，避免中途刪除導致照片遺失
            grouped_photos = {}
            for photo in st.session_state.uploaded_photos_list:
                p_date = photo['date']
                if p_date not in grouped_photos:
                    grouped_photos[p_date] = []
                grouped_photos[p_date].append(photo)
            
            #  Word 生成函式（create_report）會處理照片縮放
            doc_file = create_report(grouped_photos)
            
            st.download_button(
                label="📥 此為偷懶小幫手.docx",
                data=doc_file,
                file_name=f"監造報表_{datetime.now().strftime('%Y%m%d')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
else:
    st.info("請上傳照片再給我按，這種事還要我提醒。")


