import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from PIL import Image
from PIL.ExifTags import TAGS
import pillow_heif
import io
import os
from datetime import datetime

# 註冊 HEIC 格式支援
pillow_heif.register_heif_opener()

# -----------------------------------------
# 照片處理與日期讀取
# -----------------------------------------
def get_photo_date(image_bytes, file_name):
    """從照片的 EXIF 資訊讀取拍攝日期，若無則嘗試檔名，皆無則採今天"""
    try:
        image = Image.open(io.BytesIO(image_bytes))
        exif = image._getexif()
        if exif:
            for tag, value in exif.items():
                decoded = TAGS.get(tag, tag)
                if decoded == "DateTimeOriginal" or decoded == "DateTime":
                    # EXIF 標準格式 YYYY:MM:DD HH:MM:SS
                    date_str = value.split()[0].replace(":", "/")
                    return date_str
    except Exception:
        pass
    
    # 解析 YYYYMMDD 或 YYYY-MM-DD
    for fmt in ["%Y%m%d", "%Y-%m-%d", "%Y/%m/%d"]:
        try:
            # 比對並取前面部分
            clean_name = "".join(filter(lambda ch: ch.isdigit() or ch in "-/", file_name))
            if len(clean_name) >= 8:
                dt = datetime.strptime(clean_name[:10], fmt)
                return dt.strftime("%Y/%m/%d")
        except Exception:
            continue
            
    # 預設日期
    return datetime.today().strftime("%Y/%m/%d")

def resize_and_compress_image(image_bytes):
    """
    限制照片尺寸比例：寬度 8.0cm、高度 6.15cm。
    採用置中裁切保有需求比例，並使用 300 DPI 確保不模糊。
    """
    img = Image.open(io.BytesIO(image_bytes))
    
    # 轉 RGB 避免相容性問題
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
        
    # 寬高比例 = 8.0 / 6.15 
    target_ratio = 8.0 / 6.15
    img_w, img_h = img.size
    current_ratio = img_w / img_h
    
    # 置中裁切，讓照片不變形且填滿大格子
    if current_ratio > target_ratio:
        # 裁切左右
        new_w = int(target_ratio * img_h)
        offset = (img_w - new_w) // 2
        img = img.crop((offset, 0, img_w - offset, img_h))
    elif current_ratio < target_ratio:
        # 裁切上下
        new_h = int(img_w / target_ratio)
        offset = (img_h - new_h) // 2
        img = img.crop((0, offset, img_w, img_h - offset))
        
    # 指定 300 DPI 高印刷畫質
    out_io = io.BytesIO()
    img.save(out_io, format="JPEG", quality=95, dpi=(300, 300))
    out_io.seek(0)
    return out_io


# -----------------------------------------
# 固定欄寬與不拆頁
# -----------------------------------------
def set_cell_width(cell, width_cm):
    """固定單一儲存格寬度"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(width_cm * 567))) # 1 cm = 567 twips
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def set_row_cant_split(row):
    """設定表格列不列開 (CantSplit)"""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

def set_row_height(row, height_cm):
    """設定表格列固定高度"""
    trPr = row._tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(int(height_cm * 567)))
    trHeight.set(qn('w:hRule'), 'exact') # exact 代表固定高度
    trPr.append(trHeight)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """設定儲存格內邊距"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

# -----------------------------------------
# 建立 Word 文件
# -----------------------------------------
def create_report(grouped_photos):
    doc = docx.Document()
    
    # 設定頁邊距、不溢出
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(1.2)
        section.bottom_margin = Cm(1.2)
        section.left_margin = Cm(2.2)
        section.right_margin = Cm(2.2)
    
    # 預設字型為標楷體
    style = doc.styles['Normal']
    style.font.name = '標楷體'
    style.element.rPr.get_or_add_rFonts().set(qn('w:hint'), 'eastAsia')
    style.font.size = Pt(12)

    sorted_dates = sorted(grouped_photos.keys())
    
    for date_idx, date_str in enumerate(sorted_dates):
        photos = grouped_photos[date_str]
        total_photos = len(photos)
        
        # 6 張照片為一頁
        for page_idx in range(0, max(1, total_photos), 6):
            if date_idx > 0 or page_idx > 0:
                doc.add_page_break() # 不同日期或超過6張時建立新頁面
                
            # 頁面標題
            title = doc.add_paragraph()
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # 移除多餘段落間距
            title.paragraph_format.space_before = Pt(0)
            title.paragraph_format.space_after = Pt(6)
            run = title.add_run(f"監造報表")
            run.bold = True
            run.font.size = Pt(16)
            
            # 建立 6 列 x 2 欄的表格
            table = doc.add_table(rows=6, cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.style = 'Table Grid'
            
            # 當前這頁插入的照片（最多6張）
            page_photos = photos[page_idx:page_idx + 6]
            
            # 建立 2 欄並排，排序為大格-小格-..以此類推
            for block_idx in range(3):
                r_big = block_idx * 2      # 0, 2, 4 列是大格
                r_small = block_idx * 2 + 1  # 1, 3, 5 列是小格
                
                row_big = table.rows[r_big]
                row_small = table.rows[r_small]
                
                # 設定精確高度、防網頁裂開
                set_row_height(row_big, 6.6)
                set_row_height(row_small, 0.8)
                set_row_cant_split(row_big)
                set_row_cant_split(row_small)
                
                # 處理左欄 (c=0) 與 右欄 (c=1)
                for c_idx in range(2):
                    cell_big = row_big.cells[c_idx]
                    cell_small = row_small.cells[c_idx]
                    
                    # 固定寬度 8cm
                    set_cell_width(cell_big, 8.0)
                    set_cell_width(cell_small, 8.0)
                    
                    # 垂直置中
                    cell_big.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    cell_small.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    
                    # 順序：左上(0)->右上(1)->左中(2)->右中(3)->左下(4)->右下(5)
                    photo_pos = block_idx * 2 + c_idx
                    
                    # 填入大格（照片）
                    p_para = cell_big.paragraphs[0]
                    p_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_para.paragraph_format.space_before = Pt(0)
                    p_para.paragraph_format.space_after = Pt(0)
                    
                    if photo_pos < len(page_photos):
                        p_bytes = page_photos[photo_pos]
                        compressed_io = resize_and_compress_image(p_bytes)
                        run_img = p_para.add_run()
                        run_img.add_picture(compressed_io, width=Cm(8.0))
                    else:
                        # 沒照片時，留空並維持預設段落置中
                        p_para.add_run("")
                        
                    # 填入小格 (自訂文字)
                    s_para = cell_small.paragraphs[0]
                    s_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    s_para.paragraph_format.space_before = Pt(0)
                    s_para.paragraph_format.space_after = Pt(0)
                    s_para.add_run("") 

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# -----------------------------------------
# Streamlit 介面設計
# -----------------------------------------
st.set_page_config(page_title="監造報表", page_icon="📝")

st.title("📝 監造報表")
st.caption("上傳照片自動依日期分類，並導出符合規範的 Word 報表。")

# 上傳照片按鈕
uploaded_files = st.file_uploader(
    "上傳照片", 
    type=["jpg", "jpeg", "png", "heic"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"成功上傳 {len(uploaded_files)} 張照片，整理中不要急")
    
    # 依照日期分類
    grouped_photos = {}
    
    for f in uploaded_files:
        file_bytes = f.read()
        # 讀取日期
        photo_date = get_photo_date(file_bytes, f.name)
        
        if photo_date not in grouped_photos:
            grouped_photos[photo_date] = []
        grouped_photos[photo_date].append(file_bytes)
        
    st.success("🎉 照片時間分類中！")
    
    # 顯示結果
    for d, p_list in grouped_photos.items():
        st.write(f"📅 **{d}**：共 {len(p_list)} 張照片")
        
    # 產生Word 檔案
    with st.spinner("努力產生 Word 檔，敢吵我直接當機"):
        word_file = create_report(grouped_photos)
        
    # 下載按鈕
    st.download_button(
        label="📥 此為偷懶小幫手",
        data=word_file,
        file_name=f"監造報表_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
else:
    st.warning("請上傳照片再給我按，這種事還要我提醒。")
