import streamlit as st
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from PIL import Image, ImageDraw, ImageFont, ImageOps
from PIL.ExifTags import TAGS
import pillow_heif
import io
import os
import re
from datetime import datetime

# 註冊 HEIC 格式支援
pillow_heif.register_heif_opener()

# 初始化 
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
    
    # 若正規表達式 (Regex) 解析原始檔名中的 YYYYMMDD
    match = re.search(r"(\d{4})[-_ /]?(\d{2})[-_ /]?(\d{2})", file_name)
    if match:
        try:
            date_part = f"{match.group(1)}/{match.group(2)}/{match.group(3)}"
            dt = datetime.strptime(date_part, "%Y/%m/%d")
            return dt.strftime("%Y/%m/%d"), False
        except Exception:
            pass
            
    # 預設今天日期
    return datetime.today().strftime("%Y/%m/%d"), False

def resize_and_compress_image(image_bytes, date_str, is_from_exif):
    """
    處理照片尺寸與浮水印：
    - 自動修正手機拍攝時的 EXIF 旋轉問題。
    - 橫式照片：置中裁切填滿。
    - 直式照片：等比例縮放且限制高度不超出儲存格。
    - 調整尺寸後：若無原生時間，在右下角新增時間文字。
    """
    raw_img = Image.open(io.BytesIO(image_bytes))
    
    # 讀取手機 EXIF 的旋轉資訊並將照片轉正
    img = ImageOps.exif_transpose(raw_img)
    
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
        
    target_pixel_w = 945  # 8.0cm 對應pixel (300 DPI)
    target_ratio = 8.0 / 6.15
    target_pixel_h = int(target_pixel_w / target_ratio)  # 726 像素 (約 6.15cm，安全高度)
    
    img_w, img_h = img.size
    
    # 先進行裁切與基礎縮放 
    if img_w < img_h:
        # 直式照片縮小至固定格子內，但不要進行裁剪
        ratio_w = target_pixel_w / img_w
        ratio_h = target_pixel_h / img_h
        scale_ratio = min(ratio_w, ratio_h)
        new_w = int(img_w * scale_ratio)
        new_h = int(img_h * scale_ratio)
        img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
    else:
        # 橫式照片照舊：置中裁切並縮放
        current_ratio = img_w / img_h
        if current_ratio > target_ratio:
            crop_w = int(target_ratio * img_h)
            offset = (img_w - crop_w) // 2
            img = img.crop((offset, 0, img_w - offset, img_h))
        elif current_ratio < target_ratio:
            crop_h = int(img_w / target_ratio)
            offset = (img_h - crop_h) // 2
            img = img.crop((0, offset, img_w, img_h - offset))
        img = img.resize((target_pixel_w, target_pixel_h), Image.Resampling.LANCZOS)

    # 在標準化尺寸後，繪製固定大小的字體
    if not is_from_exif:
        draw = ImageDraw.Draw(img)
        w, h = img.size
        
        #  300 DPI 下，10 Pt 的字體大小固定為 42 pixels
        font_size = 42 
        
        try:
            font = ImageFont.truetype("msjh.ttc", font_size)  # 微軟正黑體
        except IOError:
            try:
                font = ImageFont.truetype("Arial.ttf", font_size)
            except IOError:
                font = ImageFont.load_default()

        # 計算文字寬高
        try:
            text_bbox = draw.textbbox((0, 0), date_str, font=font)
            text_w = text_bbox[2] - text_bbox[0]
            text_h = text_bbox[3] - text_bbox[1]
        except AttributeError:
            text_w, text_h = draw.textsize(date_str, font=font)
            
        # 右下角留邊
        x = w - text_w - 20
        y = h - text_h - 20
        
        # 繪製黑邊
        border_thickness = 3
        for dx in range(-border_thickness, border_thickness + 1):
            for dy in range(-border_thickness, border_thickness + 1):
                if dx != 0 or dy != 0:
                    draw.text((x + dx, y + dy), date_str, font=font, fill="black")
                    
        # 繪製主體白字
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
                    
                    if photo_pos < len(page_photos):
                        photo_data = page_photos[photo_pos]
                        
                        # 讀取圖片以確認橫直式比例，並精準插入 Word
                        img_obj = Image.open(io.BytesIO(photo_data["bytes"]))
                        img_w, img_h = img_obj.size
                        
                        run_img = p_para.add_run()
                        if img_w < img_h:
                            # 直式照片：縮小至固定格子內（限制高度不超標）
                            run_img.add_picture(io.BytesIO(photo_data["bytes"]), height=Cm(6.15))
                        else:
                            # 橫式照片：維持原程式縮放（限制寬度填滿）
                            run_img.add_picture(io.BytesIO(photo_data["bytes"]), width=Cm(8.0))
                            
                        # 填入說明文字
                        s_para = cell_small.paragraphs[0]
                        s_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        s_para.paragraph_format.space_before = Pt(0)
                        s_para.paragraph_format.space_after = Pt(0)
                        
                        run_desc = s_para.add_run(photo_data["description"])
                        run_desc.font.name = '標楷體'
                        run_desc.font.size = Pt(11)
                    else:
                        s_para = cell_small.paragraphs[0]
                        s_para.paragraph_format.space_after = Pt(0)
                        
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# -----------------------------------------
# Streamlit 前端網頁介面
# -----------------------------------------
st.title("📸 偷懶小幫手()")
st.caption("⚠️ 提醒：還是要使用時間相機app才會有浮水印喔！")

if "uploaded_photos_list" not in st.session_state:
    st.session_state.uploaded_photos_list = []

# 初始化控制清空 key
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

uploaded_files = st.file_uploader(
    "上傳照片", 
    type=["jpg", "jpeg", "png", "heic"], 
    accept_multiple_files=True,
    key=f"file_uploader_{st.session_state.uploader_key}"
)

if uploaded_files:
    existing_names = {p["name"] for p in st.session_state.uploaded_photos_list}
    for f in uploaded_files:
        if f.name not in existing_names:
            file_bytes = f.read()
            
            # 嘗試抓取拍攝日期
            date_str, is_from_exif = get_photo_date(file_bytes, f.name)
            
            # 如果不是從真實 EXIF 抓到的，或者抓到的是今天(錯誤)日期，直接判定為「未偵測到日期」
            if not is_from_exif:
                # 略過浮水印：傳入空字串或 None
                processed_io = resize_and_compress_image(file_bytes, date_str="", is_from_exif=False)
                # 網頁上的欄位改為空白，或讓使用者手動輸入
                default_display_date = "" 
            else:
                # 有抓到正確 EXIF 正常壓印
                processed_io = resize_and_compress_image(file_bytes, date_str, is_from_exif)
                default_display_date = date_str

            st.session_state.uploaded_photos_list.append({
                "name": f.name,
                "bytes": processed_io.getvalue(),
                "display_date": default_display_date,
                "description": ""
            })

# 圖片管理與編輯介面
if st.session_state.uploaded_photos_list:
    
    # 一鍵刪除功能鍵清除日期與文字
    col_btn1, col_btn2 = st.columns([1, 3])
    with col_btn1:
        if st.button("🗑️ 一鍵刪除", type="secondary", use_container_width=True):
            for idx in range(len(st.session_state.uploaded_photos_list)):
                date_key = f"date_input_{idx}"
                desc_key = f"desc_input_{idx}"
                if date_key in st.session_state:
                    del st.session_state[date_key]
                if desc_key in st.session_state:
                    del st.session_state[desc_key]
            
            # 2. 清除照片資料與重置上傳鍵
            st.session_state.uploaded_photos_list = []
            st.session_state.uploader_key += 1
            st.rerun()

    # 防呆:點擊別處或按 Enter 時才存檔
    def update_photo_data(index, field_key, session_key):
        st.session_state.uploaded_photos_list[index][field_key] = st.session_state[session_key]

    # 分組資料
    grouped_ui = {}
    for idx, photo in enumerate(st.session_state.uploaded_photos_list):
        grouped_ui.setdefault(photo["display_date"], []).append((idx, photo))
        
    # 排序: 讓沒有時間的排在最上面
    sorted_groups = sorted(grouped_ui.keys(), key=lambda x: (x != "", x))
    
    for d_str in sorted_groups:
        
        # -----------------------------------------------------------------
        # 沒有確定的時間時，直接呈現不變動
        # -----------------------------------------------------------------
        if d_str == "":
            st.markdown("請手動輸入時間（日期判定錯誤/未偵測到日期）")
            
            for original_idx, photo in grouped_ui[d_str]:
                col1, col2 = st.columns([1, 3])  
                with col1:
                    st.image(photo["bytes"], use_container_width=True)
                with col2:
                    date_key = f"date_input_{original_idx}"
                    st.text_input(
                        "補充拍攝日期_自己手動吧~", 
                        value=photo["display_date"], 
                        placeholder="輸入日期（例：YYYY/MM/DD）",
                        key=date_key,
                        on_change=update_photo_data,
                        args=(original_idx, "display_date", date_key)
                    )
                    
                    desc_key = f"desc_input_{original_idx}"
                    st.text_input(
                        "照片說明", 
                        value=photo["description"], 
                        key=desc_key,
                        on_change=update_photo_data,
                        args=(original_idx, "description", desc_key)
                    )
                st.write("---")
                
        # -----------------------------------------------------------------
        # 有確定的時間，使用 st.expander 可縮放功能
        # -----------------------------------------------------------------
        else:
            photo_count = len(grouped_ui[d_str])
            # 使用 expander 建立可折疊區塊，預設設為 False
            with st.expander(f"📅 日期：{d_str} (共 {photo_count} 張照片)", expanded=False):
                
                for original_idx, photo in grouped_ui[d_str]:
                    col1, col2 = st.columns([1, 3])
                    with col1:
                        st.image(photo["bytes"], use_container_width=True)
                    with col2:
                        date_key = f"date_input_{original_idx}"
                        st.text_input(
                            "修改日期_如有問題自己改喔~", 
                            value=photo["display_date"], 
                            key=date_key,
                            on_change=update_photo_data,
                            args=(original_idx, "display_date", date_key)
                        )
                        
                        desc_key = f"desc_input_{original_idx}"
                        st.text_input(
                            "照片說明", 
                            value=photo["description"], 
                            key=desc_key,
                            on_change=update_photo_data,
                            args=(original_idx, "description", desc_key)
                        )
                    st.write("---")

    # 封裝分組數據傳給報表生成函數
    final_grouped_photos = {}
    for photo in st.session_state.uploaded_photos_list:
        report_photo_node = {
            "name": photo["name"],
            "bytes": photo["bytes"], 
            "date": photo["display_date"], 
            "description": photo["description"]
        }
        final_grouped_photos.setdefault(photo["display_date"], []).append(report_photo_node)

    try:
        word_file = create_report(final_grouped_photos)
        
        # 獲取今天的日期及時間字串，避免 import 衝突
        import datetime
        current_date_str = datetime.datetime.now().strftime('%Y%m%d')
        
        st.download_button(
            label="🚀 按下去後就可以節省很多時間喔",
            data=word_file,
            file_name=f"監造報表_{current_date_str}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"產生 Word 報表失敗，給我檢查請處資料有沒有好嗎 :(：{e}")
