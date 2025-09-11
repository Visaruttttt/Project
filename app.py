import streamlit as st
import pandas as pd
import google.generativeai as genai
from PIL import Image
import os
from dotenv import load_dotenv
import io
import json
from PIL import Image 

# ติดตั้งไลบรารีนี้ก่อน: pip install PyMuPDF
import fitz 

# --- การตั้งค่าพื้นฐาน ---
load_dotenv() # โหลดค่าจากไฟล์ .env
st.set_page_config(page_title="ระบบสแกนใบเสร็จ", page_icon="🧾", layout="wide")

# ✅ ตรวจสอบว่าโหลด API KEY มาหรือไม่
api_key = os.getenv("GOOGLE_API_KEY")
if not api_key:
    st.error("❌ ไม่พบ GOOGLE_API_KEY ในไฟล์ .env กรุณาตรวจสอบอีกครั้ง")
    st.stop()
else:
    st.success("✅ โหลด GOOGLE_API_KEY สำเร็จ")
    st.write("API Key (บางส่วน):", api_key[:10] + "*****")

# --- การตั้งค่า Google Gemini API ---
try:
    genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
except Exception as e:
    st.error(f"เกิดข้อผิดพลาดในการตั้งค่า API: {e} กรุณาตรวจสอบว่าคุณได้ตั้งค่า GOOGLE_API_KEY ในไฟล์ .env ถูกต้อง")
    st.stop()
    
# ลดขนาดรูปภาพ โดยใช้ Libery Pillow
def compress_image(image, max_size_mb=4):
    """บีบอัดรูปภาพให้มีขนาดไม่เกินที่กำหนด"""
    img_byte_arr = io.BytesIO()
    image.save(img_byte_arr, format='JPEG', quality=85)
    img_byte_arr = img_byte_arr.getvalue()
    
    # ถ้าขนาดไฟล์ยังใหญ่เกินไป ให้ลดคุณภาพลงอีก
    while len(img_byte_arr) > max_size_mb * 1024 * 1024:
        img_byte_arr = io.BytesIO()
        image = image.resize((int(image.width * 0.9), int(image.height * 0.9)), Image.Resampling.LANCZOS)
        image.save(img_byte_arr, format='JPEG', quality=85)
        img_byte_arr = img_byte_arr.getvalue()
    return Image.open(io.BytesIO(img_byte_arr))

# --- ฟังก์ชันหลักในการประมวลผล ---

def get_gemini_response(input_text, image):
    """
    ฟังก์ชันสำหรับส่งคำสั่งและรูปภาพไปยัง Gemini Pro Vision
    และรับผลลัพธ์กลับมา
    """
    model = genai.GenerativeModel('gemini-1.5-flash')
    if image:
        response = model.generate_content([input_text, image])
        return response.text
    return "กรุณาอัปโหลดรูปภาพ"

def extract_data_from_image(image):
    """
    ดึงข้อมูลจากรูปภาพใบเสร็จโดยใช้ Gemini
    """
    prompt = """
    คุณคือผู้เชี่ยวชาญด้านการประมวลผลเอกสาร
    จากรูปภาพใบเสร็จ/ใบกำกับภาษีที่ให้มา โปรดดึงข้อมูลต่อไปนี้:
    1. ชื่อบริษัท (company_name)
    2. เลขประจำตัวผู้เสียภาษี (tax_id)
    3. เลขที่ใบกำกับภาษี หรือ เลขที่เอกสาร (invoice_number)
    4. วันที่ (date)
    5. รายการสินค้า/บริการทั้งหมด (items): โดยแต่ละรายการควรมี 'description' (คำอธิบาย), 'quantity' (จำนวน), 'unit_price' (ราคาต่อหน่วย), และ 'total_price' (ราคารวม)
    6. ยอดรวม (grand_total)

    โปรดตอบกลับเป็นรูปแบบ JSON ที่ถูกต้องเท่านั้น โดยไม่ต้องมีคำอธิบายหรือข้อความอื่นใดๆ นำหน้าหรือต่อท้าย
    หากข้อมูลส่วนใดไม่ปรากฏในภาพ ให้ใช้ค่าเป็น null
    ตัวอย่าง JSON:
    {
        "company_name": "ชื่อบริษัทตัวอย่าง",
        "tax_id": "1234567890123",
        "invoice_number": "INV-00123",
        "date": "2024-09-02",
        "items": [
            {"description": "สินค้า A", "quantity": 2, "unit_price": 50.0, "total_price": 100.0},
            {"description": "บริการ B", "quantity": 1, "unit_price": 250.0, "total_price": 250.0}
        ],
        "grand_total": 350.0
    }
    """
    try:
        response_text = get_gemini_response(prompt, image)
        # ทำความสะอาด response เพื่อให้เป็น JSON ที่ถูกต้อง
        cleaned_json = response_text.strip().replace("```json", "").replace("```", "")
        return json.loads(cleaned_json)
    except (json.JSONDecodeError, ValueError) as e:
        st.error(f"ไม่สามารถแปลงผลลัพธ์เป็น JSON ได้: {e}")
        st.code(response_text) # แสดงผลลัพธ์ที่ได้จาก AI เพื่อดีบัก
        return None
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการเรียก Gemini API: {e}")
        return None

def convert_to_excel(data_list):
    """
    แปลง List ของ Dictionary ที่ได้จากการสแกน
    ให้กลายเป็นไฟล์ Excel ในหน่วยความจำ (in-memory)
    """
    if not data_list:
        return None

    # สร้าง DataFrame หลักสำหรับข้อมูลทั่วไป
    main_data = []
    # สร้าง DataFrame สำหรับรายการสินค้า
    items_data = []

    for idx, data in enumerate(data_list):
        receipt_id = f"ใบเสร็จ_{idx+1}"
        main_info = {
            "ID ใบเสร็จ": receipt_id,
            "ชื่อบริษัท": data.get("company_name"),
            "เลขผู้เสียภาษี": data.get("tax_id"),
            "เลขใบกำกับภาษี": data.get("invoice_number"),
            "วันที่": data.get("date"),
            "ยอดรวม": data.get("grand_total")
        }
        main_data.append(main_info)

        if data.get("items"):
            for item in data["items"]:
                item_info = {
                    "ID ใบเสร็จ": receipt_id,
                    "รายการ": item.get("description"),
                    "จำนวน": item.get("quantity"),
                    "ราคาต่อหน่วย": item.get("unit_price"),
                    "ราคารวม": item.get("total_price")
                }
                items_data.append(item_info)

    df_main = pd.DataFrame(main_data)
    df_items = pd.DataFrame(items_data)

    # เขียนลง Excel buffer
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_main.to_excel(writer, sheet_name='สรุปใบเสร็จ', index=False)
        df_items.to_excel(writer, sheet_name='รายการสินค้า', index=False)

    processed_data = output.getvalue()
    return processed_data


# --- ส่วนของ Streamlit UI ---

st.title("🧾 ระบบสแกนใบเสร็จเพื่อลง Excel")
st.write("อัปโหลดไฟล์ใบเสร็จ (JPG, PNG, PDF) เพื่อดึงข้อมูลออกมาเป็นตารางและดาวน์โหลดเป็น Excel")

# จัดการ Session State เพื่อเก็บข้อมูลที่ประมวลผลแล้ว
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = []

uploaded_files = st.file_uploader(
    "เลือกไฟล์ใบเสร็จของคุณ",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    # เคลียร์ข้อมูลเก่าเมื่อมีการอัปโหลดใหม่
    if st.button("เริ่มต้นประมวลผลใหม่"):
        st.session_state.processed_data = []
        st.rerun()

    for uploaded_file in uploaded_files:
        st.markdown(f"---")
        st.subheader(f"ไฟล์: `{uploaded_file.name}`")
        
        col1, col2 = st.columns(2)

        images = []
        with col1:
            if uploaded_file.type == "application/pdf":
                st.write("กำลังแปลง PDF เป็นรูปภาพ...")
                # --- ใช้ PyMuPDF แทน pdf2image ---
                try:
                    doc = fitz.open(stream=uploaded_file.getvalue(), filetype="pdf")
                    # แปลงหน้าแรกของ PDF เป็นรูปภาพ
                    page = doc[0]
                    pix = page.get_pixmap()
                    images.append(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
                except Exception as e:
                    st.error(f"ไม่สามารถแปลงไฟล์ PDF ได้: {e}")
                    continue
                
                st.image(images[0], caption="หน้าแรกของ PDF", use_container_width=True)
            else:
                # ถ้าเป็นไฟล์รูปภาพ
                try:
                    images.append(Image.open(uploaded_file))
                    st.image(images[0], caption="รูปภาพที่อัปโหลด", use_container_width=True)
                except Exception as e:
                    st.error(f"ไม่สามารถแสดงไฟล์ได้: {e}")
                    continue

        with col2:
            if st.button(f"ดึงข้อมูลจาก `{uploaded_file.name}`", key=uploaded_file.name):
                with st.spinner("กำลังประมวลผลด้วย AI..."):
                    if images:
                        # บีบอัดรูปภาพก่อนส่งไปให้ AI
                        compressed_image = compress_image(images[0])
                        extracted_data = extract_data_from_image(compressed_image)
                        
                        if extracted_data:
                            st.session_state.processed_data.append(extracted_data)
                            st.success("ดึงข้อมูลสำเร็จ!")
                            st.json(extracted_data)
                        else:
                            st.error("ไม่สามารถดึงข้อมูลได้ กรุณาลองใหม่")
                    else:
                        st.warning("ไม่พบรูปภาพให้ประมวลผล")

# --- ส่วนของการดาวน์โหลด Excel ---
if st.session_state.processed_data:
    st.markdown("---")
    st.header("สรุปข้อมูลทั้งหมด")

    # แสดง DataFrame สรุป
    df_display = pd.DataFrame([{
        "บริษัท": d.get("company_name"),
        "เลขใบกำกับ": d.get("invoice_number"),
        "วันที่": d.get("date"),
        "ยอดรวม": d.get("grand_total")
    } for d in st.session_state.processed_data])
    st.dataframe(df_display)

    excel_data = convert_to_excel(st.session_state.processed_data)
    if excel_data:
        st.download_button(
            label="📥 ดาวน์โหลดข้อมูลเป็น Excel",
            data=excel_data,
            file_name="receipt_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )