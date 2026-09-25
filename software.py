import streamlit as st
import os
import io
import csv
import tempfile
import zipfile
import re
import difflib
import random
import concurrent.futures
import platform
import shutil
import streamlit.components.v1 as components
from pypdf import PdfWriter, PdfReader
from PIL import Image, ImageEnhance, ImageFilter, ImageStat, ImageChops
import docx
from docx.shared import Inches, Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

# --- KRUTIDEV TO UNICODE LOGIC ---
def convert_krutidev_to_unicode(text):
    if not text: return ""
    array_one = [
        "ñ","Q+Z","sas","aa",")Z","ZZ","‘","’","“","”",
        "å","ƒ","„","…","†","‡","ˆ","‰","Š","‹",
        "¶","d+","[k+","x+","T+","M+","<+","Q+",";+","j+","u+",
        "Ùk","Ù","ä","–","—","é","™","=kk","f=k",
        "à","á","â","ã","ºz","º","í","{k","{","=","«",
        "Nî","Vî","Bî","Mî","<î","|","K","}",
        "J","Vª","Mª","<ªª","Nª","Ø","Ý","nzZ","æ","ç","Á","xz","#",":",
        "v‚","vks","vkS","vk","v","b±","Ã","bZ","b","m","Å",",s",",","_",
        "d","D","k","K","[k","[","x","X","?","Œ","p","P","N","t","T",">","÷","ß","M","<",".k",".",
        "r","R","Fk","F","n","èk","è","u","U","i","I","Q","¶","c","C","Hk","H","e","E",";","¸","j","y","Y","y","Y","o","O",
        "'k","'","\"k","\"","l","L","g",
        "È","z",
        "Ì","Í","Î","Ï","Ñ","Ò","Ó","Ô","Ö","Ø","Ù","Ü","Ý","Þ","ß","à","á","â","ã","ä","å","æ","ç","è","é","ê","ë","ì","í","î","ï","ð","ñ","ò","ó","ô","õ","ö","÷","ø","ù","ú","û","ü","ý","þ","ÿ",
        "A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z",
        "a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z",
        "0","1","2","3","4","5","6","7","8","9"
    ]
    array_two = [
        "॰","QZ+","sa","a","Z)","Z","\"","\"","'","'",
        "०","१","२","३","४","५","६","७","८","९",
        "४","क़","ख़","ग़","ज़","ड़","ढ़","फ़","य़","ऱ","ऩ",
        "त्त","त्त्","क्त","दृ","कृ","न्न","न्न्","=k","f=",
        "ह्न","ह्य","हृ","ह्म","क्र","र्को","द्द","क्ष","क्ष्","त्र","त्र्",
        "छ्य","ट्य","ठ्य","ड्य","ढ्य","।","ज्ञ","द्व",
        "श्र","ट्र","ड्र","ढ्र","छ्र","क्र","फ्र","र्द्र","द्र","प्र","प्र","ग्र","रु","रू",
        "ऑ","ओ","औ","आ","अ","ईं","ई","ई","इ","उ","ऊ","ऐ","ए","ऋ",
        "क","क्","क","क्","ख","ख्","ग","ग्","घ","घ्","च","च्","छ","ज","ज्","झ","झ्","ञ","ट","ठ","ड","ढ","ण","ण्",
        "त","त्","थ","थ्","द","ध","ध्","न","न्","प","प्","फ","फ्","ब","ब्","भ","भ्","म","म्","य","य्","र","ल","ल्","ळ","ळ्","व","व्",
        "श","श्","ष","ष्","स","स्","ह",
        "ीं","्र",
        "द्द","ट्ट","ट्ठ","ड्ड","कृ","भ","्य","ड्ढ","झ्","क्र","त्त्","श","श्","व","व","र","र","ह","ह","क्क","क","क","ल","ल","भ","भ","श","श","ष","ष","स","स","स","स","त्र","त्र","छ","ट","ठ","ड","ढ","ण","ण","त","त","थ",
        "ं","इ","च","क्","म्","ँ","ळ","ः","्","श्र","क्","स्","म्","छ","ब्","प्","ृ","र्","ै","त्","न्","व्","ॅ","ग्","ल्","्",
        "ा","ि","च्","क","म","ि","ह","ी","ं","र","ा","स","म","द","ब","प","ु","र","े","त","न","व","ू","ग","ल","र्",
        "०","१","२","३","४","५","६","७","८","९"
    ]
    for a1, a2 in zip(array_one, array_two):
        text = text.replace(a1, a2)
    text = re.sub(r'ि([क-ह])(्[क-ह])*', r'\1\2ि', text)
    text = re.sub(r'([क-ह])([ािीुूृेैोौंँ]*)र्', r'र्\1\2', text)
    return text

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="दस्तावेज़ सेतु",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- PREMIUM STYLING & UI ENHANCEMENTS ---
st.markdown("""
    <style>
    /* Background Color */
    .stApp { background-color: #F4F6F7; }
    
    /* Enhanced Main Header */
    .main-header {
        background: linear-gradient(135deg, #6b0f0f 0%, #3d0707 100%);
        padding: 40px 20px;
        border-radius: 16px;
        text-align: center;
        color: white;
        margin-bottom: 35px;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.2);
        border-bottom: 4px solid #F39C12;
    }
    .org-title { font-size: 18px; letter-spacing: 3px; font-weight: 700; color: #E0E0E0; text-transform: uppercase; margin-bottom: 8px; }
    .main-title { color: #F39C12; font-size: 58px; font-weight: 900; margin: 0; text-shadow: 3px 3px 6px rgba(0,0,0,0.4); letter-spacing: 2px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .portal-subtitle { font-size: 16px; color: #FADBD8; margin-top: 10px; font-weight: 500; font-style: italic; }
    
    /* Section Headings */
    h2 { color: #7A1212 !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; border-bottom: 3px solid #F39C12; padding-bottom: 10px; margin-top: 40px; font-weight: 800; text-transform: uppercase; letter-spacing: 1px; }
    
    /* Premium Expanders */
    [data-testid="stExpander"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E0E0E0 !important;
        border-radius: 10px !important;
        margin-bottom: 12px !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05) !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease !important;
    }
    [data-testid="stExpander"]:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 15px rgba(0,0,0,0.12) !important;
        border-color: #7A1212 !important;
    }
    [data-testid="stExpander"] summary,
    [data-testid="stExpander"] summary p,
    [data-testid="stExpander"] summary span,
    .streamlit-expanderHeader,
    .streamlit-expanderHeader p {
        color: #1A252F !important;
        font-weight: 700 !important;
        font-size: 16px !important;
    }
    [data-testid="stExpander"] summary svg {
        fill: #7A1212 !important;
        color: #7A1212 !important;
    }

    /* Fixed Footer */
    .footer { position: fixed; left: 0; bottom: 0; width: 100%; background: linear-gradient(135deg, #1B2631 0%, #2C3E50 100%); color: #FADBD8; text-align: center; padding: 15px; font-size: 14px; font-weight: 500; letter-spacing: 0.5px; z-index: 100; box-shadow: 0 -4px 15px rgba(0,0,0,0.2); }
    </style>
""", unsafe_allow_html=True)

# --- SIDEBAR ---
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/c/cb/New_Parliament_Building%2C_New_Delhi_-_Mar_2024.jpg", width="stretch")
st.sidebar.markdown("### ⚙️ System Actions")
if st.sidebar.button("🧹 Clear Workspace & Cache", width="stretch"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

st.sidebar.markdown("<br><br><br><br>", unsafe_allow_html=True)
st.sidebar.caption("v3.13 Cloud Build (Bugfix)")

# --- HEADER & LIVE SEARCH ---
st.markdown("""
    <div class="main-header">
        <div class="org-title">Rajya Sabha Secretariat</div>
        <div class="main-title">दस्तावेज़ सेतु</div>
        <div class="portal-subtitle">Secure Intranet Document Utility & Conversion Portal</div>
    </div>
""", unsafe_allow_html=True)

st.markdown("### 🔍 Live Tool Search")
search_query = st.text_input("Search", "", placeholder="Start typing to instantly find a tool (e.g., 'CGHS', 'Gate Pass', 'Directory', 'Word')...", label_visibility="collapsed").strip()

def is_match(title, keywords=""):
    if not search_query: return True
    search_words = search_query.lower().split()
    target_text = (title + " " + keywords).lower()
    for word in search_words:
        if word not in target_text:
            return False
    return True

# Pre-calculate visibility for all 55 tools
t1 = is_match("1. Merge Multiple PDFs (Advanced Reorder)", "combine join append together")
t2 = is_match("2. Extract Specific Pages from PDF", "keep pull specific part")
t3 = is_match("3. Split PDF into Individual Pages", "divide break separate")
t4 = is_match("4. Delete Specific Pages from PDF", "remove cut exclude trash")
t5 = is_match("5. Reverse PDF Page Order (Back to Front)", "flip backwards order")
t6 = is_match("6. Generate Official Notice Draft Template", "word docx memo letter typing")
t7 = is_match("7. Generate Visitor Pass (Auto-Fill & Multiple Visitors)", "entry security guest aadhaar")
t8 = is_match("8. Generate Gate Pass (Systems Division Auto-Fill)", "exit material hardware equipment")
t9 = is_match("9. Generate ITDC Catering Note (Meeting Arrangements)", "tea biscuits food water")
t10 = is_match("10. Generate Green Note Sheet (फाइल नोटिंग) Template", "noting file margins approval")
t11 = is_match("11. Convert Single Image to PDF", "jpg png picture format")
t12 = is_match("12. Combine Multiple Images to One PDF", "join pictures photos format")
t13 = is_match("13. Convert Word Document to PDF", "docx format")
t14 = is_match("14. Convert PDF to Word Document", "docx format editable")
t15 = is_match("15. Scan PDF to Notepad (Full OCR - Scanned Docs)", "text extract picture readable")
t16 = is_match("16. Fast Text Extract (Digital PDFs Only)", "notepad txt pull words")
t17 = is_match("17. Extract Text from Word to Notepad", "pull txt words")
t18 = is_match("18. Extract All Embedded Images from PDF", "pull pictures photos out extract")
t19 = is_match("19. Extract PDF Tables to Excel (.csv)", "spreadsheet data row column")
t20 = is_match("20. Password Protect / Encrypt a PDF", "lock secure hide password")
t21 = is_match("21. Unlock / Remove PDF Password", "decrypt open remove password")
t22 = is_match("22. Apply Custom Watermark / Stamp to PDF", "confidential text background draft")
t23 = is_match("23. Add Page Numbers (Bates Numbering)", "pagination footer numbers")
t24 = is_match("24. Insert Blank Pages (For Duplex Printing)", "white empty add page")
t25 = is_match("25. Rotate PDF Pages (90°)", "turn flip landscape portrait")
t26 = is_match("26. Compress / Optimize PDF File Size", "reduce smaller shrink mb kb")
t27 = is_match("27. View Hidden PDF Metadata", "author date properties info")
t28 = is_match("28. Remove Image Background & Change Color (AI Offline)", "transparent color red green blue solid clear")
t29 = is_match("29. Resize & Compress Image", "shrink dimensions pixels smaller")
t30 = is_match("30. KrutiDev 010 to Unicode (Mangal) Converter", "hindi font translation legacy type")
t31 = is_match("31. True PDF Redaction (ऑटोमैटिक ब्लैकआउट)", "hide blackout remove text hide personal")
t32 = is_match("32. Bulk Generator (Excel/CSV to Word Mail Merge)", "mass generate template multiple")
t33 = is_match("33. Add Table of Contents (Clickable Bookmarks) to PDF", "index sidebar links click")
t34 = is_match("34. PDF Grayscale / B&W Converter (Printer Toner Saver)", "black white monochrome ink saver")
t35 = is_match("35. Visual PDF Comparison (Draft Diff Tool)", "compare changes track different")
t36 = is_match("36. Searchable 'Sandwich' PDF (Invisible OCR Overlay)", "ctrl-f find text scan ocr")
t37 = is_match("37. Digital Facsimile Signature & Stamp Placer", "sign approve stamp electronic")
t38 = is_match("38. N-Up & Booklet Layout Maker (Multiple Pages on 1 Sheet)", "paper saving multiple layout")
t39 = is_match("39. Institutional Bates Stamping (Custom Prefix Legal Numbering)", "legal prefix serial number stamp")
t40 = is_match("40. PDF Auto-Crop & White Margin Trimmer", "remove borders edge crop whitespace")
t41 = is_match("41. Parliamentary & Admin Glossary (सचिवालयीन शब्दकोश)", "dictionary hindi translation meaning")
t42 = is_match("42. Make PDF Look Scanned (डिजिटल को 'स्कैन-लुक' दें)", "fake noisy tilt old scan effect")
t43 = is_match("43. Quick Dak/Diary Receipt Stamper (ई-डायरी मुहर)", "received date stamp dak letter")
t44 = is_match("44. PDF A4 Page Standardizer (प्रिंटर पेपर-साइज फिक्सर)", "resize a4 format uniform paper")
t45 = is_match("45. Scanned Book / Landscape Page Slicer (हाफ-पेज स्प्लिटर)", "cut half dual book landscape")
t46 = is_match("46. Offline Telephone & Room Directory (सचिवालय टेलीफोन निर्देशिका)", "phone book contact number room directory")
t47 = is_match("47. CGHS Medical Reimbursement Claim Form - MRC(S) (चिकित्सा प्रतिपूर्ति दावा प्रपत्र)", "cghs medical reimbursement claim form mrc health hospital bill")
t48 = is_match("48. Scanned PDF Whitener & De-Shadow Cleaner (दस्तावेज़ बैकग्राउंड क्लीनर)", "clean sharp contrast grey white")
t49 = is_match("49. Smart Auto-PII Redactor (आधार / पैन / मोबाइल ऑटो-ब्लैकआउट)", "hide aadhaar pan email phone personal")
t50 = is_match("50. Manual Duplex Printing Assistant (ऑड-ईवन प्रिंट हेल्पर)", "print double sided even odd printer")
t51 = is_match("51. PDF Color vs B&W Page Audit & Splitter (प्रिंटर बजट सेवर)", "split audit color black white printer ink")
t52 = is_match("52. PDF Embedded Portfolio / Attachment Packer (डिजिटल मिसल टूल)", "attach embed missil pack files portfolio")
t53 = is_match("53. Smart Booklet Imposition (Center-Staple Maker)", "booklet fold staple print saddle layout")
t54 = is_match("54. Application for Medical Test/Treatment Permission", "medical diagnostic test treatment permission cghs hospital health form")
t55 = is_match("55. Convert PDF to eBook (EPUB)", "epub kindle ebook format read book")

if not any([t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13, t14, t15, t16, t17, t18, t19, t20, t21, t22, t23, t24, t25, t26, t27, t28, t29, t30, t31, t32, t33, t34, t35, t36, t37, t38, t39, t40, t41, t42, t43, t44, t45, t46, t47, t48, t49, t50, t51, t52, t53, t54, t55]):
    st.warning("⚠️ No utilities match your search. Try adjusting your keywords.")

# ==========================================
# CATEGORY 1: CORE FILE MANAGEMENT
# ==========================================
if any([t1, t2, t3, t4, t5]):
    st.markdown("## 📁 Core File Management")
    if t1:
        with st.expander("1. Merge Multiple PDFs (Advanced Reorder)"):
            uploaded_files = st.file_uploader("Upload PDF files to merge", type=["pdf"], accept_multiple_files=True, key="merge_pdf")
            if uploaded_files and st.button("Merge PDFs", key="btn_merge"):
                merger = PdfWriter()
                for f in uploaded_files: merger.append(f)
                out = io.BytesIO()
                merger.write(out)
                merger.close()
                st.success("PDFs merged successfully!")
                st.download_button("Download Merged PDF", data=out.getvalue(), file_name="merged_output.pdf", mime="application/pdf")
    if t2:
        with st.expander("2. Extract Specific Pages from PDF"):
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="ext_pdf")
            if uploaded_file:
                reader = PdfReader(uploaded_file)
                total_p = len(reader.pages)
                st.info(f"Total pages in PDF: {total_p}")
                c1, c2 = st.columns(2)
                start = c1.number_input("Start Page", min_value=1, max_value=total_p, value=1)
                end = c2.number_input("End Page", min_value=start, max_value=total_p, value=total_p)
                if st.button("Extract Pages", key="btn_ext"):
                    writer = PdfWriter()
                    for i in range(int(start) - 1, int(end)): writer.add_page(reader.pages[i])
                    out = io.BytesIO()
                    writer.write(out)
                    writer.close()
                    st.success("Pages extracted successfully!")
                    st.download_button("Download Extracted PDF", data=out.getvalue(), file_name="extracted_pages.pdf", mime="application/pdf")
    if t3:
        with st.expander("3. Split PDF into Individual Pages"):
            uploaded_file = st.file_uploader("Upload PDF to Split", type=["pdf"], key="split_pdf")
            if uploaded_file and st.button("Split Pages", key="btn_split"):
                reader = PdfReader(uploaded_file)
                st.success(f"PDF has {len(reader.pages)} pages. Use Extract Pages above to download ranges.")
    if t4:
        with st.expander("4. Delete Specific Pages from PDF"):
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="del_pdf")
            pages_to_del = st.text_input("Enter page numbers to delete (comma-separated, e.g., 1, 3)", key="del_pages_input")
            if uploaded_file and pages_to_del and st.button("Delete Pages", key="btn_del"):
                try:
                    del_list = [int(p.strip()) for p in pages_to_del.split(",")]
                    reader = PdfReader(uploaded_file)
                    writer = PdfWriter()
                    for i, page in enumerate(reader.pages):
                        if (i + 1) not in del_list: writer.add_page(page)
                    out = io.BytesIO()
                    writer.write(out)
                    out.seek(0)
                    st.success("Selected pages deleted successfully!")
                    st.download_button("Download Trimmed PDF", data=out, file_name="trimmed_output.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t5:
        with st.expander("5. Reverse PDF Page Order (Back to Front)"):
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="rev_pdf")
            if uploaded_file and st.button("Reverse Order", key="btn_rev"):
                reader = PdfReader(uploaded_file)
                writer = PdfWriter()
                for page in reversed(reader.pages): writer.add_page(page)
                out = io.BytesIO()
                writer.write(out)
                out.close()
                st.success("PDF pages reversed successfully!")
                st.download_button("Download Reversed PDF", data=out.getvalue(), file_name="reversed_output.pdf", mime="application/pdf")

# ==========================================
# CATEGORY 2: CONVERSIONS & DRAFTING
# ==========================================
if any([t6, t7, t8, t9, t10, t11, t12, t13, t14, t55]):
    st.markdown("## 📝 Conversions & Drafting")
    if t6:
        with st.expander("6. Generate Official Notice Draft Template"):
            if st.button("Generate Template (.docx)", key="btn_template"):
                doc = docx.Document()
                for s in doc.sections: s.top_margin = Inches(1); s.bottom_margin = Inches(1); s.left_margin = Inches(1); s.right_margin = Inches(1)
                p_header = doc.add_paragraph()
                p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p_header.add_run("RAJYA SABHA SECRETARIAT\nBRANCH NAME")
                r.bold = True; r.font.size = Pt(14)
                doc.add_paragraph()
                p_sub = doc.add_paragraph()
                p_sub.add_run("Subject: ").bold = True
                p_sub.add_run("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX")
                doc.add_paragraph("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX.")
                doc.add_paragraph("2.\t XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX.")
                doc.add_paragraph("3.\tXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX.")
                doc.add_paragraph("4.\tXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX.")
                doc.add_paragraph()
                p_sig = doc.add_paragraph()
                p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_sig.add_run("(NAME)\nDesignation\nXX.XX.XXXX\nTel: XXXXXXX")
                doc.add_paragraph()
                doc.add_paragraph("To,\nXXXXXXXXXXXXXXXXXXXXXXX\nXXXXXXXXXXXXXXXXXXXXXXX\nXXXXXXXXXXXXXXXXXXXXXXX")
                out = io.BytesIO()
                doc.save(out); out.seek(0)
                st.success("Notice Template generated successfully!")
                st.download_button("Download Draft Template", data=out, file_name="Notice_Template.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    if t7:
        with st.expander("7. Generate Visitor Pass (Auto-Fill & Multiple Visitors)"):
            c1, c2 = st.columns(2)
            vp_officer = c1.selectbox("NAME OF THE OFFICER", ["Navneet Joon", "Sandeep Pandey", "Devyanshu Pal", "Ankit Chansoria", "Agam Mittal", "R.Arivazhagan"], key="vp_off")
            vp_desig = c2.selectbox("DESIGNATION", ["Deputy Secretary", "Under Secretary", "Executive Officer"], key="vp_desig")
            c3, c4 = st.columns(2)
            vp_room = c3.text_input("ROOM NO.", value="209, 2nd Floor, Parliament House Annexe", key="vp_room")
            vp_tel = c4.text_input("OFFICE TELEPHONE NO.", value="011-23034325", key="vp_tel")
            if 'visitor_count' not in st.session_state: st.session_state['visitor_count'] = 1
            v_col1, v_col2 = st.columns(2)
            if v_col1.button("➕ Add Visitor", key="add_v"): st.session_state['visitor_count'] += 1
            if v_col2.button("➖ Remove Visitor", key="rem_v") and st.session_state['visitor_count'] > 1: st.session_state['visitor_count'] -= 1
            visitor_names, visitor_aadhaars = [], []
            for i in range(st.session_state['visitor_count']):
                vc1, vc2 = st.columns(2)
                name = vc1.text_input(f"Visitor {i+1} Name", key=f"v_name_{i}")
                aadhaar = vc2.text_input(f"Visitor {i+1} ID Details", key=f"v_aadh_{i}")
                visitor_names.append(name.strip() if name else "")
                visitor_aadhaars.append(aadhaar.strip() if aadhaar else "")
            c7, c8 = st.columns(2)
            vp_datetime = c7.text_input("DATE AND TIME", key="vp_dt")
            vp_purpose = c8.text_input("PURPOSE OF VISIT", value="Official Meeting", key="vp_purp")
            vp_gadgets = st.selectbox("Allow Mobile & Laptop?", ["✅ Yes (Allow gadgets and print line)", "❌ No (Do not include gadget permission line)"], key="vp_gadgets")
            if st.button("Generate Visitor Pass (.docx)", key="btn_vp"):
                doc = docx.Document()
                for s in doc.sections: s.top_margin = Inches(1); s.bottom_margin = Inches(1); s.left_margin = Inches(1); s.right_margin = Inches(1)
                p_head = doc.add_paragraph()
                p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_head = p_head.add_run("PARLIAMENT OF INDIA\nRAJYA SABHA SECRETARIAT\n[SYSTEMS DIVISION]")
                r_head.bold = True; r_head.font.size = Pt(13)
                doc.add_paragraph()
                table = doc.add_table(rows=0, cols=3)
                table.columns[0].width = Inches(2.2); table.columns[1].width = Inches(0.2); table.columns[2].width = Inches(3.8)
                combined_visitors = []
                for n, a in zip(visitor_names, visitor_aadhaars):
                    if n or a:
                        line = n if n else "Unknown Visitor"
                        if a: line += f" (ID: {a})"
                        combined_visitors.append(line)
                vp_visitor_str = "\n".join(combined_visitors) if combined_visitors else ""
                fields = [("NAME OF THE OFFICER", vp_officer), ("DESIGNATION", vp_desig), ("ROOM NO.", vp_room), ("OFFICE TELEPHONE NO.", vp_tel), ("NAME OF THE VISITOR(S)", vp_visitor_str)]
                for label, val in fields:
                    row = table.add_row().cells
                    p_label = row[0].paragraphs[0]
                    p_label.add_run(label).bold = True
                    row[1].text = ":"
                    row[2].text = val
                doc.add_paragraph()
                if "Yes" in vp_gadgets:
                    p_note = doc.add_paragraph()
                    p_note.add_run("(Above person may be allowed along with his Mobile phone and Laptop with adapters/chargers)").bold = True
                doc.add_paragraph()
                p_dt = doc.add_paragraph()
                p_dt.add_run("DATE AND TIME:\t\t").bold = True; p_dt.add_run(vp_datetime)
                p_purp = doc.add_paragraph()
                p_purp.add_run("PURPOSE OF VISIT:\t").bold = True; p_purp.add_run(vp_purpose)
                doc.add_paragraph("\n\n")
                p_sig = doc.add_paragraph()
                p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_sig.add_run("(SIGNATURE WITH STAMP)").bold = True
                out = io.BytesIO()
                doc.save(out); out.seek(0)
                st.success("Visitor Pass generated successfully!")
                st.download_button("Download Visitor Pass", data=out, file_name="Visitor_Pass.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    if t8:
        with st.expander("8. Generate Gate Pass (Systems Division Auto-Fill)"):
            c1, c2 = st.columns(2)
            gp_officer = c1.text_input("Name of officer/Section", key="gp_off")
            gp_date = c2.text_input("Date of Issue", key="gp_date")
            gp_item = st.text_area("Item Description", key="gp_item")
            gp_make = st.text_input("Make", key="gp_make")
            c3, c4 = st.columns(2)
            gp_favour = c3.text_input("Issue in favour of Shri", key="gp_fav")
            gp_org = c4.text_input("Of (organization)", key="gp_org")
            c5, c6 = st.columns(2)
            gp_from = c5.text_input("Valid for movement from", key="gp_from")
            gp_to = c6.text_input("To", key="gp_to")
            c7, c8 = st.columns(2)
            gp_gate = c7.text_input("Through Gate No.", key="gp_gate")
            gp_ondate = c8.text_input("On Date", key="gp_ondate")
            c9, c10 = st.columns(2)
            gp_between = c9.text_input("Between (Time)", key="gp_bet")
            gp_and = c10.text_input("To (Time)", key="gp_and")
            c11, c12 = st.columns(2)
            gp_sig_name = c11.text_input("Issuer Name (for Signature)", key="gp_sig_name")
            gp_sig_desig = c12.text_input("Issuer Designation", key="gp_sig_desig")
            if st.button("Generate Gate Pass (.docx)", key="btn_gp"):
                doc = docx.Document()
                for s in doc.sections: s.top_margin = Inches(1); s.bottom_margin = Inches(1); s.left_margin = Inches(1); s.right_margin = Inches(1)
                p_head = doc.add_paragraph()
                p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r1 = p_head.add_run("RAJYA SABHA SECRETARIAT\n"); r1.underline = True; r1.font.size = Pt(14)
                r2 = p_head.add_run("SYSTEMS DIVISION"); r2.font.size = Pt(13)
                p_title = doc.add_paragraph()
                p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                rt = p_title.add_run("\nGATE PASS\n"); rt.underline = True; rt.font.size = Pt(13)
                table = doc.add_table(rows=0, cols=2)
                table.autofit = False
                def add_gp_row(lbl, val):
                    row = table.add_row()
                    row.cells[0].width = Inches(2.3); row.cells[1].width = Inches(4.2)
                    row.cells[0].paragraphs[0].add_run(lbl)
                    r = row.cells[1].paragraphs[0].add_run(f" {val} " if val else " " * 40)
                    r.underline = True
                add_gp_row("Name of officer/Section:", gp_officer)
                add_gp_row("Date of Issue:", gp_date)
                add_gp_row("Item Description:", gp_item)
                add_gp_row("Make:", gp_make)
                add_gp_row("Issue in favour of Shri:", gp_favour)
                add_gp_row("Of (organization):", gp_org)
                add_gp_row("Valid for movement from:", gp_from)
                add_gp_row("To:", gp_to)
                add_gp_row("Through Gate No.:", gp_gate)
                add_gp_row("On Date:", gp_ondate)
                add_gp_row("Between (Time):", f"{gp_between} to {gp_and}")
                doc.add_paragraph("\n")
                p_sig = doc.add_paragraph()
                p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_sig.add_run(f"{gp_sig_name or '(Name)'}\n").bold = True
                p_sig.add_run(f"{gp_sig_desig or 'DESIGNATION'}")
                out = io.BytesIO()
                doc.save(out); out.seek(0)
                st.success("Gate Pass generated successfully!")
                st.download_button("Download Gate Pass", data=out, file_name="Gate_Pass.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    if t9:
        with st.expander("9. Generate ITDC Catering Note (Meeting Arrangements)"):
            c1, c2 = st.columns(2)
            mtg_name = c1.text_input("Meeting / Committee Name", value="FBEC", key="itdc_mtg")
            mtg_date = c2.text_input("Meeting Date", value="24.08.2026", key="itdc_mdate")
            c3, c4 = st.columns(2)
            mtg_time = c3.text_input("Meeting Time", value="4:00 PM", key="itdc_time")
            mtg_venue = c4.text_input("Venue / Room No.", value="Room No: 124, First Floor, PHA", key="itdc_venue")
            c5, c6 = st.columns(2)
            tea_cups = c5.text_input("Number of Tea Cups", value="10", key="itdc_tea")
            biscuits = c6.text_input("Number of Biscuit Packets", value="5", key="itdc_bisc")
            c7, c8 = st.columns(2)
            itdc_off = c7.selectbox("NAME OF THE OFFICER", ["Navneet Joon", "Sandeep Pandey", "Devyanshu Pal", "Ankit Chansoria", "Agam Mittal", "R.Arivazhagan"], key="itdc_off_sel")
            itdc_desig = c8.selectbox("DESIGNATION", ["Deputy Secretary", "Under Secretary", "Executive Officer"], key="itdc_desig_sel")
            c9, c10 = st.columns(2)
            itdc_issue_date = c9.text_input("Issue Date", value="21.08.2026", key="itdc_idate")
            itdc_tel = c10.text_input("Tel No.", value="23034068", key="itdc_tel")
            if st.button("Generate ITDC Note (.docx)", key="btn_itdc"):
                doc = docx.Document()
                for s in doc.sections: s.top_margin = Inches(1); s.bottom_margin = Inches(1); s.left_margin = Inches(1.5); s.right_margin = Inches(1)
                p_head = doc.add_paragraph()
                p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_head = p_head.add_run("RAJYA SABHA SECRETARIAT\nSYSTEMS DIVISION")
                r_head.bold = True; r_head.font.size = Pt(13)
                doc.add_paragraph()
                p_sub = doc.add_paragraph()
                p_sub.add_run(f"Subject: Arrangement of Water Bottles, Glasses, Tea, Biscuits for {mtg_name} Meeting.").bold = True
                p_body = doc.add_paragraph()
                p_body.add_run(f"India Tourism Development Corporation (ITDC) Ltd., Parliament House Annexe (PHA), is requested to arrange water bottles, glasses, tea ({tea_cups} cups), biscuits ({biscuits} packets) for the Meeting of {mtg_name} scheduled as follows:\n\nVenue:\t{mtg_venue}\nDate:\t{mtg_date}\nTime:\t{mtg_time}")
                doc.add_paragraph("\n2. The above arrangements may kindly be ensured accordingly.\n")
                p_sig = doc.add_paragraph()
                p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_sig.add_run(f"({itdc_off})\n").bold = True
                p_sig.add_run(f"{itdc_desig}\n{itdc_issue_date}\nTel: {itdc_tel}")
                doc.add_paragraph("\nTo,\nThe Manager,\nIndia Tourism Development Corporation,\nParliament House Annexe,\nNew Delhi.")
                out = io.BytesIO()
                doc.save(out); out.seek(0)
                st.success("ITDC Catering Note generated successfully!")
                st.download_button("Download ITDC Note", data=out, file_name="ITDC_Catering_Note.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    if t10:
        with st.expander("10. Generate Green Note Sheet (फाइल नोटिंग) Template"):
            c1, c2 = st.columns(2)
            gn_file_no = c1.text_input("File No.", value="RS/Sys/2026/...", key="gn_file")
            gn_date = c2.text_input("Date", value="21.08.2026", key="gn_date")
            gn_subject = st.text_input("Subject", value="Regarding...", key="gn_sub")
            gn_body = st.text_area("Note Description (टिप्पणी)", value="Put up for approval please.", height=150, key="gn_body")
            c3, c4 = st.columns(2)
            gn_off = c3.selectbox("NAME OF THE OFFICER", ["Navneet Joon", "Sandeep Pandey", "Devyanshu Pal", "Ankit Chansoria", "Agam Mittal", "R.Arivazhagan"], key="gn_off_sheet")
            gn_desig = c4.selectbox("DESIGNATION", ["Deputy Secretary", "Under Secretary", "Executive Officer"], key="gn_desig_sheet")
            if st.button("Generate Note Sheet (.docx)", key="btn_green_note"):
                doc = docx.Document()
                for s in doc.sections: s.top_margin = Inches(1); s.bottom_margin = Inches(1); s.left_margin = Inches(1.5); s.right_margin = Inches(1)
                p_head = doc.add_paragraph()
                p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_head.add_run("RAJYA SABHA SECRETARIAT\nSYSTEMS DIVISION\n").bold = True
                p_head.add_run("(NOTE SHEET)").bold = True
                doc.add_paragraph()
                doc.add_paragraph(f"F.No.: {gn_file_no}").runs[0].bold = True
                p_sub = doc.add_paragraph()
                p_sub.add_run("Subject: ").bold = True; p_sub.add_run(gn_subject).bold = True
                doc.add_paragraph()
                p_b = doc.add_paragraph(gn_body)
                p_b.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                doc.add_paragraph("\n")
                p_sig = doc.add_paragraph()
                p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_sig.add_run(f"({gn_off})\n").bold = True
                p_sig.add_run(f"{gn_desig}\n{gn_date}")
                out = io.BytesIO()
                doc.save(out); out.seek(0)
                st.success("Green Note Sheet generated successfully!")
                st.download_button("Download Note Sheet", data=out, file_name="Green_Note_Sheet.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    if t11:
        with st.expander("11. Convert Single Image to PDF"):
            img_file = st.file_uploader("Upload Image (JPG/PNG)", type=["jpg", "jpeg", "png"], key="img_pdf")
            if img_file and st.button("Convert to PDF", key="btn_img_pdf"):
                image = Image.open(img_file).convert('RGB')
                out = io.BytesIO()
                image.save(out, format='PDF')
                st.success("Image converted successfully!")
                st.download_button("Download PDF", data=out.getvalue(), file_name="converted_image.pdf", mime="application/pdf")
    if t12:
        with st.expander("12. Combine Multiple Images to One PDF"):
            uploaded_files = st.file_uploader("Upload Multiple Images", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="m_imgs")
            if uploaded_files and st.button("Combine Images", key="btn_m_imgs"):
                image_list = [Image.open(f).convert('RGB') for f in uploaded_files]
                if image_list:
                    out = io.BytesIO()
                    image_list[0].save(out, format='PDF', save_all=True, append_images=image_list[1:])
                    st.success("Images combined into PDF successfully!")
                    st.download_button("Download Combined PDF", data=out.getvalue(), file_name="combined_images.pdf", mime="application/pdf")
    if t13:
        with st.expander("13. Convert Word Document to PDF"):
            if platform.system() == "Linux":
                st.info("ℹ️ Note: Microsoft Word to PDF conversion relies on MS Office software and is supported on local Windows/Mac environments.")
            else:
                uploaded_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="w_pdf")
                if uploaded_file and st.button("Convert to PDF", key="btn_w_pdf"):
                    try:
                        from docx2pdf import convert as docx_convert
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_in, tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_out:
                            tmp_in.write(uploaded_file.read())
                            tmp_in_path, tmp_out_path = tmp_in.name, tmp_out.name
                        docx_convert(tmp_in_path, tmp_out_path)
                        with open(tmp_out_path, "rb") as f_out: pdf_bytes = f_out.read()
                        st.success("Word converted to PDF successfully!")
                        st.download_button("Download PDF", data=pdf_bytes, file_name="converted_word.pdf", mime="application/pdf")
                    except Exception as e: st.error(f"Conversion error: {e}")
                    finally:
                        if 'tmp_in_path' in locals() and os.path.exists(tmp_in_path): os.remove(tmp_in_path)
                        if 'tmp_out_path' in locals() and os.path.exists(tmp_out_path): os.remove(tmp_out_path)
    if t14:
        with st.expander("14. Convert PDF to Word Document"):
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_w")
            if uploaded_file and st.button("Convert to Word", key="btn_pdf_w"):
                try:
                    from pdf2docx import Converter
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in, tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_out:
                        tmp_in.write(uploaded_file.read())
                        tmp_in_path, tmp_out_path = tmp_in.name, tmp_out.name
                    cv = Converter(tmp_in_path)
                    cv.convert(tmp_out_path, start=0, end=None); cv.close()
                    with open(tmp_out_path, "rb") as f_out: docx_bytes = f_out.read()
                    st.success("PDF converted to Word successfully!")
                    st.download_button("Download Word Document", data=docx_bytes, file_name="converted_pdf.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e: st.error(f"Error: {e}")
                finally:
                    if 'tmp_in_path' in locals() and os.path.exists(tmp_in_path): os.remove(tmp_in_path)
                    if 'tmp_out_path' in locals() and os.path.exists(tmp_out_path): os.remove(tmp_out_path)
    if t55:
        with st.expander("55. Convert PDF to eBook (EPUB)"):
            ebook_pdf = st.file_uploader("Upload PDF Document", type=["pdf"], key="ebook_pdf")
            c1, c2 = st.columns(2)
            ebook_title = c1.text_input("eBook Title", value="Rajya Sabha Document", key="ebook_title")
            ebook_author = c2.text_input("Author Name", value="Secretariat", key="ebook_author")
            if ebook_pdf and st.button("Convert to EPUB", key="btn_epub"):
                try:
                    import ebooklib; from ebooklib import epub; import fitz
                    with st.spinner("Extracting text and compiling eBook..."):
                        doc = fitz.open(stream=ebook_pdf.read(), filetype="pdf")
                        book = epub.EpubBook()
                        book.set_identifier(f"rs-epub-{random.randint(1000,9999)}")
                        book.set_title(ebook_title); book.set_language('en'); book.add_author(ebook_author)
                        full_html = ""
                        for page in doc:
                            for p in page.get_text("text").split('\n\n'):
                                clean_p = p.replace('\n', ' ').strip()
                                if clean_p:
                                    clean_p = clean_p.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                                    full_html += f"<p>{clean_p}</p>\n"
                        c1 = epub.EpubHtml(title='Document Content', file_name='content.xhtml', lang='en')
                        c1.content = f'<h1>{ebook_title}</h1>' + full_html
                        book.add_item(c1)
                        book.toc = (epub.Link('content.xhtml', 'Content', 'content'),)
                        book.add_item(epub.EpubNcx()); book.add_item(epub.EpubNav())
                        book.spine = ['nav', c1]
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".epub") as tmp_epub:
                            tmp_epub_path = tmp_epub.name
                        epub.write_epub(tmp_epub_path, book)
                        with open(tmp_epub_path, "rb") as f: epub_bytes = f.read()
                        st.success("eBook compiled successfully!")
                        st.download_button("📥 Download EPUB eBook", data=epub_bytes, file_name=f"{ebook_title.replace(' ', '_')}.epub", mime="application/epub+zip")
                except Exception as e: st.error(f"Error compiling eBook: {e}")
                finally:
                    if 'tmp_epub_path' in locals() and os.path.exists(tmp_epub_path): os.remove(tmp_epub_path)

# ==========================================
# CATEGORY 3: DATA & TEXT EXTRACTION
# ==========================================
if any([t15, t16, t17, t18, t19]):
    st.markdown("## 📊 Data & Text Extraction")
    if t15:
        with st.expander("15. Scan PDF to Notepad (Full OCR - Scanned Docs)"):
            ocr_pdf = st.file_uploader("Upload Scanned PDF", type=["pdf"], key="ocr_pdf")
            if ocr_pdf and st.button("Run Full OCR", key="btn_ocr"):
                try:
                    import pytesseract; from pdf2image import convert_from_bytes
                    images = convert_from_bytes(ocr_pdf.read(), dpi=200)
                    text_out = ""
                    for i, img in enumerate(images):
                        text_out += f"--- Page {i+1} ---\n" + pytesseract.image_to_string(img) + "\n\n"
                    st.success("OCR completed successfully!")
                    st.download_button("Download OCR Text", data=text_out, file_name="ocr_extracted.txt", mime="text/plain")
                except Exception as e: st.error(f"OCR Error: {e}")
    if t16:
        with st.expander("16. Fast Text Extract (Digital PDFs Only)"):
            uploaded_file = st.file_uploader("Upload Digital PDF", type=["pdf"], key="txt_pdf")
            if uploaded_file and st.button("Extract Text", key="btn_txt_pdf"):
                reader = PdfReader(uploaded_file)
                text_content = "".join([f"--- Page {i+1} ---\n{p.extract_text()}\n\n" for i, p in enumerate(reader.pages) if p.extract_text()])
                st.success("Text extracted successfully!")
                st.download_button("Download Text File", data=text_content, file_name="extracted_text.txt", mime="text/plain")
    if t17:
        with st.expander("17. Extract Text from Word to Notepad"):
            uploaded_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="w_txt")
            if uploaded_file and st.button("Extract Word Text", key="btn_w_txt"):
                doc = docx.Document(uploaded_file)
                text_content = "\n".join([p.text for p in doc.paragraphs])
                st.success("Text extracted successfully!")
                st.download_button("Download Text File", data=text_content, file_name="word_text.txt", mime="text/plain")
    if t18:
        with st.expander("18. Extract All Embedded Images from PDF"):
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="img_ext")
            if uploaded_file and st.button("Extract Images", key="btn_img_ext"):
                try:
                    reader = PdfReader(uploaded_file)
                    zip_buffer = io.BytesIO()
                    count = 0
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for page_num, page in enumerate(reader.pages):
                            for img_idx, img_obj in enumerate(page.images):
                                count += 1
                                zip_file.writestr(f"page_{page_num+1}_img_{img_idx+1}_{img_obj.name}", img_obj.data)
                    if count > 0:
                        st.success(f"Extracted {count} embedded images!")
                        st.download_button("📥 Download Images (ZIP)", data=zip_buffer.getvalue(), file_name="extracted_images.zip", mime="application/zip")
                    else: st.warning("No embedded images found.")
                except Exception as e: st.error(f"Error: {e}")
    if t19:
        with st.expander("19. Extract PDF Tables to Excel (.csv)"):
            uploaded_file = st.file_uploader("Upload PDF with Tables", type=["pdf"], key="tbl_pdf")
            if uploaded_file and st.button("Extract Tables", key="btn_tbl_pdf"):
                try:
                    import pdfplumber
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(uploaded_file.read()); tmp_path = tmp.name
                    csv_buffer = io.StringIO()
                    writer = csv.writer(csv_buffer)
                    found = False
                    with pdfplumber.open(tmp_path) as pdf:
                        for page in pdf.pages:
                            for table in page.extract_tables():
                                found = True
                                for row in table: writer.writerow([cell if cell is not None else "" for cell in row])
                                writer.writerow([])
                    if found:
                        st.success("Tables extracted successfully!")
                        st.download_button("Download Table (.csv)", data=csv_buffer.getvalue(), file_name="extracted_tables.csv", mime="application/csv")
                    else: st.warning("No structured tables found.")
                except Exception as e: st.error(f"Error: {e}")
                finally:
                    if 'tmp_path' in locals() and os.path.exists(tmp_path): os.remove(tmp_path)

# ==========================================
# CATEGORY 4: SECURITY & DOCUMENT POLISH
# ==========================================
if any([t20, t21, t22, t23, t24, t25, t26, t27]):
    st.markdown("## 🔒 Security & Document Polish")
    if t20:
        with st.expander("20. Password Protect / Encrypt a PDF"):
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="lock_pdf")
            pwd = st.text_input("Enter Password", type="password", key="lock_pwd")
            if uploaded_file and pwd and st.button("Protect PDF", key="btn_lock"):
                reader = PdfReader(uploaded_file)
                writer = PdfWriter()
                for page in reader.pages: writer.add_page(page)
                writer.encrypt(pwd)
                out = io.BytesIO()
                writer.write(out); out.seek(0)
                st.success("PDF protected successfully!")
                st.download_button("Download Protected PDF", data=out, file_name="protected.pdf", mime="application/pdf")
    if t21:
        with st.expander("21. Unlock / Remove PDF Password"):
            uploaded_file = st.file_uploader("Upload Locked PDF", type=["pdf"], key="unlock_pdf")
            pwd = st.text_input("Enter Current Password", type="password", key="unlock_pwd")
            if uploaded_file and pwd and st.button("Unlock PDF", key="btn_unlock"):
                try:
                    reader = PdfReader(uploaded_file)
                    if reader.is_encrypted: reader.decrypt(pwd)
                    writer = PdfWriter()
                    for page in reader.pages: writer.add_page(page)
                    out = io.BytesIO()
                    writer.write(out); out.seek(0)
                    st.success("PDF unlocked successfully!")
                    st.download_button("Download Unlocked PDF", data=out, file_name="unlocked.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t22:
        with st.expander("22. Apply Custom Watermark / Stamp to PDF"):
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="wm_pdf")
            wm_text = st.text_input("Watermark Text", value="CONFIDENTIAL", key="wm_txt")
            if uploaded_file and wm_text and st.button("Apply Watermark", key="btn_wm"):
                reader = PdfReader(uploaded_file)
                writer = PdfWriter()
                page_width, page_height = A4
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=A4)
                can.setFont("Helvetica-Bold", 65)
                can.setFillColorRGB(0.5, 0.5, 0.5, alpha=0.3)
                can.translate(page_width / 2, page_height / 2); can.rotate(45)
                can.drawCentredString(0, 0, wm_text); can.save(); packet.seek(0)
                wm_page = PdfReader(packet).pages[0]
                for page in reader.pages:
                    page.merge_page(wm_page); writer.add_page(page)
                out = io.BytesIO()
                writer.write(out); out.seek(0)
                st.success("Watermark applied successfully!")
                st.download_button("Download Watermarked PDF", data=out, file_name="watermarked.pdf", mime="application/pdf")
    if t23:
        with st.expander("23. Add Page Numbers (Bates Numbering)"):
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="bates_pdf")
            if uploaded_file and st.button("Add Page Numbers", key="btn_bates"):
                reader = PdfReader(uploaded_file)
                writer = PdfWriter()
                total_pages = len(reader.pages)
                for i, page in enumerate(reader.pages):
                    packet = io.BytesIO()
                    can = canvas.Canvas(packet, pagesize=(float(page.mediabox.width), float(page.mediabox.height)))
                    can.setFont("Helvetica", 10)
                    can.drawCentredString(float(page.mediabox.width) / 2, 20, f"Page {i + 1} of {total_pages}")
                    can.save(); packet.seek(0)
                    page.merge_page(PdfReader(packet).pages[0])
                    writer.add_page(page)
                out = io.BytesIO()
                writer.write(out); out.seek(0)
                st.success("Page numbers added successfully!")
                st.download_button("Download Numbered PDF", data=out, file_name="numbered.pdf", mime="application/pdf")
    if t24:
        with st.expander("24. Insert Blank Pages (For Duplex Printing)"):
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="blank_pdf")
            insert_at = st.number_input("Insert blank page AFTER page number (0 for beginning)", min_value=0, value=0, step=1, key="blank_num")
            if uploaded_file and st.button("Insert Blank Page", key="btn_blank"):
                reader = PdfReader(uploaded_file)
                writer = PdfWriter()
                for i, page in enumerate(reader.pages):
                    if i == insert_at and insert_at == 0: writer.add_blank_page(width=page.mediabox.width, height=page.mediabox.height)
                    writer.add_page(page)
                    if i + 1 == insert_at: writer.add_blank_page(width=page.mediabox.width, height=page.mediabox.height)
                out = io.BytesIO()
                writer.write(out); out.seek(0)
                st.success("Blank page inserted successfully!")
                st.download_button("Download Updated PDF", data=out, file_name="blank_page_output.pdf", mime="application/pdf")
    if t25:
        with st.expander("25. Rotate PDF Pages (90°)"):
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="rot_pdf")
            if uploaded_file and st.button("Rotate Pages 90°", key="btn_rot"):
                reader = PdfReader(uploaded_file)
                writer = PdfWriter()
                for page in reader.pages:
                    page.rotate(90); writer.add_page(page)
                out = io.BytesIO()
                writer.write(out); out.seek(0)
                st.success("PDF rotated successfully!")
                st.download_button("Download Rotated PDF", data=out.getvalue(), file_name="rotated.pdf", mime="application/pdf")
    if t26:
        with st.expander("26. Compress / Optimize PDF File Size"):
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="comp_pdf")
            if uploaded_file and st.button("Compress PDF Streams", key="btn_comp"):
                reader = PdfReader(uploaded_file)
                writer = PdfWriter()
                for page in reader.pages: writer.add_page(page)
                for page in writer.pages: page.compress_content_streams()
                out = io.BytesIO()
                writer.write(out); out.seek(0)
                st.success("PDF streams compressed successfully!")
                st.download_button("Download Compressed PDF", data=out, file_name="compressed.pdf", mime="application/pdf")
    if t27:
        with st.expander("27. View Hidden PDF Metadata"):
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="meta_pdf")
            if uploaded_file and st.button("View Metadata", key="btn_meta"):
                meta = PdfReader(uploaded_file).metadata
                if meta:
                    meta_str = "\n".join([f"{k.strip('/')}: {v}" for k, v in meta.items()])
                    st.text_area("PDF Metadata", value=meta_str, height=200)
                else: st.info("No metadata found.")

# ==========================================
# CATEGORY 5: IMAGE UTILITIES
# ==========================================
if any([t28, t29]):
    st.markdown("## 🖼️ Image Utilities")
    if t28:
        with st.expander("28. Remove Image Background & Change Color (AI Offline)"):
            bg_img_file = st.file_uploader("Upload Image", type=["png", "jpg", "jpeg"], key="bg_rem_img")
            c1, c2 = st.columns(2)
            bg_option = c1.selectbox("Select New Background", ["Transparent", "White", "Red", "Pink", "Green", "Yellow", "Custom Color"], key="bg_rem_opt")
            custom_color = "#FFFFFF"
            if bg_option == "Custom Color": custom_color = c2.color_picker("Pick a Custom Color", "#3498db", key="bg_rem_color")
            if bg_img_file and st.button("Process Image", key="btn_bg_rem"):
                try:
                    from rembg import remove as rembg_remove
                    with st.spinner("AI is analyzing and replacing background..."):
                        output_bytes = rembg_remove(bg_img_file.read())
                        if bg_option == "Transparent":
                            final_output, file_ext, mime_val = output_bytes, "png", "image/png"
                        else:
                            color_map = {"White": "#FFFFFF", "Red": "#FF0000", "Pink": "#FFC0CB", "Green": "#008000", "Yellow": "#FFFF00"}
                            selected_hex = custom_color if bg_option == "Custom Color" else color_map[bg_option]
                            transparent_img = Image.open(io.BytesIO(output_bytes)).convert("RGBA")
                            solid_bg = Image.new("RGBA", transparent_img.size, selected_hex)
                            solid_bg.paste(transparent_img, (0, 0), transparent_img)
                            final_io = io.BytesIO()
                            solid_bg.convert("RGB").save(final_io, format="JPEG", quality=95)
                            final_output, file_ext, mime_val = final_io.getvalue(), "jpg", "image/jpeg"
                    st.success("Background processed successfully!")
                    st.download_button(f"Download Processed Image (.{file_ext})", data=final_output, file_name=f"bg_processed.{file_ext}", mime=mime_val)
                except Exception as e: st.error(f"Error: {e}")
    if t29:
        with st.expander("29. Resize & Compress Image"):
            res_img_file = st.file_uploader("Upload Image to Resize/Compress", type=["png", "jpg", "jpeg"], key="res_img")
            if res_img_file:
                orig_img = Image.open(res_img_file)
                st.info(f"Original Dimensions: **{orig_img.size[0]} x {orig_img.size[1]} pixels**")
                c1, c2 = st.columns(2)
                new_width = c1.number_input("New Width (Pixels)", min_value=10, value=orig_img.size[0], key="res_w")
                new_height = c2.number_input("New Height (Pixels)", min_value=10, value=orig_img.size[1], key="res_h")
                quality = st.slider("Compression Quality", min_value=1, max_value=100, value=85, key="res_q")
                if st.button("Resize & Compress Image", key="btn_res"):
                    resized_img = orig_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    out_img = io.BytesIO()
                    if orig_img.format == "PNG" and quality > 90:
                        resized_img.save(out_img, format="PNG", optimize=True)
                        mime_type, ext = "image/png", "png"
                    else:
                        resized_img.convert("RGB").save(out_img, format="JPEG", quality=quality, optimize=True)
                        mime_type, ext = "image/jpeg", "jpg"
                    st.success("Image resized and compressed successfully!")
                    st.download_button("Download Processed Image", data=out_img.getvalue(), file_name=f"processed_image.{ext}", mime=mime_type)

# ==========================================
# CATEGORY 6: RAJBHASHA & HINDI TOOLS
# ==========================================
if any([t30]):
    st.markdown("## 🇮🇳 Rajbhasha & Hindi Tools")
    if t30:
        with st.expander("30. KrutiDev 010 to Unicode (Mangal) Converter"):
            kruti_text = st.text_area("Paste KrutiDev Text Here:", height=150, key="kruti_input")
            if st.button("Convert to Unicode", key="btn_convert_kruti"):
                if kruti_text.strip():
                    unicode_text = convert_krutidev_to_unicode(kruti_text)
                    st.success("Conversion Successful!")
                    st.text_area("Unicode (Mangal) Output:", value=unicode_text, height=150, key="unicode_output")
                else: st.warning("Please enter text to convert.")

# ==========================================
# CATEGORY 7: ADVANCED & SMART UTILITIES
# ==========================================
if any([t31, t32, t33, t34]):
    st.markdown("## 🚀 Advanced & Smart Utilities")
    if t31:
        with st.expander("31. True PDF Redaction (ऑटोमैटिक ब्लैकआउट)"):
            redact_pdf = st.file_uploader("Upload PDF to Redact", type=["pdf"], key="redact_pdf")
            redact_text = st.text_input("Enter Exact Text to Redact (Hide)", key="redact_text")
            if redact_pdf and redact_text and st.button("Apply Permanent Blackout", key="btn_redact"):
                try:
                    import fitz
                    with fitz.open(stream=redact_pdf.read(), filetype="pdf") as doc:
                        found_count = 0
                        for page in doc:
                            areas = page.search_for(redact_text)
                            for rect in areas:
                                found_count += 1
                                page.add_redact_annot(rect, fill=(0, 0, 0))
                            page.apply_redactions()
                        if found_count > 0:
                            out_pdf = io.BytesIO()
                            doc.save(out_pdf); out_pdf.seek(0)
                            st.success(f"Success! Redacted '{redact_text}' {found_count} times.")
                            st.download_button("Download Redacted PDF", data=out_pdf, file_name="redacted_document.pdf", mime="application/pdf")
                        else: st.warning(f"Could not find '{redact_text}' in this PDF.")
                except Exception as e: st.error(f"Error: {e}")
    if t32:
        with st.expander("32. Bulk Generator (Excel/CSV to Word Mail Merge)"):
            template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"], key="mm_template")
            data_file = st.file_uploader("Upload Data File (.csv)", type=["csv"], key="mm_data")
            if template_file and data_file and st.button("Generate Bulk Documents", key="btn_mm"):
                try:
                    reader = csv.DictReader(io.StringIO(data_file.getvalue().decode("utf-8")))
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for i, row in enumerate(reader):
                            doc = docx.Document(template_file)
                            for p in doc.paragraphs:
                                for key, val in row.items():
                                    ph = f"<<{key}>>"
                                    if ph in p.text:
                                        for run in p.runs:
                                            if ph in run.text: run.text = run.text.replace(ph, str(val))
                                        if ph in p.text: p.text = p.text.replace(ph, str(val))
                            for table in doc.tables:
                                for r in table.rows:
                                    for cell in r.cells:
                                        for p in cell.paragraphs:
                                            for key, val in row.items():
                                                ph = f"<<{key}>>"
                                                if ph in p.text:
                                                    for run in p.runs:
                                                        if ph in run.text: run.text = run.text.replace(ph, str(val))
                                                    if ph in p.text: p.text = p.text.replace(ph, str(val))
                            doc_io = io.BytesIO()
                            doc.save(doc_io)
                            zip_file.writestr(f"Document_{i+1}.docx", doc_io.getvalue())
                            template_file.seek(0)
                    st.success("Bulk documents generated successfully!")
                    st.download_button("Download All (ZIP)", data=zip_buffer.getvalue(), file_name="Bulk_Documents.zip", mime="application/zip")
                except Exception as e: st.error(f"Error: {e}")
    if t33:
        with st.expander("33. Add Table of Contents (Clickable Bookmarks) to PDF"):
            toc_pdf = st.file_uploader("Upload PDF", type=["pdf"], key="toc_pdf")
            toc_text = st.text_area("Format: Page Number, Title on each line (e.g., 1, Introduction)", height=150, key="toc_text")
            if toc_pdf and toc_text and st.button("Add Bookmarks", key="btn_toc"):
                try:
                    reader = PdfReader(toc_pdf)
                    writer = PdfWriter()
                    for page in reader.pages: writer.add_page(page)
                    for line in toc_text.strip().split('\n'):
                        if ',' in line:
                            p_str, title = line.split(',', 1)
                            p_num = int(p_str.strip()) - 1
                            if 0 <= p_num < len(reader.pages): writer.add_outline_item(title.strip(), p_num)
                    out_pdf = io.BytesIO()
                    writer.write(out_pdf); out_pdf.seek(0)
                    st.success("Bookmarks added successfully!")
                    st.download_button("Download Bookmarked PDF", data=out_pdf, file_name="bookmarked.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t34:
        with st.expander("34. PDF Grayscale / B&W Converter (Printer Toner Saver)"):
            bw_pdf_file = st.file_uploader("Upload Color PDF", type=["pdf"], key="bw_pdf")
            dpi_val = st.slider("Print Quality (DPI)", min_value=72, max_value=300, value=150, step=10, key="bw_dpi")
            if bw_pdf_file and st.button("Convert to Grayscale", key="btn_bw"):
                try:
                    import fitz
                    with fitz.open(stream=bw_pdf_file.read(), filetype="pdf") as doc, fitz.open() as out_pdf:
                        for page in doc:
                            pix = page.get_pixmap(colorspace=fitz.csGRAY, dpi=dpi_val, alpha=False)
                            new_page = out_pdf.new_page(width=page.rect.width, height=page.rect.height)
                            new_page.insert_image(page.rect, pixmap=pix)
                        out_bytes = io.BytesIO()
                        out_pdf.save(out_bytes); out_bytes.seek(0)
                        st.success("Successfully converted to Grayscale!")
                        st.download_button("Download Grayscale PDF", data=out_bytes, file_name="Grayscale_Doc.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")

# ==========================================
# CATEGORY 8: ENTERPRISE UTILITIES
# ==========================================
if any([t35, t36, t37, t38, t39, t40]):
    st.markdown("## 🏢 Enterprise & Commercial Grade Utilities")
    if t35:
        with st.expander("35. Visual PDF Comparison (Draft Diff Tool)"):
            pdf1_file = st.file_uploader("Upload Original PDF", type=["pdf"], key="diff_pdf1")
            pdf2_file = st.file_uploader("Upload Modified PDF", type=["pdf"], key="diff_pdf2")
            if pdf1_file and pdf2_file and st.button("Compare Texts", key="btn_diff"):
                t1_txt = "\n".join([page.extract_text() for page in PdfReader(pdf1_file).pages if page.extract_text()])
                t2_txt = "\n".join([page.extract_text() for page in PdfReader(pdf2_file).pages if page.extract_text()])
                html_diff = difflib.HtmlDiff().make_file(t1_txt.splitlines(), t2_txt.splitlines(), "Original", "Modified")
                st.success("Comparison Complete!")
                components.html(html_diff, height=600, scrolling=True)
    if t36:
        with st.expander("36. Searchable 'Sandwich' PDF (Invisible OCR Overlay)"):
            ocr_pdf_file = st.file_uploader("Upload Scanned PDF to make Searchable", type=["pdf"], key="sandwich_pdf")
            if ocr_pdf_file and st.button("Create Searchable PDF", key="btn_sandwich"):
                try:
                    import pytesseract; from pdf2image import convert_from_bytes
                    images = convert_from_bytes(ocr_pdf_file.read(), dpi=200)
                    merger = PdfWriter()
                    def process_ocr(img): return pytesseract.image_to_pdf_or_hocr(img, extension='pdf')
                    with concurrent.futures.ThreadPoolExecutor() as executor:
                        pdf_bytes_list = list(executor.map(process_ocr, images))
                    for pdf_bytes in pdf_bytes_list:
                        merger.add_page(PdfReader(io.BytesIO(pdf_bytes)).pages[0])
                    out_pdf = io.BytesIO()
                    merger.write(out_pdf); out_pdf.seek(0)
                    st.success("Searchable PDF created successfully!")
                    st.download_button("Download Searchable PDF", data=out_pdf, file_name="Searchable_Doc.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t37:
        with st.expander("37. Digital Facsimile Signature & Stamp Placer"):
            stamp_pdf = st.file_uploader("Upload PDF Document", type=["pdf"], key="stamp_pdf")
            stamp_img = st.file_uploader("Upload Stamp/Signature (PNG)", type=["png", "jpg", "jpeg"], key="stamp_img")
            c1, c2 = st.columns(2)
            stamp_page = c1.number_input("Page Number (0 for All Pages)", min_value=0, value=1, step=1, key="stamp_page_num")
            stamp_pos = c2.selectbox("Position", ["Bottom-Right", "Bottom-Left", "Top-Right", "Top-Left", "Center"], key="stamp_pos")
            if stamp_pdf and stamp_img and st.button("Place Stamp/Signature", key="btn_place_stamp"):
                try:
                    import fitz
                    with fitz.open(stream=stamp_pdf.read(), filetype="pdf") as doc:
                        img_bytes = stamp_img.read()
                        pages_to_stamp = range(len(doc)) if stamp_page == 0 else [stamp_page - 1]
                        for p_num in pages_to_stamp:
                            if 0 <= p_num < len(doc):
                                page = doc[p_num]; rect = page.rect
                                sw, sh, m = 150, 75, 30
                                if stamp_pos == "Bottom-Right": tr = fitz.Rect(rect.width - sw - m, rect.height - sh - m, rect.width - m, rect.height - m)
                                elif stamp_pos == "Bottom-Left": tr = fitz.Rect(m, rect.height - sh - m, m + sw, rect.height - m)
                                elif stamp_pos == "Top-Right": tr = fitz.Rect(rect.width - sw - m, m, rect.width - m, m + sh)
                                elif stamp_pos == "Top-Left": tr = fitz.Rect(m, m, m + sw, m + sh)
                                else: tr = fitz.Rect((rect.width - sw)/2, (rect.height - sh)/2, (rect.width + sw)/2, (rect.height + sh)/2)
                                page.insert_image(tr, stream=img_bytes)
                        out_pdf = io.BytesIO()
                        doc.save(out_pdf); out_pdf.seek(0)
                        st.success("Stamp/Signature placed successfully!")
                        st.download_button("Download Stamped PDF", data=out_pdf, file_name="Stamped.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t38:
        with st.expander("38. N-Up & Booklet Layout Maker (Multiple Pages on 1 Sheet)"):
            nup_pdf = st.file_uploader("Upload PDF to Format", type=["pdf"], key="nup_pdf")
            nup_layout = st.selectbox("Select Layout", ["2-Up (2 Pages per Sheet)", "4-Up (4 Pages per Sheet)", "8-Up (8 Pages per Sheet)"], key="nup_layout")
            if nup_pdf and st.button("Generate Layout", key="btn_nup"):
                try:
                    import fitz
                    with fitz.open(stream=nup_pdf.read(), filetype="pdf") as doc, fitz.open() as out_doc:
                        cols, rows = (2, 1) if "2-Up" in nup_layout else ((2, 2) if "4-Up" in nup_layout else (4, 2))
                        chunk_size = cols * rows
                        for start_idx in range(0, len(doc), chunk_size):
                            p_ref = doc[start_idx]
                            w, h = p_ref.rect.width, p_ref.rect.height
                            new_page = out_doc.new_page(width=w * cols, height=h * rows)
                            for i in range(chunk_size):
                                doc_idx = start_idx + i
                                if doc_idx < len(doc):
                                    col, row = i % cols, i // cols
                                    tr = fitz.Rect(col * w, row * h, (col + 1) * w, (row + 1) * h)
                                    new_page.show_pdf_page(tr, doc, doc_idx)
                        out_pdf = io.BytesIO()
                        out_doc.save(out_pdf); out_pdf.seek(0)
                        st.success("Layout created successfully!")
                        st.download_button("Download N-Up PDF", data=out_pdf, file_name="N_Up_Format.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t39:
        with st.expander("39. Institutional Bates Stamping (Custom Prefix Legal Numbering)"):
            bates_adv_pdf = st.file_uploader("Upload Document for Stamping", type=["pdf"], key="bates_adv_pdf")
            c1, c2 = st.columns(2)
            b_prefix = c1.text_input("Bates Prefix", value="RS/SYS/2026/", key="b_prefix")
            b_start = c2.number_input("Starting Number", min_value=1, value=1, step=1, key="b_start")
            if bates_adv_pdf and st.button("Apply Bates Stamping", key="btn_bates_adv"):
                try:
                    import fitz
                    with fitz.open(stream=bates_adv_pdf.read(), filetype="pdf") as doc:
                        for i, page in enumerate(doc):
                            b_num = f"{b_prefix}{str(b_start + i).zfill(4)}"
                            page.insert_text(fitz.Point(page.rect.width - 150, 30), b_num, fontsize=11, fontname="helv", color=(1, 0, 0))
                        out_pdf = io.BytesIO()
                        doc.save(out_pdf); out_pdf.seek(0)
                        st.success("Bates stamping completed!")
                        st.download_button("Download Stamped PDF", data=out_pdf, file_name="Bates_Doc.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t40:
        with st.expander("40. PDF Auto-Crop & White Margin Trimmer"):
            crop_pdf_file = st.file_uploader("Upload PDF to Crop", type=["pdf"], key="crop_pdf")
            if crop_pdf_file and st.button("Auto-Crop Margins", key="btn_crop"):
                try:
                    import fitz
                    with fitz.open(stream=crop_pdf_file.read(), filetype="pdf") as doc:
                        for page in doc:
                            tr = page.get_text("rect")
                            if not tr.is_empty and not tr.is_infinite:
                                cb = fitz.Rect(max(0, tr.x0 - 10), max(0, tr.y0 - 10), min(page.rect.width, tr.x1 + 10), min(page.rect.height, tr.y1 + 10))
                                page.set_cropbox(cb)
                        out_pdf = io.BytesIO()
                        doc.save(out_pdf); out_pdf.seek(0)
                        st.success("Margins cropped successfully!")
                        st.download_button("Download Cropped PDF", data=out_pdf, file_name="Cropped.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")

# ==========================================
# CATEGORY 9: SECRETARIAT SPECIFIC
# ==========================================
if any([t41, t42, t43, t44, t45, t46, t47, t54]):
    st.markdown("## 🌟 Secretariat Specific Utilities")
    if t41:
        with st.expander("41. Parliamentary & Admin Glossary (सचिवालयीन शब्दकोश)"):
            glossary = {
                "Adjournment": "स्थगन", "Adjournment sine die": "अनिश्चित काल के लिए स्थगन", "Prorogation": "सत्रावसान",
                "Quorum": "गणपूर्ति", "Whip": "सचेतक", "Zero Hour": "शून्य काल", "Question Hour": "प्रश्न काल",
                "Starred Question": "तारांकित प्रश्न", "Unstarred Question": "अतारांकित प्रश्न", "Point of Order": "औचित्य प्रश्न",
                "Breach of Privilege": "विशेषाधिकार हनन", "Laying on the Table": "पटल पर रखना", "Casting Vote": "निर्णायक मत",
                "Resolution": "संकल्प", "Motion": "प्रस्ताव", "Bill": "विधेयक", "Act": "अधिनियम", "Amendment": "संशोधन",
                "Office Memorandum (O.M.)": "कार्यालय ज्ञापन", "Circular": "परिपत्र", "Notification": "अधिसूचना",
                "Gazette": "राजपत्र", "Endorsement": "पृष्ठांकन", "Annexure": "अनुलग्नक", "Appendix": "परिशिष्ट",
                "Minutes of Meeting": "कार्यवृत्त", "Agenda": "कार्यसूची", "Ex-officio": "पदेन", "Vetting": "संवीक्षा",
                "Concurrence": "सहमति", "Delegation of Power": "शक्तियों का प्रत्यायोजन", "Sanction": "स्वीकृति"
            }
            term = st.text_input("Type term to search:", key="gloss_search")
            if term:
                matches = {k: v for k, v in glossary.items() if term.lower() in k.lower()}
                for k, v in matches.items(): st.success(f"**{k}** ➔ {v}")
    if t42:
        with st.expander("42. Make PDF Look Scanned (डिजिटल को 'स्कैन-लुक' दें)"):
            fake_scan_pdf = st.file_uploader("Upload Clean PDF", type=["pdf"], key="fake_scan_pdf")
            if fake_scan_pdf and st.button("Apply Scan Effect", key="btn_fake_scan"):
                try:
                    import fitz
                    with fitz.open(stream=fake_scan_pdf.read(), filetype="pdf") as doc:
                        out_pdf_bytes = io.BytesIO()
                        merger = PdfWriter()
                        for page in doc:
                            pix = page.get_pixmap(dpi=150)
                            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
                            img = img.rotate(random.uniform(-1.2, 1.2), fillcolor=255, expand=False)
                            img = ImageEnhance.Contrast(img).enhance(1.8)
                            img = ImageEnhance.Brightness(img).enhance(0.9)
                            tpdf = io.BytesIO(); img.save(tpdf, format="PDF"); tpdf.seek(0)
                            merger.add_page(PdfReader(tpdf).pages[0])
                        merger.write(out_pdf_bytes); out_pdf_bytes.seek(0)
                        st.success("Scan effect applied successfully!")
                        st.download_button("Download Scanned PDF", data=out_pdf_bytes, file_name="Scanned_Copy.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t43:
        with st.expander("43. Quick Dak/Diary Receipt Stamper (ई-डायरी मुहर)"):
            dak_pdf_file = st.file_uploader("Upload Incoming Dak (PDF)", type=["pdf"], key="dak_pdf_file")
            c1, c2 = st.columns(2)
            diary_no = c1.text_input("Enter Diary No.", value="1234/Sys/26", key="diary_no")
            diary_date = c2.text_input("Enter Date", value="03-09-2026", key="diary_date")
            if dak_pdf_file and st.button("Apply Diary Stamp", key="btn_dak"):
                try:
                    import fitz
                    with fitz.open(stream=dak_pdf_file.read(), filetype="pdf") as doc:
                        p = doc[0]; rw, rh, m = 240, 80, 20
                        r = fitz.Rect(p.rect.width - rw - m, m, p.rect.width - m, m + rh)
                        p.draw_rect(r, color=(1, 0, 0), width=2)
                        p.insert_text(fitz.Point(r.x0 + 10, r.y0 + 20), "RECEIVED - RAJYA SABHA SEC.", color=(1,0,0), fontsize=10, fontname="Helvetica-Bold")
                        p.insert_text(fitz.Point(r.x0 + 10, r.y0 + 45), f"Diary No : {diary_no}", color=(1,0,0), fontsize=10, fontname="Helvetica")
                        p.insert_text(fitz.Point(r.x0 + 10, r.y0 + 65), f"Date     : {diary_date}", color=(1,0,0), fontsize=10, fontname="Helvetica")
                        out_pdf = io.BytesIO()
                        doc.save(out_pdf); out_pdf.seek(0)
                        st.success("Diary stamp applied!")
                        st.download_button("Download Stamped Dak", data=out_pdf, file_name="Diarized.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t44:
        with st.expander("44. PDF A4 Page Standardizer (प्रिंटर पेपर-साइज फिक्सर)"):
            a4_pdf_file = st.file_uploader("Upload Mixed-Size PDF", type=["pdf"], key="a4_pdf")
            if a4_pdf_file and st.button("Standardize to A4", key="btn_a4"):
                try:
                    import fitz
                    with fitz.open(stream=a4_pdf_file.read(), filetype="pdf") as doc, fitz.open() as out_doc:
                        a4_w, a4_h = fitz.paper_size("a4")
                        for i in range(len(doc)):
                            np = out_doc.new_page(width=a4_w, height=a4_h)
                            np.show_pdf_page(np.rect, doc, i)
                        out_pdf = io.BytesIO()
                        out_doc.save(out_pdf); out_pdf.seek(0)
                        st.success("Standardized to A4!")
                        st.download_button("Download A4 PDF", data=out_pdf, file_name="Standardized_A4.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t45:
        with st.expander("45. Scanned Book / Landscape Page Slicer (हाफ-पेज स्प्लिटर)"):
            slice_pdf_file = st.file_uploader("Upload Wide/Landscape PDF", type=["pdf"], key="slice_pdf")
            if slice_pdf_file and st.button("Slice in Half", key="btn_slice"):
                try:
                    import fitz
                    with fitz.open(stream=slice_pdf_file.read(), filetype="pdf") as doc, fitz.open() as out_doc:
                        for i in range(len(doc)):
                            p = doc[i]
                            if p.rect.width > p.rect.height:
                                hw = p.rect.width / 2
                                p1 = out_doc.new_page(width=hw, height=p.rect.height)
                                p1.show_pdf_page(p1.rect, doc, i, clip=fitz.Rect(0, 0, hw, p.rect.height))
                                p2 = out_doc.new_page(width=hw, height=p.rect.height)
                                p2.show_pdf_page(p2.rect, doc, i, clip=fitz.Rect(hw, 0, p.rect.width, p.rect.height))
                            else: out_doc.insert_pdf(doc, from_page=i, to_page=i)
                        out_pdf = io.BytesIO()
                        out_doc.save(out_pdf); out_pdf.seek(0)
                        st.success("Sliced successfully!")
                        st.download_button("Download Sliced PDF", data=out_pdf, file_name="Sliced_Portrait.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t46:
        with st.expander("46. Offline Telephone & Room Directory (सचिवालय टेलीफोन निर्देशिका)"):
            tel_directory = [
                {"name": "Shri CP Radhakrishnan (Hon'ble Chairman, Rajya Sabha)", "office_phone": "23094953, 23094954", "room": "RS-14, PH"},
                {"name": "Shri Harivansh (Hon'ble Deputy Chairman)", "office_phone": "23083029, 23083030", "room": "RS-07, PH"},
                {"name": "Shri P.C. Mody (Secretary-General)", "office_phone": "23083035, 23083036", "room": "RS-08, PH"},
                {"name": "Dr. Kushal Kumar Pathak (Joint Secretary - Systems & CISO)", "office_phone": "23034967", "room": "34, GF, Samvidhan Sadan"},
                {"name": "Systems Division", "office_phone": "23034325, 23034074", "room": "121, III-F, Samvidhan Sadan"},
                {"name": "IT Complaints / Helpdesk", "office_phone": "23034126, 23034718", "room": "helpdesk.msp@supportgov.in"}
            ]
            q = st.text_input("Search Name / Section:", key="tel_search")
            if q:
                for res in [e for e in tel_directory if q.lower() in e['name'].lower()]:
                    st.success(f"👤 **{res['name']}**\n\n📞 **Phone:** {res['office_phone']} | 🚪 **Room:** {res['room']}")
    if t47:
        with st.expander("47. CGHS Medical Reimbursement Claim Form - MRC(S) (चिकित्सा प्रतिपूर्ति दावा प्रपत्र)"):
            c1, c2 = st.columns(2)
            cghs_name = c1.text_input("1(a) Name of Card Holder & Desig.", key="cghs_name")
            cghs_ben_id = c2.text_input("1(b) CGHS Ben ID No.", key="cghs_ben_id")
            c3, c4 = st.columns(2)
            cghs_emp_code = c3.text_input("1(c) Employee Code No.", key="cghs_emp_code")
            cghs_ward = c4.selectbox("1(d) Ward Entitlement", ["General", "Semi-Pvt.", "Pvt."], key="cghs_ward")
            c5, c6 = st.columns(2)
            cghs_basic = c5.text_input("1(e) Basic Pay (₹)", key="cghs_basic")
            cghs_contact = c6.text_input("1(f) Mobile No. & Email", key="cghs_contact")
            cghs_address = st.text_area("Full Address", height=68, key="cghs_address")
            p1, p2, p3 = st.columns(3)
            cghs_pat_name = p1.text_input("Patient's Name", key="cghs_pat_name")
            cghs_pat_id = p2.text_input("Patient's Ben ID", key="cghs_pat_id")
            cghs_pat_rel = p3.selectbox("Relationship", ["Self", "Spouse", "Son", "Daughter", "Father", "Mother", "Other"], key="cghs_pat_rel")
            cghs_hosp = st.text_input("Hospital Name & Address", key="cghs_hosp")
            a1, a2, a3 = st.columns(3)
            amt_opd = a1.number_input("OPD Claimed", min_value=0.0, value=0.0, step=100.0, key="amt_opd")
            amt_indoor = a2.number_input("Indoor Claimed", min_value=0.0, value=0.0, step=100.0, key="amt_indoor")
            amt_tests = a3.number_input("Tests Claimed", min_value=0.0, value=0.0, step=100.0, key="amt_tests")
            total_claim = amt_opd + amt_indoor + amt_tests
            st.info(f"💰 Total Amount: ₹ {total_claim:,.2f}")
            b1, b2 = st.columns(2)
            b_name = b1.text_input("Bank Name", key="b_name")
            b_acc = b2.text_input("SB Account No.", key="b_acc")
            b3, b4 = st.columns(2)
            b_ifsc = b3.text_input("IFSC Code", key="b_ifsc")
            b_micr = b4.text_input("MICR Code", key="b_micr")
            d1, d2 = st.columns(2)
            dec_date = d1.text_input("Date", value="09.09.2026", key="dec_date")
            dec_place = d2.text_input("Place", value="New Delhi", key="dec_place")
            if st.button("Generate CGHS Form MRC(S) (.docx)", key="btn_cghs_form"):
                doc = docx.Document()
                for s in doc.sections:
                    s.page_width, s.page_height = Mm(210), Mm(297)
                    s.top_margin, s.bottom_margin, s.left_margin, s.right_margin = Inches(0.5), Inches(0.5), Inches(0.6), Inches(0.6)
                p_hdr = doc.add_paragraph()
                p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_hdr.add_run("FORM-MRC (S) (For serving employees)\nCENTRAL GOVERNMENT HEALTH SCHEME\nMEDICAL REIMBURSEMENT CLAIM FORM\n").bold = True
                table = doc.add_table(rows=0, cols=3)
                table.autofit = False
                def add_mrc(n, d, v):
                    r = table.add_row()
                    r.cells[0].width, r.cells[1].width, r.cells[2].width = Inches(0.5), Inches(3.2), Inches(3.3)
                    r.cells[0].paragraphs[0].add_run(n).bold = True
                    r.cells[1].paragraphs[0].add_run(d)
                    r.cells[2].paragraphs[0].add_run(f": {v}")
                add_mrc("1.", "Card Holder Name & Desig.", cghs_name)
                add_mrc("", "CGHS Ben ID / Employee Code", f"{cghs_ben_id} / {cghs_emp_code}")
                add_mrc("", "Ward / Basic Pay", f"{cghs_ward} / ₹ {cghs_basic}")
                add_mrc("", "Address & Mobile", f"{cghs_address} | {cghs_contact}")
                add_mrc("2.", "Patient Name & Ben ID", f"{cghs_pat_name} ({cghs_pat_rel}) / ID: {cghs_pat_id}")
                add_mrc("3.", "Hospital Name & Address", cghs_hosp)
                add_mrc("4.", "Total Claimed Amount", f"₹ {total_claim:,.2f}")
                add_mrc("5.", "Bank Details", f"{b_name} | A/C: {b_acc} | IFSC: {b_ifsc}")
                doc.add_paragraph("\nDECLARATION: I hereby declare that statements made are true to the best of my knowledge.").runs[0].font.size = Pt(9.5)
                doc.add_paragraph(f"\nDate: {dec_date}\tPlace: {dec_place}\t\t\t\tSignature of Card Holder").runs[0].bold = True
                out = io.BytesIO()
                doc.save(out); out.seek(0)
                st.success("Form generated!")
                st.download_button("📥 Download CGHS MRC(S)", data=out, file_name="CGHS_MRC_S.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    if t54:
        with st.expander("54. Application for Medical Test/Treatment Permission"):
            c0, c1, c2, c3 = st.columns([1, 2, 2, 2])
            med_sal = c0.selectbox("Salutation", ["Shri", "Smt.", "Km.", "Ms.", "Mr."], key="med_sal")
            med_emp_name = c1.text_input("1. Name of Employee", value="SAURABH BATRA", key="med_emp_name")
            med_desig = c2.text_input("2. Designation", value="Senior Assistant", key="med_desig")
            med_pay = c3.text_input("3. Basic Pay (₹)", key="med_pay")
            p1, p2 = st.columns(2)
            med_patient = p1.text_input("4. Name of Patient", key="med_patient")
            med_relation = p2.selectbox("5. Relation with Employee", ["Self", "Spouse", "Son", "Daughter", "Father", "Mother", "Other"], key="med_rel")
            t1_col, t2_col = st.columns(2)
            med_recom = t1_col.selectbox("6. Recommended by", ["CMO, CGHS Dispensary", "Specialist, Govt. Hospital", "Authorised Medical Attendant (AMA)"], key="med_recom")
            med_date = t2_col.text_input("7. Date of Prescription slip(s)", key="med_date")
            med_tests = st.text_area("8. Details of Diagnostic tests/Medical Treatment", key="med_tests")
            med_hosp = st.text_input("9. Name of Diagnostic Centre/Hospital", key="med_hosp")
            cgh1, cgh2, cgh3 = st.columns(3)
            med_cghs = cgh1.text_input("10(a) CGHS Card No.", key="med_cghs")
            med_disp = cgh2.text_input("10(b) Dispensary Name & No.", key="med_disp")
            med_ama = cgh3.text_input("11. Name of AMA (If applicable)", key="med_ama")
            o1, o2, o3 = st.columns(3)
            med_branch = o1.text_input("Branch", value="Systems Division", key="med_branch")
            med_tel = o2.text_input("Tel. No.", key="med_tel")
            med_app_date = o3.text_input("Date", key="med_app_date")
            if st.button("Generate Permission Form (.docx)", key="btn_med_perm"):
                doc = docx.Document()
                style = doc.styles['Normal']
                style.font.name = 'Arial'
                style.font.size = Pt(11)
                style.paragraph_format.space_after = Pt(6)
                style.paragraph_format.line_spacing = 1.15

                for s in doc.sections:
                    s.page_width, s.page_height = Mm(210), Mm(297)
                    s.top_margin, s.bottom_margin, s.left_margin, s.right_margin = Inches(0.4), Inches(0.4), Inches(0.6), Inches(0.6)

                p_hdr = doc.add_paragraph()
                p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_hdr.paragraph_format.space_after = Pt(8)
                r_h1 = p_hdr.add_run("APPLICATION FOR GRANT OF PERMISSION FOR DIAGNOSTIC TESTS/MEDICAL TREATMENT\n")
                r_h1.bold = True; r_h1.font.size = Pt(12)
                r_h2 = p_hdr.add_run("[Test/Treatment is to be taken by the official after getting written permission from the Office]")
                r_h2.font.size = Pt(10)

                p1 = doc.add_paragraph()
                p1.add_run("1. Name of the Employee (in capital letters): ")
                p1.add_run(f"{med_sal} {med_emp_name.upper()}").bold = True
                
                doc.add_paragraph(f"2. Designation: {med_desig}")
                doc.add_paragraph(f"3. Basic Pay: {med_pay}")
                doc.add_paragraph(f"4. Name of the Patient: {med_patient}")
                doc.add_paragraph(f"5. Relation with the Emjployee: {med_relation}")

                cmo_check = "(✔)" if med_recom == "CMO, CGHS Dispensary" else "( )"
                spec_check = "(✔)" if med_recom == "Specialist, Govt. Hospital" else "( )"
                ama_check = "(✔)" if med_recom == "Authorised Medical Attendant (AMA)" else "( )"

                doc.add_paragraph("6. Diagnostic Tests/Treatment recommended by: [Please(✔) against the relevant head]")
                doc.add_paragraph(f"    (a) CMO, CGHS Dispensary {cmo_check}      (b) Specialist, Govt. Hospital {spec_check}")
                doc.add_paragraph(f"    (c) Authorised Medical Attendant [for beneficiary not covered under CGHS] {ama_check}")
                doc.add_paragraph(f"7. Date of Prescription slip (s): {med_date}")

                # --- TABLE 8 & 9 ---
                t_8_9 = doc.add_table(rows=2, cols=2)
                t_8_9.style = 'Table Grid'
                t_8_9.autofit = False
                for r in t_8_9.rows: r.cells[0].width, r.cells[1].width = Inches(3.5), Inches(3.5)
                t_8_9.cell(0, 0).paragraphs[0].add_run("8. Details of the Diagnostic tests/Medical Treatment").bold = True
                t_8_9.cell(0, 1).paragraphs[0].add_run("9. Name of Diagnostic Centre/Hospital where\nMedical Diagnostic test/Treatment is to be taken").bold = True
                t_8_9.cell(1, 0).text = med_tests
                t_8_9.cell(1, 1).text = med_hosp

                doc.add_paragraph("10. To be filled by beneficiary covered under CGHS")

                # --- TABLE 10 ---
                t_10 = doc.add_table(rows=2, cols=2)
                t_10.style = 'Table Grid'
                t_10.autofit = False
                for r in t_10.rows: r.cells[0].width, r.cells[1].width = Inches(3.5), Inches(3.5)
                t_10.cell(0, 0).text = "(a)  CGHS Card No."
                t_10.cell(0, 1).text = med_cghs
                t_10.cell(1, 0).text = "(b)  Name & Number of the Dispensary"
                t_10.cell(1, 1).text = med_disp

                # --- POINT 11 ---
                p11 = doc.add_paragraph()
                p11.add_run("11.   To be filled by beneficiary ")
                r_not = p11.add_run("not")
                r_not.underline = True
                p11.add_run(" covered under CGHS")
                
                p11_a = doc.add_paragraph()
                p11_a.add_run("      (a) Name of the Authorised Medical Attendant (AMA): Dr. ")
                if med_ama:
                    p11_a.add_run(f"{med_ama}").bold = True
                else:
                    p11_a.add_run("_________________________")

                doc.add_paragraph("12. I have enclosed the photocopy of the following documents:\n"
                                  "    (a) Prescription slip issued by the doctor.\n"
                                  "        [The name of the doctor, dispensary, date and stamp should be clearly visible and legible]\n"
                                  "    (b) CGHS Card\n"
                                  "    (c) Order of appointment of AMA [for beneficiary not covered under CGHS]")

                doc.add_paragraph("13. I may kindly be granted permission for the above mentioned Test/Treatment.")

                # --- RIGHT-ALIGNED FOOTER BLOCK ---
                footer_table = doc.add_table(rows=4, cols=2)
                footer_table.autofit = False
                for r in footer_table.rows:
                    r.cells[0].width = Inches(4.3)
                    r.cells[1].width = Inches(2.7)

                p0 = footer_table.cell(0, 1).paragraphs[0]
                p0.add_run("Signature: ____________________")

                def add_f_line(idx, lbl, val):
                    p = footer_table.cell(idx, 1).paragraphs[0]
                    p.add_run(lbl)
                    pad_len = max(0, 20 - len(str(val)))
                    padded = f" {val}" + "\u00A0" * pad_len
                    p.add_run(padded).underline = True

                add_f_line(1, "Branch:", med_branch)
                add_f_line(2, "Tel.No.:", med_tel)
                add_f_line(3, "Date:", med_app_date)

                out = io.BytesIO()
                doc.save(out); out.seek(0)
                st.success("Permission Form generated exactly matching A4 format!")
                st.download_button("📥 Download Exact Permission Form", data=out, file_name="Medical_Test_Permission_Exact_Form.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# ==========================================
# CATEGORY 10: ULTIMATE PRODUCTIVITY
# ==========================================
if any([t48, t49, t50, t51, t52, t53]):
    st.markdown("## 🏆 Ultimate Productivity Utilities")
    if t48:
        with st.expander("48. Scanned PDF Whitener & De-Shadow Cleaner (दस्तावेज़ बैकग्राउंड क्लीनर)"):
            wh_file = st.file_uploader("Upload Scanned PDF", type=["pdf"], key="wh_pdf")
            if wh_file and st.button("Clean & Whiten Background", key="btn_wh"):
                try:
                    import fitz
                    with fitz.open(stream=wh_file.read(), filetype="pdf") as doc:
                        out_pdf_bytes = io.BytesIO()
                        merger = PdfWriter()
                        for page in doc:
                            pix = page.get_pixmap(dpi=200)
                            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
                            img = ImageEnhance.Contrast(img).enhance(2.0)
                            img = img.point(lambda p: 255 if p > 150 else p)
                            tpdf = io.BytesIO(); img.save(tpdf, format="PDF"); tpdf.seek(0)
                            merger.add_page(PdfReader(tpdf).pages[0])
                        merger.write(out_pdf_bytes); out_pdf_bytes.seek(0)
                        st.success("Background cleaned successfully!")
                        st.download_button("Download Cleaned PDF", data=out_pdf_bytes, file_name="Cleaned_White.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t49:
        with st.expander("49. Smart Auto-PII Redactor (आधार / पैन / मोबाइल ऑटो-ब्लैकआउट)"):
            pii_pdf = st.file_uploader("Upload Digital PDF", type=["pdf"], key="pii_pdf")
            c1, c2, c3, c4 = st.columns(4)
            h_aadhaar = c1.checkbox("Aadhaar Numbers", value=True)
            h_pan = c2.checkbox("PAN Numbers", value=True)
            h_mobile = c3.checkbox("Mobile Numbers", value=True)
            h_email = c4.checkbox("Email IDs", value=True)
            if pii_pdf and st.button("Auto-Redact Sensitive Info", key="btn_pii"):
                try:
                    import fitz
                    with fitz.open(stream=pii_pdf.read(), filetype="pdf") as doc:
                        patterns = []
                        if h_aadhaar: patterns.append(r"\b\d{4}[\s\-]?\d{4}[\s\-]?\d{4}\b")
                        if h_pan: patterns.append(r"\b[A-Z]{5}\d{4}[A-Z]\b")
                        if h_mobile: patterns.append(r"\b(?:\+91|91)?[\s\-]?\d{10}\b")
                        if h_email: patterns.append(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b")
                        count = 0
                        for page in doc:
                            t = page.get_text("text")
                            for pat in patterns:
                                for match in re.findall(pat, t):
                                    for r in page.search_for(match):
                                        page.add_redact_annot(r, fill=(0, 0, 0)); count += 1
                            page.apply_redactions()
                        out_pdf = io.BytesIO()
                        doc.save(out_pdf); out_pdf.seek(0)
                        st.success(f"Success! Blocked {count} items.")
                        st.download_button("Download Redacted PDF", data=out_pdf, file_name="PII_Redacted.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t50:
        with st.expander("50. Manual Duplex Printing Assistant (ऑड-ईवन प्रिंट हेल्पर)"):
            duplex_pdf = st.file_uploader("Upload PDF", type=["pdf"], key="duplex_pdf")
            if duplex_pdf and st.button("Generate Duplex Print Files", key="btn_duplex"):
                try:
                    reader = PdfReader(duplex_pdf)
                    w_odd, w_even = PdfWriter(), PdfWriter()
                    even_pages = []
                    for i in range(len(reader.pages)):
                        if i % 2 == 0: w_odd.add_page(reader.pages[i])
                        else: even_pages.append(reader.pages[i])
                    for p in reversed(even_pages): w_even.add_page(p)
                    odd_out, even_out = io.BytesIO(), io.BytesIO()
                    w_odd.write(odd_out); w_even.write(even_out)
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        zip_file.writestr("1_Print_First_Odd_Pages.pdf", odd_out.getvalue())
                        zip_file.writestr("2_Print_Second_Even_Pages_Reversed.pdf", even_out.getvalue())
                    st.success("Files prepared successfully!")
                    st.download_button("Download Duplex ZIP", data=zip_buffer.getvalue(), file_name="Manual_Duplex.zip", mime="application/zip")
                except Exception as e: st.error(f"Error: {e}")
    if t51:
        with st.expander("51. PDF Color vs B&W Page Audit & Splitter (प्रिंटर बजट सेवर)"):
            audit_pdf = st.file_uploader("Upload PDF to Audit", type=["pdf"], key="audit_pdf")
            if audit_pdf and st.button("Audit and Split Pages", key="btn_audit"):
                try:
                    import fitz
                    with fitz.open(stream=audit_pdf.read(), filetype="pdf") as doc, fitz.open() as bw_doc, fitz.open() as col_doc:
                        bw_c, col_c = 0, 0
                        for i in range(len(doc)):
                            pix = doc[i].get_pixmap(colorspace=fitz.csRGB, dpi=36)
                            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                            gray = img.convert('L').convert('RGB')
                            diff = ImageChops.difference(img, gray)
                            stat = ImageStat.Stat(diff)
                            if sum(stat.mean) > 1.5:
                                col_doc.insert_pdf(doc, from_page=i, to_page=i); col_c += 1
                            else:
                                bw_doc.insert_pdf(doc, from_page=i, to_page=i); bw_c += 1
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            if bw_c > 0:
                                bwb = io.BytesIO(); bw_doc.save(bwb); zip_file.writestr(f"BW_Pages_({bw_c}).pdf", bwb.getvalue())
                            if col_c > 0:
                                cb = io.BytesIO(); col_doc.save(cb); zip_file.writestr(f"Color_Pages_({col_c}).pdf", cb.getvalue())
                        st.success(f"Audit Complete! {bw_c} B&W, {col_c} Color.")
                        st.download_button("Download Split Audit ZIP", data=zip_buffer.getvalue(), file_name="Audit_Split.zip", mime="application/zip")
                except Exception as e: st.error(f"Error: {e}")
    if t52:
        with st.expander("52. PDF Embedded Portfolio / Attachment Packer (डिजिटल मिसल टूल)"):
            main_pdf = st.file_uploader("1. Upload Main PDF", type=["pdf"], key="main_pdf")
            attach_files = st.file_uploader("2. Upload Attachments", accept_multiple_files=True, key="attach_files")
            if main_pdf and attach_files and st.button("Pack into Portfolio PDF", key="btn_pack"):
                try:
                    import fitz
                    with fitz.open(stream=main_pdf.read(), filetype="pdf") as doc:
                        for f in attach_files: doc.embfile_add(f.name, f.read(), filename=f.name)
                        out_pdf = io.BytesIO()
                        doc.save(out_pdf); out_pdf.seek(0)
                        st.success(f"Embedded {len(attach_files)} files!")
                        st.download_button("Download Portfolio PDF", data=out_pdf, file_name="Portfolio.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")
    if t53:
        with st.expander("53. Smart Booklet Imposition (Center-Staple Maker)"):
            booklet_pdf = st.file_uploader("Upload PDF Document", type=["pdf"], key="booklet_pdf")
            sheet_size = st.selectbox("Printer Paper Size", ["A4 (A5 folded booklet)", "A3 (A4 folded booklet)", "Letter (Half-Letter folded)"], key="booklet_size")
            if booklet_pdf and st.button("Generate Booklet", key="btn_booklet"):
                try:
                    import fitz
                    with fitz.open(stream=booklet_pdf.read(), filetype="pdf") as doc, fitz.open() as out_doc:
                        rem = len(doc) % 4
                        if rem != 0:
                            for _ in range(4 - rem): doc.new_page(width=doc[-1].rect.width, height=doc[-1].rect.height)
                        tp = len(doc)
                        sw, sh = fitz.paper_size("a4-l") if "A4" in sheet_size else (fitz.paper_size("a3-l") if "A3" in sheet_size else fitz.paper_size("letter-l"))
                        hw = sw / 2
                        for k in range(tp // 4):
                            fp = out_doc.new_page(width=sw, height=sh)
                            fp.show_pdf_page(fitz.Rect(0, 0, hw, sh), doc, tp - 1 - 2*k)
                            fp.show_pdf_page(fitz.Rect(hw, 0, sw, sh), doc, 2*k)
                            bp = out_doc.new_page(width=sw, height=sh)
                            bp.show_pdf_page(fitz.Rect(0, 0, hw, sh), doc, 2*k + 1)
                            bp.show_pdf_page(fitz.Rect(hw, 0, sw, sh), doc, tp - 2 - 2*k)
                        out_pdf = io.BytesIO()
                        out_doc.save(out_pdf); out_pdf.seek(0)
                        st.success("Booklet generated!")
                        st.download_button("Download Print-Ready Booklet", data=out_pdf, file_name="Booklet.pdf", mime="application/pdf")
                except Exception as e: st.error(f"Error: {e}")

# --- FOOTER ---
st.markdown("<br><br><br><br>", unsafe_allow_html=True)
st.markdown("""
    <div class="footer">
        For secure and seamless document conversion, use दस्तावेज़ सेतु—because confidentiality should never be compromised.
    </div>
""", unsafe_allow_html=True)
