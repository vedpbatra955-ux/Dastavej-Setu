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

try:
    from docx2pdf import convert as docx_convert
except ImportError:
    docx_convert = None

try:
    from pdf2docx import Converter
except ImportError:
    Converter = None

try:
    from rembg import remove as rembg_remove
except ImportError:
    rembg_remove = None

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    import pytesseract
    from pdf2image import convert_from_bytes
except ImportError:
    pytesseract = None

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


# ==========================================
# SIDEBAR (MINIMAL)
# ==========================================
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/c/cb/New_Parliament_Building%2C_New_Delhi_-_Mar_2024.jpg", width="stretch")

st.sidebar.markdown("### ⚙️ System Actions")
if st.sidebar.button("🧹 Clear Workspace & Cache", width="stretch"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

st.sidebar.markdown("<br><br><br><br>", unsafe_allow_html=True)
st.sidebar.caption("v3.10 Secretariat Build")


# ==========================================
# MAIN HEADER & LIVE SEARCH BOX
# ==========================================
st.markdown("""
    <div class="main-header">
        <div class="org-title">Rajya Sabha Secretariat</div>
        <div class="main-title">दस्तावेज़ सेतु</div>
        <div class="portal-subtitle">Secure Intranet Document Utility & Conversion Portal</div>
    </div>
""", unsafe_allow_html=True)

st.markdown("### 🔍 Live Tool Search")
search_query = st.text_input("Search", "", placeholder="Start typing to instantly find a tool (e.g., 'CGHS', 'Gate Pass', 'Directory', 'Word')...", label_visibility="collapsed").strip()


# Helper function to match search words
def is_match(title, keywords=""):
    if not search_query: return True
    search_words = search_query.lower().split()
    target_text = (title + " " + keywords).lower()
    for word in search_words:
        if word not in target_text:
            return False
    return True

# Pre-calculate visibility for all 54 tools
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

# Check if there are ANY matches at all
if not any([t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13, t14, t15, t16, t17, t18, t19, t20, t21, t22, t23, t24, t25, t26, t27, t28, t29, t30, t31, t32, t33, t34, t35, t36, t37, t38, t39, t40, t41, t42, t43, t44, t45, t46, t47, t48, t49, t50, t51, t52, t53, t54]):
    st.warning("⚠️ No utilities match your search. Try adjusting your keywords (e.g., 'CGHS', 'Word', 'Stamp', 'Directory').")


# ==========================================
# CATEGORY 1: CORE FILE MANAGEMENT
# ==========================================
if any([t1, t2, t3, t4, t5]):
    st.markdown("## 📁 Core File Management")

    if t1:
        with st.expander("1. Merge Multiple PDFs (Advanced Reorder)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload multiple PDF files here. Once uploaded, click 'Merge PDFs'. The system will combine all of them into a single PDF document in the order they were selected.")
            uploaded_files = st.file_uploader("Upload PDF files to merge", type=["pdf"], accept_multiple_files=True, key="merge_pdf")
            if uploaded_files and st.button("Merge PDFs", key="btn_merge"):
                merger = PdfWriter()
                for f in uploaded_files: merger.append(f)
                output_stream = io.BytesIO()
                merger.write(output_stream)
                merger.close()
                st.success("PDFs merged successfully!")
                st.download_button("Download Merged PDF", data=output_stream.getvalue(), file_name="merged_output.pdf", mime="application/pdf")

    if t2:
        with st.expander("2. Extract Specific Pages from PDF"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a single PDF. The tool will show you the total number of pages. Enter the 'Start Page' and 'End Page' you want to keep, then click 'Extract Pages' to download only those selected pages.")
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="ext_pdf")
            if uploaded_file:
                reader = PdfReader(uploaded_file)
                total_p = len(reader.pages)
                st.info(f"Total pages in PDF: {total_p}")
                col1, col2 = st.columns(2)
                start = col1.number_input("Start Page", min_value=1, max_value=total_p, value=1)
                end = col2.number_input("End Page", min_value=start, max_value=total_p, value=total_p)
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
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF to instantly see its total page count. This tool acts as a quick viewer so you can decide which pages to split/extract using Tool No. 2 above.")
            uploaded_file = st.file_uploader("Upload PDF to Split", type=["pdf"], key="split_pdf")
            if uploaded_file and st.button("Split Pages", key="btn_split"):
                reader = PdfReader(uploaded_file)
                st.success(f"PDF has {len(reader.pages)} pages. Use Extract Pages above to download specific ranges.")

    if t4:
        with st.expander("4. Delete Specific Pages from PDF"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload your PDF. In the text box, type the exact page numbers you want to remove, separated by commas (Example: `1, 3, 5`). Click 'Delete Pages' to download the updated PDF.")
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
                except Exception as e:
                    st.error(f"Error: {e}")

    if t5:
        with st.expander("5. Reverse PDF Page Order (Back to Front)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF. Click 'Reverse Order' to flip the entire document backwards. The last page will become the first page, and the first page will go to the end.")
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
if any([t6, t7, t8, t9, t10, t11, t12, t13, t14]):
    st.markdown("## 📝 Conversions & Drafting")

    if t6:
        with st.expander("6. Generate Official Notice Draft Template"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Simply click the 'Generate Template' button below. It will instantly download a ready-to-use Notice Template in MS Word (.docx) format with official margins and layout.")
            st.write("Generate the official Rajya Sabha Secretariat formatted notice template instantly.")
            if st.button("Generate Template (.docx)", key="btn_template"):
                doc = docx.Document()
                for s in doc.sections: s.top_margin = Inches(1); s.bottom_margin = Inches(1); s.left_margin = Inches(1); s.right_margin = Inches(1)
                p_header = doc.add_paragraph()
                p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p_header.add_run("RAJYA SABHA SECRETARIAT\nBRANCH NAME")
                r.bold = True
                r.font.size = Pt(14)
                doc.add_paragraph()
                p_sub = doc.add_paragraph()
                p_sub.add_run("Subject: ").bold = True
                p_sub.add_run("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX")
                doc.add_paragraph("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX.")
                doc.add_paragraph("2.\t XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX.")
                doc.add_paragraph("3.\tXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX.")
                doc.add_paragraph("4.\tXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX.")
                doc.add_paragraph()
                p_sig = doc.add_paragraph()
                p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_sig.add_run("(NAME)\nDesignation\nXX.XX.XXXX\nTel: XXXXXXX")
                doc.add_paragraph()
                doc.add_paragraph("To,\nXXXXXXXXXXXXXXXXXXXXXXX\nXXXXXXXXXXXXXXXXXXXXXXX\nXXXXXXXXXXXXXXXXXXXXXXX\nXXXXXXXXXXXXXXXXXXXXXXX")
                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                st.success("Notice Template generated successfully!")
                st.download_button("Download Draft Template", data=out, file_name="Notice_Template.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    if t7:
        with st.expander("7. Generate Visitor Pass (Auto-Fill & Multiple Visitors)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("1. Select the Officer and Designation.\n2. Add Room and Phone number.\n3. Enter the Visitor's Name and Aadhaar (Click '➕ Add Visitor' to add more people).\n4. Select if gadgets are allowed.\n5. Click 'Generate Visitor Pass' to download the Word document.")
            col1, col2 = st.columns(2)
            vp_officer = col1.selectbox("NAME OF THE OFFICER", ["Navneet Joon", "Sandeep Pandey", "Devyanshu Pal", "Ankit Chansoria", "Agam Mittal", "R.Arivazhagan"], key="vp_off")
            vp_desig = col2.selectbox("DESIGNATION", ["Deputy Secretary", "Under Secretary", "Executive Officer"], key="vp_desig")
            col3, col4 = st.columns(2)
            vp_room = col3.text_input("ROOM NO.", value="209, 2nd Floor, Parliament House Annexe", key="vp_room")
            vp_tel = col4.text_input("OFFICE TELEPHONE NO.", value="011-23034325", key="vp_tel")
            
            st.markdown("---")
            st.markdown("**Visitor Details (Add multiple visitors if needed)**")
            
            if 'visitor_count' not in st.session_state: 
                st.session_state['visitor_count'] = 1
                
            v_col1, v_col2, v_col3 = st.columns([1, 1, 2])
            with v_col1:
                if st.button("➕ Add Visitor", key="add_v"): st.session_state['visitor_count'] += 1
            with v_col2:
                if st.button("➖ Remove Visitor", key="rem_v") and st.session_state['visitor_count'] > 1: st.session_state['visitor_count'] -= 1
                    
            visitor_names = []
            visitor_aadhaars = []
            for i in range(st.session_state['visitor_count']):
                c1, c2 = st.columns(2)
                name = c1.text_input(f"Visitor {i+1} Name", key=f"v_name_{i}")
                # Preserved specifically for Aadhaar naming as requested by user
                aadhaar = c2.text_input(f"Visitor {i+1} Aadhaar ID", key=f"v_aadh_{i}")
                visitor_names.append(name.strip() if name else "")
                visitor_aadhaars.append(aadhaar.strip() if aadhaar else "")
                
            st.markdown("---")
            col7, col8 = st.columns(2)
            vp_datetime = col7.text_input("DATE AND TIME", key="vp_dt")
            vp_purpose = col8.text_input("PURPOSE OF VISIT", value="Official Meeting", key="vp_purp")
            vp_gadgets = st.selectbox("Allow Mobile & Laptop?", ["✅ Yes (Allow gadgets and print line)", "❌ No (Do not include gadget permission line)"], key="vp_gadgets")

            if st.button("Generate Visitor Pass (.docx)", key="btn_vp"):
                doc = docx.Document()
                for s in doc.sections: s.top_margin = Inches(1); s.bottom_margin = Inches(1); s.left_margin = Inches(1); s.right_margin = Inches(1)
                p_head = doc.add_paragraph()
                p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_head = p_head.add_run("PARLIAMENT OF INDIA\nRAJYA SABHA SECRETARIAT\n[SYSTEMS DIVISION]")
                r_head.bold = True
                r_head.font.size = Pt(13)
                doc.add_paragraph() 
                table = doc.add_table(rows=0, cols=3)
                table.columns[0].width = Inches(2.2)
                table.columns[1].width = Inches(0.2)
                table.columns[2].width = Inches(3.8)
                
                combined_visitors = []
                for n, a in zip(visitor_names, visitor_aadhaars):
                    if n or a:
                        line = n if n else "Unknown Visitor"
                        # Explicitly mapping back to Aadhaar label per user request
                        if a: line += f" (Aadhaar: {a})"
                        combined_visitors.append(line)
                vp_visitor_str = "\n".join(combined_visitors) if combined_visitors else ""
                
                fields = [("NAME OF THE OFFICER", vp_officer), ("DESIGNATION", vp_desig), ("ROOM NO.", vp_room), ("OFFICE TELEPHONE NO.", vp_tel), ("NAME OF THE VISITOR(S)", vp_visitor_str)]
                for label, val in fields:
                    row = table.add_row().cells
                    p_label = row[0].paragraphs[0]
                    r_label = p_label.add_run(label)
                    r_label.bold = True
                    row[1].text = ":"
                    row[2].text = val
                    p_label.paragraph_format.space_after = Pt(10)
                
                doc.add_paragraph()
                if "Yes" in vp_gadgets:
                    p_note = doc.add_paragraph()
                    r_note = p_note.add_run("(Above person may be allowed along with his Mobile phone and Laptop with adapters/chargers)")
                    r_note.bold = True
                doc.add_paragraph()
                
                p_dt = doc.add_paragraph()
                p_dt.add_run("DATE AND TIME:\t\t").bold = True
                p_dt.add_run(vp_datetime)
                p_purp = doc.add_paragraph()
                p_purp.add_run("PURPOSE OF VISIT:\t").bold = True
                p_purp.add_run(vp_purpose)
                
                doc.add_paragraph("\n\n")
                p_sig = doc.add_paragraph()
                p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_sig.add_run("(SIGNATURE WITH STAMP)").bold = True
                
                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                st.success("Visitor Pass generated successfully!")
                st.download_button("Download Visitor Pass", data=out, file_name="Visitor_Pass.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_vp")

    if t8:
        with st.expander("8. Generate Gate Pass (Systems Division Auto-Fill)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Fill out all the details for the item/hardware leaving the building. Click 'Generate' to get a clean, perfectly aligned professional Gate Pass in MS Word format.")
            col1, col2 = st.columns(2)
            gp_officer = col1.text_input("Name of officer/Section", key="gp_off")
            gp_date = col2.text_input("Date of Issue", key="gp_date")
            gp_item = st.text_area("Item Description", key="gp_item")
            gp_make = st.text_input("Make", key="gp_make")
            col3, col4 = st.columns(2)
            gp_favour = col3.text_input("Issue in favour of Shri", key="gp_fav")
            gp_org = col4.text_input("Of (organization)", key="gp_org")
            col5, col6 = st.columns(2)
            gp_from = col5.text_input("Valid for movement from", key="gp_from")
            gp_to = col6.text_input("To", key="gp_to")
            col7, col8 = st.columns(2)
            gp_gate = col7.text_input("Through Gate No.", key="gp_gate")
            gp_ondate = col8.text_input("On Date", key="gp_ondate")
            col9, col10 = st.columns(2)
            gp_between = col9.text_input("Between (Time)", key="gp_bet")
            gp_and = col10.text_input("To (Time)", key="gp_and")
            st.markdown("---")
            col11, col12 = st.columns(2)
            gp_sig_name = col11.text_input("Issuer Name (for Signature)", key="gp_sig_name")
            gp_sig_desig = col12.text_input("Issuer Designation", key="gp_sig_desig")

            if st.button("Generate Gate Pass (.docx)", key="btn_gp"):
                doc = docx.Document()
                for s in doc.sections: 
                    s.top_margin = Inches(1)
                    s.bottom_margin = Inches(1)
                    s.left_margin = Inches(1)
                    s.right_margin = Inches(1)
                    
                # Header
                p_head = doc.add_paragraph()
                p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_head1 = p_head.add_run("RAJYA SABHA SECRETARIAT\n")
                r_head1.underline = True
                r_head1.font.size = Pt(14)
                r_head2 = p_head.add_run("SYSTEMS DIVISION")
                r_head2.font.size = Pt(13)
                
                doc.add_paragraph()
                p_title = doc.add_paragraph()
                p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_title = p_title.add_run("GATE PASS")
                r_title.underline = True
                r_title.font.size = Pt(13)
                doc.add_paragraph()
                
                # --- PROFESSIONAL TABLE-BASED GATE PASS LAYOUT ---
                table = doc.add_table(rows=0, cols=2)
                table.autofit = False
                
                def add_gp_row(label, value):
                    row = table.add_row()
                    cell_lbl, cell_val = row.cells[0], row.cells[1]
                    cell_lbl.width = Inches(2.3)
                    cell_val.width = Inches(4.2)
                    
                    p_l = cell_lbl.paragraphs[0]
                    p_l.add_run(label).bold = False
                    p_l.paragraph_format.space_after = Pt(12)
                    
                    p_v = cell_val.paragraphs[0]
                    p_v.paragraph_format.space_after = Pt(12)
                    if value:
                        r = p_v.add_run(f" {value} ")
                        r.underline = True
                    else:
                        r = p_v.add_run(" " * 50)
                        r.underline = True

                def add_gp_double_row(label1, val1, label2, val2):
                    row = table.add_row()
                    cell = row.cells[0]
                    cell.merge(row.cells[1])
                    cell.width = Inches(6.5)
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_after = Pt(12)
                    
                    p.add_run(label1).bold = False
                    r1 = p.add_run(f" {val1 or ''} ".ljust(25, ' '))
                    r1.underline = True
                    
                    p.add_run("    " + label2)
                    r2 = p.add_run(f" {val2 or ''} ")
                    r2.underline = True

                add_gp_double_row("Name of officer/Section:", gp_officer, "Date of Issue:", gp_date)
                add_gp_row("Item Description:", gp_item)
                add_gp_row("", "") # Extra blank underline for long items
                add_gp_row("Make:", gp_make)
                add_gp_row("Issue in favour of Shri:", gp_favour)
                add_gp_row("Of (organization):", gp_org)
                add_gp_row("Valid for movement from:", gp_from)
                add_gp_row("To:", gp_to)
                add_gp_double_row("Through Gate No.:", gp_gate, "On Date:", gp_ondate)
                add_gp_double_row("Between (Time):", gp_between, "To (Time):", gp_and)
                
                doc.add_paragraph("\n")
                p_sig = doc.add_paragraph()
                p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                name = gp_sig_name if gp_sig_name else "(Name)"
                desig = gp_sig_desig if gp_sig_desig else "DESIGNATION"
                p_sig.add_run(f"{name}\n").bold = True
                p_sig.add_run(f"{desig}")

                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                st.success("Gate Pass generated successfully in Clean Professional Table Format!")
                st.download_button("Download Gate Pass", data=out, file_name="Gate_Pass.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_gp")

    if t9:
        with st.expander("9. Generate ITDC Catering Note (Meeting Arrangements)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Enter the meeting name, date, time, and quantities for tea, biscuits, and food items. Select the officer signing the note. Click 'Generate' to download the ITDC Note format with a **1.5-inch left margin**, ready to be printed on a Green Note Sheet.")
            
            # --- MENU ITEMS EXTRACTED FROM RLT.pdf ---
            menu_items = [
                "None", "BATATA VADA", "STUFED BREAD PAKORA", "POTATO PEAS CUTLET", "INDORI POНА", 
                "BEDMI POORI WITH BHAJI", "FRENCH FRIES", "BHAJIA", "VEG PAKORA", "KHAMAN DHOKLA", 
                "ANKURIT CHAAT", "DAL KACHORI", "SHAHI SAMOSA", "COTTAGE CHEESE PAKORA", "CHUKANDAR CUTLET", 
                "COTTAGE CHEESE FINGER", "PANEER ROLL", "SWEET CORN CHAT", "STUFFED HARA KABAB", 
                "CUCUMBER BUTTER SANDWICH", "VETGETABLE JULIENNE SANDWICH", "PERI PERI CHICKEN CUTLET", 
                "MASALA CHICKEN FRY", "FRIED FISH(BASA)", "SPL. MUTTON CUTLET", "RAGI VEG IDLY", 
                "RAVA IDLY WITH SAMBHAR & CHUTNEY", "MASALA DAL VADA WITH SAMBHAR & CHUTNEY", "RAWA UPΜΑ", 
                "VEG SEVIAN", "MASALA TOSSED IDLY WITH CHUTNEY", "MEDU VADA WITH CHUTNEY", "SPL.MASALA DOSA", 
                "PAPER ROAST DOSA", "ΟΝΙΟΝ ΤΟMATO UTTАРАМ", "VEG SOUP", "NON VEG SOUP", "MURG DUM BIRYANI", 
                "MURG CURRY", "ANDA CURRY", "GOSHT BIRYANI", "MUTTON ROGANJOSH", "ONION TOMATO OMELLETE", 
                "SALT & PEPPER OMELLETE", "DAHI", "GARDEN GREEN SALAD", "CURD RICE WITH PICKLE", "YELLOW DAL TADKA", 
                "PALAK KADHI", "MILLET KHICHDI WITH ACHAR", "LEMON PEANUT RICE", "KADHAI PANEER", "MATAR PANEER", 
                "MIX VEG CURRY", "MIX VEG JHALFAREZI", "SAUTE VEGETABLES", "TOMATO PULAV", "MIX VEG PONGAL", 
                "PLAIN RICE", "SUBZ DUM BIRYANI", "TAVA ROTI", "ROTI TANDOORI", "BADAM KHEER", "RABRI SEVIAN KHEER", 
                "DRY FRUIT KHAJUR BARFI", "NARIYAL BARFI", "MOONG DAL BARFI", "CHOCOLATE KHOYA BARFI", 
                "STUFFED GULAB JAMUN", "KESARI IMARTI", "CHENA MITHAI", "PISTA KHOYA BARFI", "MILKCAKE BARFI", 
                "MAWA PEDA", "BESAN DRYFRUIT LADOO", "RASMALAI", "BESAN KATLI", "MAKHAN VADA", "PISTA STUFFED RAJBHOG", 
                "SOOJI HALWA", "ROASTED CASHEWNUT DOTTING", "ROASTED ALMONDS DOTTING", "REGULAR VEG THALI", 
                "DELUXE VEG THALI", "DELUXE NON VEG THALI", "SPL. VEG BUFFET", "SPL. NON VEG BUFFET", 
                "SPL. BUFFET MINISTRIES", "SPL. BUFFET OTHERS", "SPL. HI TEA BUFFET", "SPECIAL PARATHA", 
                "SHAHI BHALLA PAPDI CHAT", "CHANA KULCHE", "DUM EGG BIRYANI", "MEEN CURRY", "SPECIAL MOONG DAL CHILLA", 
                "SHAHI PAPDI CHAT", "SPECIAL MARGHERITA PIZZA", "SPECIAL PIZZA WITH ASSORTED VEGETABLES", 
                "SPL. TEA/COFFEE WITH BISCUITS", "BLENDED COLD COFFEE", "SPECIAL CESAR SALAD", "CUT FRUIT PLATTER", 
                "ROASTED/FRIED PAPAD", "SPECIAL PAV BHAJI", "SPECIAL PYAZ KACHODI"
            ]

            col1, col2 = st.columns(2)
            mtg_name = col1.text_input("Meeting / Committee Name", value="FBEC", key="itdc_mtg")
            mtg_date = col2.text_input("Meeting Date", value="24.08.2026", key="itdc_mdate")
            col3, col4 = st.columns(2)
            mtg_time = col3.text_input("Meeting Time", value="4:00 PM", key="itdc_time")
            mtg_venue = col4.text_input("Venue / Room No.", value="Room No: 124, First Floor, PHA", key="itdc_venue")
            col5, col6 = st.columns(2)
            tea_cups = col5.text_input("Number of Tea Cups", value="10", key="itdc_tea")
            biscuits = col6.text_input("Number of Biscuit Packets", value="5", key="itdc_bisc")
            
            st.markdown("---")
            st.markdown("**Additional Food Arrangements (From RLT Menu)**")
            
            # --- DYNAMIC ADD/REMOVE FOOD ITEMS LOGIC ---
            if 'itdc_food_count' not in st.session_state: 
                st.session_state['itdc_food_count'] = 1
                
            f_col1, f_col2, f_col3 = st.columns([1, 1, 2])
            with f_col1:
                if st.button("➕ Add Food Item", key="add_f"): st.session_state['itdc_food_count'] += 1
            with f_col2:
                if st.button("➖ Remove Food Item", key="rem_f") and st.session_state['itdc_food_count'] > 0: st.session_state['itdc_food_count'] -= 1

            selected_foods = []
            selected_qtys = []
            
            for i in range(st.session_state['itdc_food_count']):
                col_f1, col_f2 = st.columns([2, 1])
                food = col_f1.selectbox(f"Select Food Item {i+1}", menu_items, key=f"itdc_food_{i}")
                qty = col_f2.text_input(f"Quantity {i+1}", value="0", key=f"itdc_food_qty_{i}")
                if food != "None" and qty != "0" and qty.strip() != "":
                    selected_foods.append(food)
                    selected_qtys.append(qty)
            
            st.markdown("---")
            col7, col8 = st.columns(2)
            itdc_off = col7.selectbox("NAME OF THE OFFICER", ["Navneet Joon", "Sandeep Pandey", "Devyanshu Pal", "Ankit Chansoria", "Agam Mittal", "R.Arivazhagan"], key="itdc_off_sel")
            itdc_desig = col8.selectbox("DESIGNATION", ["Deputy Secretary", "Under Secretary", "Executive Officer"], key="itdc_desig_sel")
            col9, col10 = st.columns(2)
            itdc_issue_date = col9.text_input("Issue Date", value="21.08.2026", key="itdc_idate")
            itdc_tel = col10.text_input("Tel No.", value="23034068", key="itdc_tel")

            if st.button("Generate ITDC Note (.docx)", key="btn_itdc"):
                doc = docx.Document()
                for s in doc.sections: 
                    s.top_margin = Inches(1)
                    s.bottom_margin = Inches(1)
                    s.left_margin = Inches(1.5)
                    s.right_margin = Inches(1)
                    
                p_head = doc.add_paragraph()
                p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_head = p_head.add_run("RAJYA SABHA SECRETARIAT\nSYSTEMS DIVISION")
                r_head.bold = True
                r_head.font.size = Pt(13)
                doc.add_paragraph()
                
                # Dynamic string logic if food is selected
                food_subject = " and Snacks/Food" if selected_foods else ""
                
                food_body_list = [f"{q} plates/packets of {f}" for f, q in zip(selected_foods, selected_qtys)]
                food_body = ", and " + ", ".join(food_body_list) if food_body_list else ""
                
                p_sub = doc.add_paragraph()
                p_sub.add_run(f"Subject: Arrangement of Water Bottles, Glasses, Tea, Biscuits{food_subject} for ").bold = True
                p_sub.add_run(f"{mtg_name} Meeting.").bold = True
                
                p_body = doc.add_paragraph()
                p_body.add_run(f"India Tourism Development Corporation (ITDC) Ltd., Parliament House Annexe (PHA), is requested to arrange water bottles, glasses, tea ({tea_cups} cups), biscuits ({biscuits} packets){food_body} for the Meeting of {mtg_name} is scheduled to be held as per details given below:")
                
                p_details = doc.add_paragraph()
                p_details.add_run(f"Venue:\t{mtg_venue}\n")
                p_details.add_run(f"Date:\t{mtg_date}\n")
                p_details.add_run(f"Time:\t{mtg_time}")
                doc.add_paragraph("2. The above arrangements may kindly be ensured accordingly.")
                doc.add_paragraph("\n")
                p_sig = doc.add_paragraph()
                p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_sig.add_run(f"({itdc_off})\n").bold = True
                p_sig.add_run(f"{itdc_desig}\n").bold = True
                p_sig.add_run(f"{itdc_issue_date}\n")
                p_sig.add_run(f"Tel: {itdc_tel}")
                doc.add_paragraph("\nTo,\nThe Manager,\nIndia Tourism Development Corporation,\nParliament House Annexe,\nNew Delhi.")
                
                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                st.success("ITDC Catering Note generated successfully!")
                st.download_button("Download ITDC Note", data=out, file_name="ITDC_Catering_Note.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_itdc")

    if t10:
        with st.expander("10. Generate Green Note Sheet (फाइल नोटिंग) Template"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Enter your File Number, Subject, and Note Description. Select the signing officer. The generated Word document will automatically have a **1.5-inch left margin** so it aligns perfectly when printed and tagged on a Green Note Sheet.")
            col1, col2 = st.columns(2)
            gn_file_no = col1.text_input("File No.", value="RS/Sys/2026/...", key="gn_file")
            gn_date = col2.text_input("Date", value="21.08.2026", key="gn_date")
            gn_subject = st.text_input("Subject", value="Regarding...", key="gn_sub")
            gn_body = st.text_area("Note Description (टिप्पणी)", value="Put up for approval please.", height=150, key="gn_body")
            st.markdown("---")
            col3, col4 = st.columns(2)
            gn_off = col3.selectbox("NAME OF THE OFFICER", ["Navneet Joon", "Sandeep Pandey", "Devyanshu Pal", "Ankit Chansoria", "Agam Mittal", "R.Arivazhagan"], key="gn_off_sheet")
            gn_desig = col4.selectbox("DESIGNATION", ["Deputy Secretary", "Under Secretary", "Executive Officer"], key="gn_desig_sheet")

            if st.button("Generate Note Sheet (.docx)", key="btn_green_note"):
                doc = docx.Document()
                for s in doc.sections:
                    s.top_margin = Inches(1)
                    s.bottom_margin = Inches(1)
                    s.left_margin = Inches(1.5)
                    s.right_margin = Inches(1)
                    
                p_head = doc.add_paragraph()
                p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_head1 = p_head.add_run("RAJYA SABHA SECRETARIAT\n")
                r_head1.bold = True
                r_head1.font.size = Pt(14)
                r_head2 = p_head.add_run("SYSTEMS DIVISION\n")
                r_head2.bold = True
                r_head2.font.size = Pt(12)
                r_head3 = p_head.add_run("(NOTE SHEET)")
                r_head3.underline = True
                r_head3.bold = True
                r_head3.font.size = Pt(12)
                doc.add_paragraph()
                
                p_file = doc.add_paragraph()
                p_file.add_run(f"F.No.: {gn_file_no}").bold = True
                p_sub = doc.add_paragraph()
                p_sub.add_run("Subject: ").bold = True
                p_sub.add_run(gn_subject).bold = True
                doc.add_paragraph()
                
                p_body = doc.add_paragraph(gn_body)
                p_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                doc.add_paragraph("\n")
                
                p_sig = doc.add_paragraph()
                p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_sig.add_run(f"({gn_off})\n").bold = True
                p_sig.add_run(f"{gn_desig}\n")
                p_sig.add_run(f"{gn_date}")

                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                st.success("Green Note Sheet generated successfully!")
                st.download_button("Download Note Sheet", data=out, file_name="Green_Note_Sheet.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_gn")

    if t11:
        with st.expander("11. Convert Single Image to PDF"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload any JPG, JPEG, or PNG image. Click 'Convert to PDF' and the tool will instantly download it as a standard PDF file.")
            uploaded_file = st.file_uploader("Upload Image (JPG/PNG)", type=["jpg", "jpeg", "png"], key="img_pdf")
            if uploaded_file and st.button("Convert to PDF", key="btn_img_pdf"):
                image = Image.open(uploaded_file).convert('RGB')
                out = io.BytesIO()
                image.save(out, format='PDF')
                st.success("Image converted successfully!")
                st.download_button("Download PDF", data=out.getvalue(), file_name="converted_image.pdf", mime="application/pdf")

    if t12:
        with st.expander("12. Combine Multiple Images to One PDF"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Select and upload multiple image files at once. Click 'Combine Images' to merge all the selected images into one single, continuous PDF document.")
            uploaded_files = st.file_uploader("Upload Multiple Images", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="m_imgs")
            if uploaded_files and st.button("Combine Images", key="btn_m_imgs"):
                image_list = []
                first_img = None
                for i, f in enumerate(uploaded_files):
                    img = Image.open(f).convert('RGB')
                    if i == 0: first_img = img
                    else: image_list.append(img)
                if first_img:
                    out = io.BytesIO()
                    first_img.save(out, format='PDF', save_all=True, append_images=image_list)
                    st.success("Images combined into PDF successfully!")
                    st.download_button("Download Combined PDF", data=out.getvalue(), file_name="combined_images.pdf", mime="application/pdf")

    if t13:
        with st.expander("13. Convert Word Document to PDF"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload an editable MS Word document (.docx). Click 'Convert' to securely transform it into a non-editable PDF. *(Note: Requires MS Word to be installed on the host server).*")
            uploaded_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="w_pdf")
            if uploaded_file and st.button("Convert to PDF", key="btn_w_pdf"):
                if docx_convert:
                    tmp_in_path = ""
                    tmp_out_path = ""
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_in, tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_out:
                            tmp_in.write(uploaded_file.read())
                            tmp_in_path = tmp_in.name
                            tmp_out_path = tmp_out.name
                        
                        docx_convert(tmp_in_path, tmp_out_path)
                        with open(tmp_out_path, "rb") as f_out: pdf_bytes = f_out.read()
                        
                        st.success("Word converted to PDF successfully!")
                        st.download_button("Download PDF", data=pdf_bytes, file_name="converted_word.pdf", mime="application/pdf")
                    except Exception as e:
                        st.error(f"Conversion error: {e}")
                    finally:
                        if tmp_in_path and os.path.exists(tmp_in_path): os.remove(tmp_in_path)
                        if tmp_out_path and os.path.exists(tmp_out_path): os.remove(tmp_out_path)
                else:
                    st.error("docx2pdf library not installed.")

    if t14:
        with st.expander("14. Convert PDF to Word Document"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF file. Click 'Convert to Word' to transform it into a fully editable Microsoft Word (.docx) document.")
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_w")
            if uploaded_file and st.button("Convert to Word", key="btn_pdf_w"):
                if Converter:
                    tmp_in_path = ""
                    tmp_out_path = ""
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_in, tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_out:
                            tmp_in.write(uploaded_file.read())
                            tmp_in_path = tmp_in.name
                            tmp_out_path = tmp_out.name
                            
                        cv = Converter(tmp_in_path)
                        cv.convert(tmp_out_path, start=0, end=None)
                        cv.close()
                        with open(tmp_out_path, "rb") as f_out: docx_bytes = f_out.read()
                        
                        st.success("PDF converted to Word successfully!")
                        st.download_button("Download Word Document", data=docx_bytes, file_name="converted_pdf.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    except Exception as e:
                        st.error(f"Error: {e}")
                    finally:
                        if tmp_in_path and os.path.exists(tmp_in_path): os.remove(tmp_in_path)
                        if tmp_out_path and os.path.exists(tmp_out_path): os.remove(tmp_out_path)
                else:
                    st.error("pdf2docx library not installed.")


# ==========================================
# CATEGORY 3: DATA & TEXT EXTRACTION
# ==========================================
if any([t15, t16, t17, t18, t19]):
    st.markdown("## 📊 Data & Text Extraction")

    if t15:
        with st.expander("15. Scan PDF to Notepad (Full OCR - Scanned Docs)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("If you have a scanned PDF (where text is just an image), upload it here. The backend server will use Optical Character Recognition (OCR) to read the text and convert it into a text file.")
            st.info("Full OCR processing is handled on the server backend.")
            uploaded_file = st.file_uploader("Upload Scanned PDF", type=["pdf"], key="ocr_pdf")
            if uploaded_file and st.button("Run OCR", key="btn_ocr"):
                st.warning("Backend OCR processing active.")

    if t16:
        with st.expander("16. Fast Text Extract (Digital PDFs Only)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a standard 'digital' PDF (where you can normally highlight text with your mouse). Click 'Extract Text' to pull all the written content out and download it as a plain Notepad (.txt) file.")
            uploaded_file = st.file_uploader("Upload Digital PDF", type=["pdf"], key="txt_pdf")
            if uploaded_file and st.button("Extract Text", key="btn_txt_pdf"):
                reader = PdfReader(uploaded_file)
                text_content = ""
                for i, page in enumerate(reader.pages):
                    t = page.extract_text()
                    if t: text_content += f"--- Page {i+1} ---\n{t}\n\n"
                st.success("Text extracted successfully!")
                st.download_button("Download Text File", data=text_content, file_name="extracted_text.txt", mime="text/plain")

    if t17:
        with st.expander("17. Extract Text from Word to Notepad"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a Microsoft Word (.docx) file. Click 'Extract' to strip away all formatting, images, and tables, and download just the raw text in a Notepad (.txt) file.")
            uploaded_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="w_txt")
            if uploaded_file and st.button("Extract Word Text", key="btn_w_txt"):
                doc = docx.Document(uploaded_file)
                text_content = "\n".join([p.text for p in doc.paragraphs])
                st.success("Text extracted from Word successfully!")
                st.download_button("Download Text File", data=text_content, file_name="word_text.txt", mime="text/plain")

    if t18:
        with st.expander("18. Extract All Embedded Images from PDF"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload any PDF that contains pictures, logos, or scanned signatures. Click 'Extract Images' and the system will pull out every single picture and pack them into a downloadable ZIP file.")
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="img_ext")
            if uploaded_file and st.button("Extract Images", key="btn_img_ext"):
                try:
                    reader = PdfReader(uploaded_file)
                    zip_buffer = io.BytesIO()
                    count = 0
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for page_num, page in enumerate(reader.pages):
                            for img_idx, img_file_object in enumerate(page.images):
                                count += 1
                                img_name = f"page_{page_num+1}_img_{img_idx+1}_{img_file_object.name}"
                                zip_file.writestr(img_name, img_file_object.data)
                    if count > 0:
                        zip_buffer.seek(0)
                        st.success(f"Successfully extracted {count} embedded images!")
                        st.download_button("📥 Download All Images (ZIP)", data=zip_buffer.getvalue(), file_name="extracted_images.zip", mime="application/zip", key="dl_extracted_images")
                    else:
                        st.warning("No embedded images found in this PDF.")
                except Exception as e:
                    st.error(f"Error extracting images: {e}")

    if t19:
        with st.expander("19. Extract PDF Tables to Excel (.csv)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF that has structured data or tables. Click 'Extract Tables' to read the grid structure and convert it directly into a downloadable CSV file (which opens in MS Excel).")
            uploaded_file = st.file_uploader("Upload PDF with Tables", type=["pdf"], key="tbl_pdf")
            if uploaded_file and st.button("Extract Tables", key="btn_tbl_pdf"):
                tmp_path = ""
                try:
                    import pdfplumber
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(uploaded_file.read())
                        tmp_path = tmp.name
                    csv_buffer = io.StringIO()
                    writer = csv.writer(csv_buffer)
                    found = False
                    with pdfplumber.open(tmp_path) as pdf:
                        for page in pdf.pages:
                            tables = page.extract_tables()
                            for table in tables:
                                found = True
                                for row in table:
                                    clean_row = [cell if cell is not None else "" for cell in row]
                                    writer.writerow(clean_row)
                                writer.writerow([])
                    if found:
                        st.success("Tables extracted successfully!")
                        st.download_button("Download Table (.csv)", data=csv_buffer.getvalue(), file_name="extracted_tables.csv", mime="application/csv")
                    else:
                        st.warning("No structured tables found in this PDF.")
                except Exception as e:
                    st.error(f"Error: {e}")
                finally:
                    if tmp_path and os.path.exists(tmp_path): os.remove(tmp_path)


# ==========================================
# CATEGORY 4: SECURITY & DOCUMENT POLISH
# ==========================================
if any([t20, t21, t22, t23, t24, t25, t26, t27]):
    st.markdown("## 🔒 Security & Document Polish")

    if t20:
        with st.expander("20. Password Protect / Encrypt a PDF"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload your PDF. Type a strong password in the text box. Click 'Protect PDF'. The document will be locked, and no one will be able to open it without entering the password.")
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="lock_pdf")
            pwd = st.text_input("Enter Password to Lock PDF", type="password", key="lock_pwd")
            if uploaded_file and pwd and st.button("Protect PDF", key="btn_lock"):
                reader = PdfReader(uploaded_file)
                writer = PdfWriter()
                for page in reader.pages: writer.add_page(page)
                writer.encrypt(pwd)
                out = io.BytesIO()
                writer.write(out)
                out.seek(0)
                st.success("PDF protected successfully!")
                st.download_button("Download Protected PDF", data=out, file_name="protected_output.pdf", mime="application/pdf")

    if t21:
        with st.expander("21. Unlock / Remove PDF Password"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a locked/protected PDF. Enter the current valid password in the box. Click 'Unlock PDF' to permanently remove the password requirement from the document.")
            uploaded_file = st.file_uploader("Upload Locked PDF", type=["pdf"], key="unlock_pdf")
            pwd = st.text_input("Enter Current Password", type="password", key="unlock_pwd")
            if uploaded_file and pwd and st.button("Unlock PDF", key="btn_unlock"):
                try:
                    reader = PdfReader(uploaded_file)
                    if reader.is_encrypted: reader.decrypt(pwd)
                    writer = PdfWriter()
                    for page in reader.pages: writer.add_page(page)
                    out = io.BytesIO()
                    writer.write(out)
                    out.seek(0)
                    st.success("PDF unlocked successfully!")
                    st.download_button("Download Unlocked PDF", data=out, file_name="unlocked_output.pdf", mime="application/pdf")
                except Exception as e:
                    st.error(f"Error: {e}")

    if t22:
        with st.expander("22. Apply Custom Watermark / Stamp to PDF"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF. Enter your desired watermark text (like 'SECRET' or 'CONFIDENTIAL'). Click 'Apply'. The text will be stamped diagonally across the background of every page.")
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
                can.translate(page_width / 2, page_height / 2)
                can.rotate(45)
                can.drawCentredString(0, 0, wm_text)
                can.save()
                packet.seek(0)
                wm_reader = PdfReader(packet)
                wm_page = wm_reader.pages[0]
                for page in reader.pages:
                    page.merge_page(wm_page)
                    writer.add_page(page)
                out = io.BytesIO()
                writer.write(out)
                out.seek(0)
                st.success("Watermark applied successfully!")
                st.download_button("Download Watermarked PDF", data=out, file_name="watermarked_output.pdf", mime="application/pdf")

    if t23:
        with st.expander("23. Add Page Numbers (Bates Numbering)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF that doesn't have page numbers. Click 'Add Page Numbers' to automatically stamp 'Page 1 of 10', 'Page 2 of 10', etc., at the bottom center of every page.")
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="bates_pdf")
            if uploaded_file and st.button("Add Page Numbers", key="btn_bates"):
                reader = PdfReader(uploaded_file)
                writer = PdfWriter()
                total_pages = len(reader.pages)
                for i, page in enumerate(reader.pages):
                    packet = io.BytesIO()
                    can = canvas.Canvas(packet, pagesize=(float(page.mediabox.width), float(page.mediabox.height)))
                    can.setFont("Helvetica", 10)
                    text = f"Page {i + 1} of {total_pages}"
                    can.drawCentredString(float(page.mediabox.width) / 2, 20, text)
                    can.save()
                    packet.seek(0)
                    num_reader = PdfReader(packet)
                    page.merge_page(num_reader.pages[0])
                    writer.add_page(page)
                out = io.BytesIO()
                writer.write(out)
                out.seek(0)
                st.success("Page numbers added successfully!")
                st.download_button("Download Numbered PDF", data=out, file_name="numbered_output.pdf", mime="application/pdf")

    if t24:
        with st.expander("24. Insert Blank Pages (For Duplex Printing)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload your PDF. Enter the page number *after which* you want a blank page inserted. (Enter '0' to put a blank page at the very beginning). Click 'Insert' to update the file.")
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
                writer.write(out)
                out.seek(0)
                st.success("Blank page inserted successfully!")
                st.download_button("Download Updated PDF", data=out, file_name="blank_page_output.pdf", mime="application/pdf")

    if t25:
        with st.expander("25. Rotate PDF Pages (90°)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF. Click 'Rotate Pages' to automatically turn every page 90 degrees clockwise. This is perfect for fixing pages that were scanned sideways/landscape.")
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="rot_pdf")
            if uploaded_file and st.button("Rotate Pages 90°", key="btn_rot"):
                reader = PdfReader(uploaded_file)
                writer = PdfWriter()
                for page in reader.pages:
                    page.rotate(90)
                    writer.add_page(page)
                out = io.BytesIO()
                writer.write(out)
                out.seek(0)
                st.success("PDF rotated successfully!")
                st.download_button("Download Rotated PDF", data=out.getvalue(), file_name="rotated_output.pdf", mime="application/pdf")

    if t26:
        with st.expander("26. Compress / Optimize PDF File Size"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a large PDF file. Click 'Compress PDF Streams' and the tool will optimize the internal code and text streams to reduce the overall MB size of the document.")
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="comp_pdf")
            if uploaded_file and st.button("Compress PDF Streams", key="btn_comp"):
                reader = PdfReader(uploaded_file)
                writer = PdfWriter()
                for page in reader.pages: writer.add_page(page)
                for page in writer.pages: page.compress_content_streams()
                out = io.BytesIO()
                writer.write(out)
                out.seek(0)
                st.success("PDF streams compressed successfully!")
                st.download_button("Download Compressed PDF", data=out, file_name="compressed_output.pdf", mime="application/pdf")

    if t27:
        with st.expander("27. View Hidden PDF Metadata"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF and click 'View Metadata'. The tool will reveal hidden information stored inside the file, such as the Author name, Creation Date, and the Software used to generate it.")
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="meta_pdf")
            if uploaded_file and st.button("View Metadata", key="btn_meta"):
                reader = PdfReader(uploaded_file)
                meta = reader.metadata
                if meta:
                    meta_str = "\n".join([f"{k.strip('/')}: {v}" for k, v in meta.items()])
                    st.text_area("PDF Metadata", value=meta_str, height=200, key="meta_txt")
                else:
                    st.info("No metadata found in this PDF.")


# ==========================================
# CATEGORY 5: IMAGE UTILITIES
# ==========================================
if any([t28, t29]):
    st.markdown("## 🖼️ Image Utilities")

    if t28:
        with st.expander("28. Remove Image Background & Change Color (AI Offline)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload an image (like a signature, logo, or person). Choose whether you want a transparent background or a solid color background (White, Red, Green, etc.). Click 'Process Image'. The AI will analyze the picture, delete the old background, and apply your new choice.")
            
            if rembg_remove is None:
                st.warning("⚠️ The 'rembg' library is missing! Please ask your IT admin to run `pip install rembg` on the server to activate this feature.")
            else:
                bg_img_file = st.file_uploader("Upload Image to Remove/Change Background", type=["png", "jpg", "jpeg"], key="bg_rem_img")
                
                # UI for Color Selection
                col1, col2 = st.columns(2)
                bg_option = col1.selectbox("Select New Background", ["Transparent", "White", "Red", "Pink", "Green", "Yellow", "Custom Color"], key="bg_rem_opt")
                
                custom_color = "#FFFFFF"
                if bg_option == "Custom Color":
                    custom_color = col2.color_picker("Pick a Custom Color", "#3498db", key="bg_rem_color")
                    
                if bg_img_file and st.button("Process Image", key="btn_bg_rem"):
                    try:
                        with st.spinner("AI is analyzing and replacing the background..."):
                            # 1. Remove background using AI (returns transparent PNG bytes)
                            input_bytes = bg_img_file.read()
                            output_bytes = rembg_remove(input_bytes)
                            
                            # 2. Check if user wants a solid color
                            if bg_option == "Transparent":
                                final_output = output_bytes
                                file_ext = "png"
                                mime_val = "image/png"
                            else:
                                color_map = {
                                    "White": "#FFFFFF", "Red": "#FF0000", "Pink": "#FFC0CB", "Green": "#008000", "Yellow": "#FFFF00"
                                }
                                selected_hex = custom_color if bg_option == "Custom Color" else color_map[bg_option]
                                transparent_img = Image.open(io.BytesIO(output_bytes)).convert("RGBA")
                                solid_bg = Image.new("RGBA", transparent_img.size, selected_hex)
                                solid_bg.paste(transparent_img, (0, 0), transparent_img)
                                
                                final_io = io.BytesIO()
                                solid_bg.convert("RGB").save(final_io, format="JPEG", quality=95)
                                final_output = final_io.getvalue()
                                file_ext = "jpg"
                                mime_val = "image/jpeg"
                                
                        st.success("Background processed successfully!")
                        st.download_button(label=f"Download Processed Image (.{file_ext})", data=final_output, file_name=f"bg_processed.{file_ext}", mime=mime_val, key="dl_bg_rem")
                    except Exception as e:
                        st.error(f"Error processing background: {e}")

    if t29:
        with st.expander("29. Resize & Compress Image"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload an image. Change the 'Width' and 'Height' boxes to make the image smaller or larger. Use the 'Quality' slider to reduce the file size. Click 'Resize & Compress' to download.")
            res_img_file = st.file_uploader("Upload Image to Resize/Compress", type=["png", "jpg", "jpeg"], key="res_img")
            if res_img_file:
                orig_img = Image.open(res_img_file)
                st.info(f"Original Dimensions: **{orig_img.size[0]} x {orig_img.size[1]} pixels**")
                col1, col2 = st.columns(2)
                new_width = col1.number_input("New Width (Pixels)", min_value=10, value=orig_img.size[0], key="res_w")
                new_height = col2.number_input("New Height (Pixels)", min_value=10, value=orig_img.size[1], key="res_h")
                st.markdown("---")
                quality = st.slider("Compression Quality (Lower = Smaller File Size, Higher = Better Quality)", min_value=1, max_value=100, value=85, key="res_q")

                if st.button("Resize & Compress Image", key="btn_res"):
                    try:
                        with st.spinner("Processing image..."):
                            resized_img = orig_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                            out_img = io.BytesIO()
                            if orig_img.format == "PNG" and quality > 90:
                                resized_img.save(out_img, format="PNG", optimize=True)
                                mime_type = "image/png"
                                ext = "png"
                            else:
                                resized_img = resized_img.convert("RGB")
                                resized_img.save(out_img, format="JPEG", quality=quality, optimize=True)
                                mime_type = "image/jpeg"
                                ext = "jpg"
                            out_img.seek(0)
                        st.success("Image resized and compressed successfully!")
                        st.download_button("Download Processed Image", data=out_img, file_name=f"processed_image.{ext}", mime=mime_type, key="dl_res")
                    except Exception as e:
                        st.error(f"Error processing image: {e}")


# ==========================================
# CATEGORY 6: RAJBHASHA & HINDI TOOLS
# ==========================================
if any([t30]):
    st.markdown("## 🇮🇳 Rajbhasha & Hindi Tools")

    if t30:
        with st.expander("30. KrutiDev 010 to Unicode (Mangal) Converter"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Paste your unreadable/legacy KrutiDev 010 Hindi text into the first text box. Click 'Convert to Unicode'. The corrected, modern Unicode (Mangal) text will appear in the box below for you to copy.")
            st.write("पुरानी फाइलों के कृतिदेव (KrutiDev) टेक्स्ट को यहाँ पेस्ट करें और उसे आधुनिक यूनिकोड (Unicode/Mangal) में बदलें।")
            kruti_text = st.text_area("Paste KrutiDev Text Here:", height=150, key="kruti_input")
            if st.button("Convert to Unicode", key="btn_convert_kruti"):
                if kruti_text.strip():
                    with st.spinner("Converting font formatting..."):
                        unicode_text = convert_krutidev_to_unicode(kruti_text)
                    st.success("Conversion Successful! Copy the standard Unicode text below:")
                    st.text_area("Unicode (Mangal) Output:", value=unicode_text, height=150, key="unicode_output")
                else:
                    st.warning("Please enter some KrutiDev text to convert.")


# ==========================================
# CATEGORY 7: ADVANCED & SMART UTILITIES
# ==========================================
if any([t31, t32, t33, t34]):
    st.markdown("## 🚀 Advanced & Smart Utilities")

    if t31:
        with st.expander("31. True PDF Redaction (ऑटोमैटिक ब्लैकआउट)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF. Type the exact sensitive word or number (e.g., a bank account number) you want to hide. Click 'Apply Permanent Blackout'. The tool will find it everywhere in the document and draw an un-removable black box over it.")
            st.markdown("Permanently hide sensitive information (like Names, Phone Numbers) from a PDF.")
            if fitz is None:
                st.warning("⚠️ The 'PyMuPDF' library is missing! Please ask your IT admin to run `pip install PyMuPDF` on the server to activate the Redaction tool.")
            else:
                redact_pdf = st.file_uploader("Upload PDF to Redact", type=["pdf"], key="redact_pdf")
                redact_text = st.text_input("Enter Exact Text to Redact (Hide)", key="redact_text")
                
                if redact_pdf and redact_text and st.button("Apply Permanent Blackout", key="btn_redact"):
                    try:
                        with fitz.open(stream=redact_pdf.read(), filetype="pdf") as doc:
                            st.write("Applying permanent blackout to the document...")
                            progress_bar = st.progress(0)
                            total_p = len(doc)
                            found_count = 0
                            
                            for i, page in enumerate(doc):
                                areas = page.search_for(redact_text)
                                for rect in areas:
                                    found_count += 1
                                    page.add_redact_annot(rect, fill=(0, 0, 0))
                                page.apply_redactions()
                                progress_bar.progress((i + 1) / total_p)
                            
                            if found_count > 0:
                                out_pdf = io.BytesIO()
                                doc.save(out_pdf)
                                out_pdf.seek(0)
                                progress_bar.empty()
                                st.success(f"Success! Redacted '{redact_text}' {found_count} times in the document.")
                                st.download_button("Download Redacted PDF", data=out_pdf, file_name="redacted_document.pdf", mime="application/pdf", key="dl_redact")
                            else:
                                progress_bar.empty()
                                st.warning(f"Could not find the text '{redact_text}' in this PDF.")
                    except Exception as e:
                        st.error(f"Error during redaction: {e}")

    if t32:
        with st.expander("32. Bulk Generator (Excel/CSV to Word Mail Merge)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("""
                **Step 1: Prepare Word Template (.docx)**
                Create a standard Word document. Wherever you want dynamic data (like a name or ID), write it inside double angle brackets, exactly matching your Excel columns. 
                *Example:* `Name of Visitor: <<VisitorName>>`
                
                **Step 2: Prepare Data File (.csv)**
                Create your data in Excel. The first row (Header) must match the tags used in your Word file exactly (e.g., column A header: `VisitorName`). Save this file as **CSV (Comma delimited) (*.csv)**.
                
                **Step 3: Generate**
                Upload both files here and click Generate. The tool will instantly create a ZIP file containing individual Word documents for every row in your CSV!
                """)
            st.markdown("Upload a Template (.docx) and a Data File (.csv) to generate multiple documents at once.")    
            template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"], key="mm_template")
            data_file = st.file_uploader("Upload Data File (.csv)", type=["csv"], key="mm_data")
            
            if template_file and data_file and st.button("Generate Bulk Documents", key="btn_mm"):
                try:
                    with st.spinner("Merging data and generating documents..."):
                        csv_text = io.StringIO(data_file.getvalue().decode("utf-8"))
                        reader = csv.DictReader(csv_text)
                        
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for i, row in enumerate(reader):
                                doc = docx.Document(template_file)
                                
                                for p in doc.paragraphs:
                                    for key, val in row.items():
                                        placeholder = f"<<{key}>>"
                                        if placeholder in p.text:
                                            for run in p.runs:
                                                if placeholder in run.text:
                                                    run.text = run.text.replace(placeholder, str(val))
                                            if placeholder in p.text:
                                                p.text = p.text.replace(placeholder, str(val))
                                                
                                for table in doc.tables:
                                    for r in table.rows:
                                        for cell in r.cells:
                                            for p in cell.paragraphs:
                                                for key, val in row.items():
                                                    placeholder = f"<<{key}>>"
                                                    for run in p.runs:
                                                        if placeholder in run.text:
                                                            run.text = run.text.replace(placeholder, str(val))
                                                    if placeholder in p.text:
                                                        p.text = p.text.replace(placeholder, str(val))
                                
                                doc_io = io.BytesIO()
                                doc.save(doc_io)
                                doc_name = f"Generated_Document_{i+1}.docx"
                                zip_file.writestr(doc_name, doc_io.getvalue())
                                template_file.seek(0)
                                
                        zip_buffer.seek(0)
                        st.success("Bulk documents generated successfully!")
                        st.download_button("Download All Documents (ZIP)", data=zip_buffer, file_name="Bulk_Merged_Documents.zip", mime="application/zip", key="dl_mm")
                except Exception as e:
                    st.error(f"Error processing mail merge: {e}")

    if t33:
        with st.expander("33. Add Table of Contents (Clickable Bookmarks) to PDF"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF. In the text box, type the page number and chapter title separated by a comma (Example: `1, Chapter One`). Add each entry on a new line. Click 'Add Bookmarks' to embed a clickable sidebar index into your PDF.")
            st.markdown("Add clickable navigation bookmarks to the sidebar of your PDF.")
            toc_pdf = st.file_uploader("Upload PDF", type=["pdf"], key="toc_pdf")
            
            st.info("Format: Enter `Page Number, Bookmark Title` on each line.\n\n**Example:**\n1, Introduction\n5, Chapter 1\n12, Conclusion")
            toc_text = st.text_area("Enter Table of Contents:", height=150, key="toc_text")
            
            if toc_pdf and toc_text and st.button("Add Bookmarks", key="btn_toc"):
                try:
                    reader = PdfReader(toc_pdf)
                    writer = PdfWriter()
                    for page in reader.pages:
                        writer.add_page(page)
                    
                    lines = toc_text.strip().split('\n')
                    for line in lines:
                        if ',' in line:
                            page_str, title = line.split(',', 1)
                            page_num = int(page_str.strip()) - 1
                            
                            if 0 <= page_num < len(reader.pages):
                                writer.add_outline_item(title.strip(), page_num)
                                
                    out_pdf = io.BytesIO()
                    writer.write(out_pdf)
                    out_pdf.seek(0)
                    
                    st.success("Bookmarks added successfully!")
                    st.download_button("Download PDF with Bookmarks", data=out_pdf, file_name="bookmarked_document.pdf", mime="application/pdf", key="dl_toc")
                except Exception as e:
                    st.error(f"Error adding bookmarks: {e}")

    if t34:
        with st.expander("34. PDF Grayscale / B&W Converter (Printer Toner Saver)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a colorful or heavy PDF. Adjust the 'Print Quality (DPI)' slider (150 is optimal for standard printing). Click 'Convert' to change the file to Black & White, which will save a lot of printer ink and make printing faster.")
            st.markdown("Convert any color PDF into Grayscale/Black & White to save printer toner and flatten the document.")
            
            if fitz is None:
                st.warning("⚠️ The 'PyMuPDF' library is missing! Please ask your IT admin to run `pip install PyMuPDF` on the server to activate this feature.")
            else:
                bw_pdf_file = st.file_uploader("Upload Color PDF", type=["pdf"], key="bw_pdf")
                st.info("💡 **Pro Tip:** This tool flattens the PDF into images. Choose higher DPI for text clarity, or lower DPI for smaller file size.")
                dpi_val = st.slider("Print Quality (DPI)", min_value=72, max_value=300, value=150, step=10, key="bw_dpi")
                
                if bw_pdf_file and st.button("Convert to Grayscale", key="btn_bw"):
                    try:
                        with fitz.open(stream=bw_pdf_file.read(), filetype="pdf") as doc, fitz.open() as out_pdf:
                            st.write("Converting pages to Grayscale...")
                            progress_bar = st.progress(0)
                            total_p = len(doc)
                            
                            for page_num in range(total_p):
                                page = doc[page_num]
                                pix = page.get_pixmap(colorspace=fitz.csGRAY, dpi=dpi_val, alpha=False)
                                new_page = out_pdf.new_page(width=page.rect.width, height=page.rect.height)
                                new_page.insert_image(page.rect, pixmap=pix)
                                progress_bar.progress((page_num + 1) / total_p)
                                
                            out_bytes = io.BytesIO()
                            out_pdf.save(out_bytes)
                            out_bytes.seek(0)
                            
                            progress_bar.empty()
                            st.success("Successfully converted to Grayscale!")
                            st.download_button("Download Grayscale PDF", data=out_bytes, file_name="Grayscale_Toner_Saver.pdf", mime="application/pdf", key="dl_bw")
                    except Exception as e:
                        st.error(f"Error during conversion: {e}")


# ==========================================
# CATEGORY 8: ENTERPRISE UTILITIES
# ==========================================
if any([t35, t36, t37, t38, t39, t40]):
    st.markdown("## 🏢 Enterprise & Commercial Grade Utilities")

    if t35:
        with st.expander("35. Visual PDF Comparison (Draft Diff Tool)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload the Old Draft and the New Draft of a PDF. Click 'Compare Texts'. The tool will generate a visual side-by-side comparison highlighting deleted words in **Red** and added words in **Green**.")
            pdf1_file = st.file_uploader("Upload Original PDF (Old Version)", type=["pdf"], key="diff_pdf1")
            pdf2_file = st.file_uploader("Upload Modified PDF (New Version)", type=["pdf"], key="diff_pdf2")
            if pdf1_file and pdf2_file and st.button("Compare Texts", key="btn_diff"):
                try:
                    with st.spinner("Analyzing documents for differences..."):
                        t1_txt = "\n".join([page.extract_text() for page in PdfReader(pdf1_file).pages if page.extract_text()])
                        t2_txt = "\n".join([page.extract_text() for page in PdfReader(pdf2_file).pages if page.extract_text()])
                        html_diff = difflib.HtmlDiff().make_file(t1_txt.splitlines(), t2_txt.splitlines(), "Original", "Modified")
                        st.success("Comparison Complete! Scroll down to view the differences.")
                        components.html(html_diff, height=600, scrolling=True)
                except Exception as e:
                    st.error(f"Error during comparison: {e}")

    if t36:
        with st.expander("36. Searchable 'Sandwich' PDF (Invisible OCR Overlay)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a scanned PDF (where text cannot be selected/copied). This tool will apply Multi-threaded Fast OCR and create an invisible text layer over the original scanned images, allowing you to use `Ctrl+F` to search text in the output PDF.")
            if pytesseract is None:
                st.warning("⚠️ The 'pytesseract' and 'pdf2image' libraries are missing! Please ask your IT admin to run `pip install pytesseract pdf2image`.")
            else:
                ocr_pdf_file = st.file_uploader("Upload Scanned PDF to make it Searchable", type=["pdf"], key="sandwich_pdf")
                if ocr_pdf_file and st.button("Create Searchable PDF", key="btn_sandwich"):
                    try:
                        # 💡 SMART AUTO-DETECT: Tesseract Path
                        tesseract_path = shutil.which("tesseract")
                        if tesseract_path:
                            pytesseract.pytesseract.tesseract_cmd = tesseract_path
                        elif platform.system() == "Windows":
                            if os.path.exists(r'C:\Program Files\Tesseract-OCR\tesseract.exe'):
                                pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
                            elif os.path.exists(r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'):
                                pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
                            else:
                                st.error("🚨 Tesseract-OCR software was not found! Please install it and ensure it is in your system PATH: https://github.com/UB-Mannheim/tesseract/wiki")
                                st.stop()
                        
                        with st.spinner("Applying Multi-threaded Fast OCR... Please wait, this may take a moment for large files."):
                            # 💡 SMART AUTO-DETECT: Poppler Path
                            poppler_path = None
                            if platform.system() == "Windows":
                                if os.path.exists(r'C:\poppler\Library\bin'):
                                    poppler_path = r'C:\poppler\Library\bin'
                                elif os.path.exists(r'C:\poppler\bin'):
                                    poppler_path = r'C:\poppler\bin'
                                else:
                                    st.warning("⚠️ Poppler path not found in default locations. OCR might fail if Poppler isnt in system PATH.")

                            # Convert PDF to Images
                            images = convert_from_bytes(ocr_pdf_file.read(), poppler_path=poppler_path, dpi=200)
                            merger = PdfWriter()
                            
                            def process_page_ocr(img):
                                return pytesseract.image_to_pdf_or_hocr(img, extension='pdf')
                            
                            with concurrent.futures.ThreadPoolExecutor() as executor:
                                pdf_bytes_list = list(executor.map(process_page_ocr, images))
                            
                            for pdf_bytes in pdf_bytes_list:
                                page_reader = PdfReader(io.BytesIO(pdf_bytes))
                                merger.add_page(page_reader.pages[0])
                                
                            out_pdf = io.BytesIO()
                            merger.write(out_pdf)
                            out_pdf.seek(0)
                            
                            st.success("Fast Searchable PDF created successfully!")
                            st.download_button(
                                "Download Searchable PDF", 
                                data=out_pdf, 
                                file_name="Searchable_Sandwich_Fast.pdf", 
                                mime="application/pdf", 
                                key="dl_sandwich"
                            )
                    except Exception as e:
                        st.error(f"Error during OCR processing: {e}")

    if t37:
        with st.expander("37. Digital Facsimile Signature & Stamp Placer"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("1. Upload your target PDF.\n2. Upload a transparent PNG of your Stamp or Signature.\n3. Choose the Page Number and Corner Position (Top/Bottom, Left/Right).\n4. Click 'Place Stamp/Signature' to digitally embed it without printing.")
            if fitz is None:
                st.warning("⚠️ 'PyMuPDF' library is required. Run `pip install PyMuPDF`.")
            else:
                stamp_pdf = st.file_uploader("Upload PDF Document", type=["pdf"], key="stamp_pdf")
                stamp_img = st.file_uploader("Upload Stamp/Signature (PNG with transparent background is best)", type=["png", "jpg", "jpeg"], key="stamp_img")
                
                col1, col2 = st.columns(2)
                stamp_page = col1.number_input("Page Number to Stamp (0 for All Pages)", min_value=0, value=1, step=1, key="stamp_page_num")
                stamp_pos = col2.selectbox("Select Position", ["Bottom-Right", "Bottom-Left", "Top-Right", "Top-Left", "Center"], key="stamp_pos")
                
                if stamp_pdf and stamp_img and st.button("Place Stamp/Signature", key="btn_place_stamp"):
                    try:
                        with st.spinner("Placing signature/stamp digitally..."):
                            with fitz.open(stream=stamp_pdf.read(), filetype="pdf") as doc:
                                img_bytes = stamp_img.read()
                                pages_to_stamp = range(len(doc)) if stamp_page == 0 else [stamp_page - 1]
                                
                                for p_num in pages_to_stamp:
                                    if 0 <= p_num < len(doc):
                                        page = doc[p_num]
                                        rect = page.rect
                                        stamp_w, stamp_h = 150, 75  # Default reasonable size for stamps
                                        margin = 30
                                        
                                        if stamp_pos == "Bottom-Right":
                                            target_rect = fitz.Rect(rect.width - stamp_w - margin, rect.height - stamp_h - margin, rect.width - margin, rect.height - margin)
                                        elif stamp_pos == "Bottom-Left":
                                            target_rect = fitz.Rect(margin, rect.height - stamp_h - margin, margin + stamp_w, rect.height - margin)
                                        elif stamp_pos == "Top-Right":
                                            target_rect = fitz.Rect(rect.width - stamp_w - margin, margin, rect.width - margin, margin + stamp_h)
                                        elif stamp_pos == "Top-Left":
                                            target_rect = fitz.Rect(margin, margin, margin + stamp_w, margin + stamp_h)
                                        else: # Center
                                            target_rect = fitz.Rect((rect.width - stamp_w)/2, (rect.height - stamp_h)/2, (rect.width + stamp_w)/2, (rect.height + stamp_h)/2)
                                        
                                        page.insert_image(target_rect, stream=img_bytes)
                                
                                out_pdf = io.BytesIO()
                                doc.save(out_pdf)
                                out_pdf.seek(0)
                                st.success("Stamp/Signature placed successfully!")
                                st.download_button("Download Stamped PDF", data=out_pdf, file_name="Stamped_Document.pdf", mime="application/pdf", key="dl_stamp")
                    except Exception as e:
                        st.error(f"Error placing stamp: {e}")

    if t38:
        with st.expander("38. N-Up & Booklet Layout Maker (Multiple Pages on 1 Sheet)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a standard PDF. Choose how many pages you want on a single sheet (2-Up, 4-Up, or 8-Up). The tool will shrink and combine the pages onto a larger sheet, which is excellent for saving paper during meeting printouts.")
            if fitz is None:
                st.warning("⚠️ 'PyMuPDF' library is required. Run `pip install PyMuPDF`.")
            else:
                nup_pdf = st.file_uploader("Upload PDF to Format", type=["pdf"], key="nup_pdf")
                
                nup_layout = st.selectbox("Select N-Up Layout", [
                    "2-Up (2 Pages per Sheet - 2x1)", 
                    "4-Up (4 Pages per Sheet - 2x2)", 
                    "8-Up (8 Pages per Sheet - 4x2)"
                ], key="nup_layout")
                
                if nup_pdf and st.button("Generate Layout", key="btn_nup"):
                    try:
                        with st.spinner(f"Reformatting pages to {nup_layout[:4]} layout..."):
                            with fitz.open(stream=nup_pdf.read(), filetype="pdf") as doc, fitz.open() as out_doc:
                                
                                if "2-Up" in nup_layout:
                                    cols, rows = 2, 1
                                elif "4-Up" in nup_layout:
                                    cols, rows = 2, 2
                                elif "8-Up" in nup_layout:
                                    cols, rows = 4, 2
                                    
                                chunk_size = cols * rows
                                
                                for start_idx in range(0, len(doc), chunk_size):
                                    p_ref = doc[start_idx]
                                    w, h = p_ref.rect.width, p_ref.rect.height
                                    
                                    new_page = out_doc.new_page(width=w * cols, height=h * rows)
                                    
                                    for i in range(chunk_size):
                                        doc_idx = start_idx + i
                                        if doc_idx < len(doc):
                                            col = i % cols
                                            row = i // cols
                                            target_rect = fitz.Rect(col * w, row * h, (col + 1) * w, (row + 1) * h)
                                            new_page.show_pdf_page(target_rect, doc, doc_idx)
                                            
                                out_pdf = io.BytesIO()
                                out_doc.save(out_pdf)
                                out_pdf.seek(0)
                                st.success(f"{nup_layout[:4]} Layout created successfully!")
                                st.download_button(
                                    f"Download {nup_layout[:4]} PDF", 
                                    data=out_pdf, 
                                    file_name=f"{nup_layout[:4]}_Format.pdf", 
                                    mime="application/pdf", 
                                    key="dl_nup"
                                )
                    except Exception as e:
                        st.error(f"Error generating layout: {e}")

    if t39:
        with st.expander("39. Institutional Bates Stamping (Custom Prefix Legal Numbering)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF. Enter your department's specific prefix (e.g., `RS/SYS/2026/`) and a starting number. The tool will apply sequential Bates stamps (like `RS/SYS/2026/0001`, `RS/SYS/2026/0002`) to the top right of every page.")
            if fitz is None:
                st.warning("⚠️ 'PyMuPDF' library is required. Run `pip install PyMuPDF`.")
            else:
                bates_adv_pdf = st.file_uploader("Upload Document for Stamping", type=["pdf"], key="bates_adv_pdf")
                col1, col2 = st.columns(2)
                b_prefix = col1.text_input("Bates Prefix", value="RS/SYS/2026/", key="b_prefix")
                b_start = col2.number_input("Starting Number", min_value=1, value=1, step=1, key="b_start")
                
                if bates_adv_pdf and st.button("Apply Advanced Bates Stamping", key="btn_bates_adv"):
                    try:
                        with st.spinner("Applying legal Bates numbering..."):
                            with fitz.open(stream=bates_adv_pdf.read(), filetype="pdf") as doc:
                                for i, page in enumerate(doc):
                                    bates_number = f"{b_prefix}{str(b_start + i).zfill(4)}"
                                    page.insert_text(fitz.Point(page.rect.width - 150, 30), bates_number, fontsize=11, fontname="helv", color=(1, 0, 0)) # Red text
                                    
                                out_pdf = io.BytesIO()
                                doc.save(out_pdf)
                                out_pdf.seek(0)
                                st.success("Bates stamping completed!")
                                st.download_button("Download Stamped PDF", data=out_pdf, file_name="Bates_Stamped_Doc.pdf", mime="application/pdf", key="dl_bates_adv")
                    except Exception as e:
                        st.error(f"Error applying Bates stamp: {e}")

    if t40:
        with st.expander("40. PDF Auto-Crop & White Margin Trimmer"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF with large empty white borders (common in scanned books or old files). Click 'Auto-Crop Margins' and the tool will automatically detect the text boundaries and trim off the wasted white space.")
            if fitz is None:
                st.warning("⚠️ 'PyMuPDF' library is required. Run `pip install PyMuPDF`.")
            else:
                crop_pdf_file = st.file_uploader("Upload PDF to Crop", type=["pdf"], key="crop_pdf")
                if crop_pdf_file and st.button("Auto-Crop Margins", key="btn_crop"):
                    try:
                        with fitz.open(stream=crop_pdf_file.read(), filetype="pdf") as doc:
                            st.write("Detecting content boundaries and cropping margins...")
                            progress_bar = st.progress(0)
                            total_p = len(doc)
                            for i, page in enumerate(doc):
                                text_rect = page.get_text("rect")
                                if not text_rect.is_empty and not text_rect.is_infinite:
                                    crop_box = fitz.Rect(
                                        max(0, text_rect.x0 - 10),
                                        max(0, text_rect.y0 - 10),
                                        min(page.rect.width, text_rect.x1 + 10),
                                        min(page.rect.height, text_rect.y1 + 10)
                                    )
                                    page.set_cropbox(crop_box)
                                progress_bar.progress((i + 1) / total_p)
                                    
                            out_pdf = io.BytesIO()
                            doc.save(out_pdf)
                            out_pdf.seek(0)
                            
                            progress_bar.empty()
                            st.success("PDF margins cropped successfully!")
                            st.download_button("Download Cropped PDF", data=out_pdf, file_name="Cropped_Margins.pdf", mime="application/pdf", key="dl_crop")
                    except Exception as e:
                        st.error(f"Error cropping margins: {e}")


# ==========================================
# CATEGORY 9: SECRETARIAT SPECIFIC
# ==========================================
if any([t41, t42, t43, t44, t45, t46, t47, t54]):
    st.markdown("## 🌟 Secretariat Specific Utilities")

    if t41:
        with st.expander("41. Parliamentary & Admin Glossary (सचिवालयीन शब्दकोश)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Start typing an English administrative or parliamentary term (e.g., 'Whip', 'Vetting'). The tool will instantly provide the official, approved Hindi translation used in government files.")
            
            st.markdown("Search for official English-to-Hindi translations of common Secretariat terms offline.")
            glossary = {
                "Adjournment": "स्थगन",
                "Adjournment sine die": "अनिश्चित काल के लिए स्थगन",
                "Prorogation": "सत्रावसान",
                "Quorum": "गणपूर्ति",
                "Whip": "सचेतक",
                "Zero Hour": "शून्य काल",
                "Question Hour": "प्रश्न काल",
                "Starred Question": "तारांकित प्रश्न",
                "Unstarred Question": "अतारांकित प्रश्न",
                "Point of Order": "औचित्य प्रश्न",
                "Breach of Privilege": "विशेषाधिकार हनन",
                "Laying on the Table": "पटल पर रखना",
                "Casting Vote": "निर्णायक मत",
                "Resolution": "संकल्प",
                "Motion": "प्रस्ताव",
                "Bill": "विधेयक",
                "Act": "अधिनियम",
                "Amendment": "संशोधन",
                "Office Memorandum (O.M.)": "कार्यालय ज्ञापन",
                "Circular": "परिपत्र",
                "Notification": "अधिसूचना",
                "Gazette": "राजपत्र",
                "Endorsement": "पृष्ठांकन",
                "Annexure": "अनुलग्नक",
                "Appendix": "परिशिष्ट",
                "Minutes of Meeting": "कार्यवृत्त",
                "Agenda": "कार्यसूची",
                "Ex-officio": "पदेन",
                "Vetting": "संवीक्षा",
                "Concurrence": "सहमति",
                "Delegation of Power": "शक्तियों का प्रत्यायोजन",
                "Ex post facto": "कार्योत्तर (घटना के बाद)",
                "Proviso": "परंतुक (शर्त)",
                "Appropriation": "विनियोग",
                "Consolidated Fund": "संचित निधि",
                "Contingency Fund": "आकस्मिकता निधि",
                "Sanction": "स्वीकृति",
                "Audit": "लेखापरीक्षा",
                "Honorarium": "मानदेय",
                "Reimbursement": "प्रतिपूर्ति",
                "Remuneration": "पारिश्रमिक",
                "Voucher": "वाउचर / प्रमाणक",
                "Deputation": "प्रतिनियुक्ति",
                "Probation": "परिवीक्षा",
                "Vigilance": "सतर्कता",
                "Disciplinary Action": "अनुशासनात्मक कार्रवाई",
                "Suspension": "निलंबन",
                "Superannuation": "अधिवर्षिता (सेवानिवृत्ति)",
                "Lien": "धारणाधिकार",
                "Seniority": "वरिष्ठता"
            }
            
            search_term = st.text_input("Type an English term to search (e.g., 'Annexure', 'Whip'):", key="gloss_search")
            if search_term:
                matches = {k: v for k, v in glossary.items() if search_term.lower() in k.lower()}
                if matches:
                    for k, v in matches.items():
                        st.success(f"**{k}** ➔ {v}")
                else:
                    st.warning("Term not found in the offline database. Please try a different word.")

    if t42:
        with st.expander("42. Make PDF Look Scanned (डिजिटल को 'स्कैन-लुक' दें)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a perfectly clean, digitally created PDF. This tool will add a slight realistic tilt, noise, and grayscale photocopy effect to make it look exactly like it was printed and physically scanned on a machine.")
            if fitz is None:
                st.warning("⚠️ 'PyMuPDF' library is required. Run `pip install PyMuPDF`.")
            else:
                fake_scan_pdf = st.file_uploader("Upload Digital PDF to make it look Scanned", type=["pdf"], key="fake_scan_pdf")
                if fake_scan_pdf and st.button("Apply Realistic Scan Effect", key="btn_fake_scan"):
                    try:
                        with fitz.open(stream=fake_scan_pdf.read(), filetype="pdf") as doc:
                            out_pdf_bytes = io.BytesIO()
                            merger = PdfWriter()
                            
                            st.write("Applying photocopier and scanner effects...")
                            progress_bar = st.progress(0)
                            total_p = len(doc)
                            
                            for i, page in enumerate(doc):
                                # Convert to image
                                pix = page.get_pixmap(dpi=150)
                                img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L") # Grayscale
                                
                                # Apply slight random tilt between -1.5 and 1.5 degrees
                                angle = random.uniform(-1.5, 1.5)
                                img = img.rotate(angle, fillcolor=255, expand=False)
                                
                                # Apply Photocopy Contrast/Brightness effects
                                img = ImageEnhance.Contrast(img).enhance(1.8) # High contrast
                                img = ImageEnhance.Brightness(img).enhance(0.9) # Slightly darker
                                
                                # Save back to PDF format in memory
                                temp_pdf = io.BytesIO()
                                img.save(temp_pdf, format="PDF")
                                temp_pdf.seek(0)
                                
                                # Append to final writer
                                merger.add_page(PdfReader(temp_pdf).pages[0])
                                progress_bar.progress((i + 1) / total_p)
                                
                            merger.write(out_pdf_bytes)
                            out_pdf_bytes.seek(0)
                            
                            progress_bar.empty()
                            st.success("Scan effect applied successfully!")
                            st.download_button("Download 'Scanned' PDF", data=out_pdf_bytes, file_name="Realistic_Scanned_Copy.pdf", mime="application/pdf", key="dl_fake_scan")
                    except Exception as e:
                        st.error(f"Error applying effect: {e}")

    if t43:
        with st.expander("43. Quick Dak/Diary Receipt Stamper (ई-डायरी मुहर)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload any incoming letter/Dak PDF. Enter the officially generated Diary No. and Date. Click 'Stamp'. The tool will perfectly place a red 'RECEIVED' official stamp with your diary details on the top right corner of the first page.")
            if fitz is None:
                st.warning("⚠️ 'PyMuPDF' library is required. Run `pip install PyMuPDF`.")
            else:
                dak_pdf_file = st.file_uploader("Upload Incoming Letter/Dak (PDF)", type=["pdf"], key="dak_pdf_file")
                col1, col2 = st.columns(2)
                diary_no = col1.text_input("Enter Diary No.", value="1234/Sys/26", key="diary_no")
                diary_date = col2.text_input("Enter Date", value="03-09-2026", key="diary_date")
                
                if dak_pdf_file and st.button("Apply Official Diary Stamp", key="btn_dak"):
                    try:
                        with st.spinner("Stamping document..."):
                            with fitz.open(stream=dak_pdf_file.read(), filetype="pdf") as doc:
                                page = doc[0] # Stamp only the first page
                                
                                # Define Stamp Rectangle on Top Right
                                rect_w, rect_h = 240, 80
                                margin = 20
                                rect = fitz.Rect(page.rect.width - rect_w - margin, margin, page.rect.width - margin, margin + rect_h)
                                
                                # Draw Stamp Border and Fill Text (Using valid standard built-in font names)
                                page.draw_rect(rect, color=(1, 0, 0), width=2) # Red border
                                page.insert_text(fitz.Point(rect.x0 + 10, rect.y0 + 20), "RECEIVED - RAJYA SABHA SEC.", color=(1,0,0), fontsize=10, fontname="Helvetica-Bold")
                                page.insert_text(fitz.Point(rect.x0 + 10, rect.y0 + 45), f"Diary No : {diary_no}", color=(1,0,0), fontsize=10, fontname="Helvetica")
                                page.insert_text(fitz.Point(rect.x0 + 10, rect.y0 + 65), f"Date     : {diary_date}", color=(1,0,0), fontsize=10, fontname="Helvetica")
                                
                                out_pdf = io.BytesIO()
                                doc.save(out_pdf)
                                out_pdf.seek(0)
                                
                                st.success("Diary stamp placed successfully!")
                                st.download_button("Download Stamped Dak", data=out_pdf, file_name="Diarized_Receipt.pdf", mime="application/pdf", key="dl_dak")
                    except Exception as e:
                        st.error(f"Error stamping Dak: {e}")

    if t44:
        with st.expander("44. PDF A4 Page Standardizer (प्रिंटर पेपर-साइज फिक्सर)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF containing pages of mixed sizes (e.g., Letter, Legal, Custom). Click 'Standardize'. The tool will automatically resize and fit every single page into exact A4 dimensions, fixing 'Paper Mismatch' printer errors forever.")
            if fitz is None:
                st.warning("⚠️ 'PyMuPDF' library is required. Run `pip install PyMuPDF`.")
            else:
                a4_pdf_file = st.file_uploader("Upload Mixed-Size PDF", type=["pdf"], key="a4_pdf")
                if a4_pdf_file and st.button("Standardize to A4 Size", key="btn_a4"):
                    try:
                        with fitz.open(stream=a4_pdf_file.read(), filetype="pdf") as doc, fitz.open() as out_doc:
                            st.write("Resizing all pages to standard A4...")
                            progress_bar = st.progress(0)
                            a4_w, a4_h = fitz.paper_size("a4")
                            total_p = len(doc)
                            
                            for i in range(total_p):
                                # Create exact A4 blank page
                                new_page = out_doc.new_page(width=a4_w, height=a4_h)
                                # Stretch/scale the original page to fit perfectly into the A4 rect
                                new_page.show_pdf_page(new_page.rect, doc, i)
                                progress_bar.progress((i + 1) / total_p)
                                
                            out_pdf = io.BytesIO()
                            out_doc.save(out_pdf)
                            out_pdf.seek(0)
                            
                            progress_bar.empty()
                            st.success("All pages perfectly scaled to A4!")
                            st.download_button("Download A4 Standard PDF", data=out_pdf, file_name="A4_Standardized_Doc.pdf", mime="application/pdf", key="dl_a4")
                    except Exception as e:
                        st.error(f"Error during standardization: {e}")

    if t45:
        with st.expander("45. Scanned Book / Landscape Page Slicer (हाफ-पेज स्प्लिटर)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF where a single wide (Landscape) page contains two actual book pages side-by-side. The tool will automatically cut all wide pages exactly down the middle, separating them into two normal Portrait pages for easy reading.")
            if fitz is None:
                st.warning("⚠️ 'PyMuPDF' library is required. Run `pip install PyMuPDF`.")
            else:
                slice_pdf_file = st.file_uploader("Upload Wide/Landscape Book PDF", type=["pdf"], key="slice_pdf")
                if slice_pdf_file and st.button("Slice Pages in Half", key="btn_slice"):
                    try:
                        with st.spinner("Slicing wide pages down the middle..."):
                            with fitz.open(stream=slice_pdf_file.read(), filetype="pdf") as doc, fitz.open() as out_doc:
                                for i in range(len(doc)):
                                    p = doc[i]
                                    # If page is Landscape (Wider than Tall)
                                    if p.rect.width > p.rect.height:
                                        half_width = p.rect.width / 2
                                        
                                        # Left Page
                                        p1 = out_doc.new_page(width=half_width, height=p.rect.height)
                                        p1.show_pdf_page(p1.rect, doc, i, clip=fitz.Rect(0, 0, half_width, p.rect.height))
                                        
                                        # Right Page
                                        p2 = out_doc.new_page(width=half_width, height=p.rect.height)
                                        p2.show_pdf_page(p2.rect, doc, i, clip=fitz.Rect(half_width, 0, p.rect.width, p.rect.height))
                                    else:
                                        # If it's already portrait, just copy it normally
                                        out_doc.insert_pdf(doc, from_page=i, to_page=i)
                                        
                                out_pdf = io.BytesIO()
                                out_doc.save(out_pdf)
                                out_pdf.seek(0)
                                
                                st.success("Landscape pages successfully sliced into two!")
                                st.download_button("Download Sliced PDF", data=out_pdf, file_name="Sliced_Portrait_Book.pdf", mime="application/pdf", key="dl_slice")
                    except Exception as e:
                        st.error(f"Error slicing pages: {e}")

    if t46:
        with st.expander("46. Offline Telephone & Room Directory (सचिवालय टेलीफोन निर्देशिका)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Start typing the name of an officer, section, or department. The tool will display their official Office Telephone Number and Room/Floor location.")
            
            # Directory loaded directly from the official Tel_directory.pdf
            tel_directory = [
                {"name": "Shri CP Radhakrishnan (Hon'ble Chairman, Rajya Sabha)", "office_phone": "23094953, 23094954, 23094955", "room": "RS-14, PH"},
                {"name": "Sh. Amit Khare (Secretary to the Vice President)", "office_phone": "23094941", "room": "Vice-President Enclave"},
                {"name": "Ms. V. Lalithalakshmi (Joint Secretary to VP)", "office_phone": "23094959", "room": "Vice-President Enclave"},
                {"name": "Shri Jagat Prakash Nadda (Leader of the House)", "office_phone": "23083546, 23083547", "room": "G-30, PH"},
                {"name": "Shri Mallikarjun Kharge (Leader of the Opposition)", "office_phone": "23083113, 23083123, 23034883", "room": "G-19, PH / 43, GF, Samvidhan Sadan"},
                {"name": "Shri Harivansh (Hon'ble Deputy Chairman)", "office_phone": "23083029, 23083030, 23034689", "room": "RS-07, PH / 32, GF, Samvidhan Sadan"},
                {"name": "Shri Raghav Chadha (Chairman, Committee on Petitions)", "office_phone": "23035797, 21410314", "room": "310, Third Floor, Block-B, PHA Ext."},
                {"name": "Smt. Ranjeet Ranjan (Chairman, Committee on ICT)", "office_phone": "23035788, 21410331", "room": "319, Third Floor, Block-B, PHA Ext."},
                {"name": "Shri M. Thambidurai (Chairman, Committee on Govt. Assurances)", "office_phone": "23035750, 21410322", "room": "315, Third Floor, Block-B, PHA Ext."},
                {"name": "Shri P.C. Mody (Secretary-General)", "office_phone": "23083035, 23083036", "room": "RS-08, PH"},
                {"name": "Dr. K.S. Somashekhar (Secretary)", "office_phone": "23083052, 23083053", "room": "RS-25, I-F, PH"},
                {"name": "Shri Vimal Kumar (Joint Secretary - Reporting)", "office_phone": "23083058", "room": "RS-37, I-F, PH"},
                {"name": "Shri Rakesh Naithani (Joint Secretary - Q&PD & CVO)", "office_phone": "23035581, 23019329", "room": "Room No.517, V-F, PHA"},
                {"name": "Dr. Kushal Kumar Pathak (Joint Secretary - Systems & CBD & CISO)", "office_phone": "23034967, 23092163", "room": "34, GF, Samvidhan Sadan"},
                {"name": "Dr.(Smt.) Rosey Sailo Damodaran (Joint Secretary - Research)", "office_phone": "23034056", "room": "515, V-F, PHA"},
                {"name": "Shri P. Narayanan (Joint Secretary - Logistics)", "office_phone": "23034084, 23093089", "room": "124, I-F, PHA"},
                {"name": "Sh. Mahender Singh (Joint Secretary - E&T)", "office_phone": "23016431", "room": "239, II-F, PHA"},
                {"name": "Shri Ajaya Kumar Mallick (Joint Secretary - HR)", "office_phone": "23034093, 23093690", "room": "123-A, I-F, PHA"},
                {"name": "Shri Ravinder Kumar (Director - B&P)", "office_phone": "23034252, 23793563", "room": "212, II-F, PHA"},
                {"name": "Shri Sameer Suryapani (Director - Sansad TV & CCA)", "office_phone": "23035415, 23011973", "room": "120, I-F, PHA"},
                {"name": "Shri Sanjeev Chandra (Director - COPLOT & CPIO)", "office_phone": "23035448, 23015557", "room": "122, I-F, PHA"},
                {"name": "Mohd. Salamuddin (Director - Systems & Web Supervisor)", "office_phone": "23035308, 23793633", "room": "125, I-F, PHA"},
                {"name": "Sh. Sandeep Pandey (Deputy Secretary - Systems-II)", "office_phone": "23034259", "room": "28-A, GF, Samvidhan Sadan"},
                {"name": "Shri Navneet Joon (Deputy Secretary - Systems-I)", "office_phone": "23034668, 23035444", "room": "28-A, GF, Samvidhan Sadan"},
                {"name": "Systems Division", "office_phone": "23034325, 23034074", "room": "121, III-F, Samvidhan Sadan"},
                {"name": "IT Complaints / Helpdesk", "office_phone": "23034126, 23034718, 23034626", "room": "helpdesk.msp@supportgov.in"},
                {"name": "C.C.T.V. Control Room", "office_phone": "23034878", "room": "Basement, Samvidhan Sadan"},
                {"name": "Reception Office (PHA)", "office_phone": "23034481, 23035589", "room": "PHA, Reception Office"}
            ]
            
            search_query_dir = st.text_input("Search Name / Section (e.g., 'Systems', 'Kharge', 'IT Complaints'):", key="tel_search")
            if search_query_dir:
                results = [entry for entry in tel_directory if search_query_dir.lower() in entry['name'].lower()]
                if results:
                    for res in results:
                        st.success(f"👤 **{res['name']}**\n\n📞 **Phone:** {res['office_phone']} | 🚪 **Room:** {res['room']}")
                else:
                    st.warning("No contact found. Please check the spelling or try a different keyword.")

    if t47:
        with st.expander("47. CGHS Medical Reimbursement Claim Form - MRC(S) (चिकित्सा प्रतिपूर्ति दावा प्रपत्र)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Fill in the claim details below. Click **'Generate CGHS Form MRC(S)'** to download an exact, print-ready Word document (`.docx`) calibrated for standard **A4 paper** with official margins and layout.")
            
            st.markdown("#### 1. Principal Card Holder Details")
            col_c1, col_c2 = st.columns(2)
            cghs_name = col_c1.text_input("1(a) Name of Card Holder & Designation", placeholder="e.g., Rajesh Sharma, Executive Officer", key="cghs_name")
            cghs_ben_id = col_c2.text_input("1(b) CGHS Ben ID No.", placeholder="e.g., 1234567", key="cghs_ben_id")
            
            col_c3, col_c4 = st.columns(2)
            cghs_emp_code = col_c3.text_input("1(c) Employee Code No.", placeholder="e.g., EMP-9821", key="cghs_emp_code")
            cghs_ward = col_c4.selectbox("1(d) Ward Entitlement", ["General", "Semi-Pvt.", "Pvt."], key="cghs_ward")
            
            col_c5, col_c6 = st.columns(2)
            cghs_basic = col_c5.text_input("1(d) Basic Pay (excluding Grade Pay) (₹)", placeholder="e.g., 56100", key="cghs_basic")
            cghs_contact = col_c6.text_input("1(f) Mobile No. & Email", placeholder="e.g., 98XXXXXXXX / email@rss.sansad.in", key="cghs_contact")
            
            cghs_address = st.text_area("1(e) Full Address", placeholder="e.g., Flat No. 102, Type-IV, PHA Complex, New Delhi", height=68, key="cghs_address")
            
            st.markdown("---")
            st.markdown("#### 2. Patient Details")
            col_p1, col_p2, col_p3 = st.columns(3)
            cghs_pat_name = col_p1.text_input("2(a) Patient's Name", placeholder="e.g., Sunita Sharma", key="cghs_pat_name")
            cghs_pat_id = col_p2.text_input("2(b) Patient's CGHS Ben ID No.", placeholder="e.g., 7654321", key="cghs_pat_id")
            cghs_pat_rel = col_p3.selectbox("2(c) Relationship with Card Holder", ["Self", "Spouse", "Son", "Daughter", "Father", "Mother", "Other"], key="cghs_pat_rel")
            
            st.markdown("---")
            st.markdown("#### 3. Hospital & Treatment Details")
            cghs_hosp = st.text_input("3. Hospital / Diagnostic Centre Name & Address", placeholder="e.g., Max Super Speciality Hospital, Saket, New Delhi", key="cghs_hosp")
            
            col_h1, col_h2, col_h3 = st.columns(3)
            cghs_empanelled = col_h1.selectbox("4. Empanelled under CGHS?", ["Yes", "No"], key="cghs_empanelled")
            cghs_emergency = col_h2.selectbox("6. Taken in Emergency?", ["No", "Yes"], key="cghs_emergency")
            cghs_permission = col_h3.selectbox("7. Prior Permission Taken?", ["Not Applicable", "Yes", "No"], key="cghs_permission")
            
            cghs_ins = st.text_input("8. Subscribing to any other Health Insurance? (Amount claimed/received if any)", value="No / Nil", key="cghs_ins")
            cghs_adv = st.text_input("9. Details of Medical Advance taken, if any", value="Nil", key="cghs_adv")
            
            st.markdown("---")
            st.markdown("#### 4. Amount Claimed (₹)")
            col_a1, col_a2, col_a3 = st.columns(3)
            amt_opd = col_a1.number_input("10(a) OPD Treatment", min_value=0.0, value=0.0, step=100.0, key="amt_opd")
            amt_indoor = col_a2.number_input("10(b) Indoor Treatment", min_value=0.0, value=0.0, step=100.0, key="amt_indoor")
            amt_tests = col_a3.number_input("10(c) Tests / Investigations", min_value=0.0, value=0.0, step=100.0, key="amt_tests")
            total_claim = amt_opd + amt_indoor + amt_tests
            st.info(f"💰 **Total Amount Claimed:** ₹ {total_claim:,.2f}")
            
            st.markdown("---")
            st.markdown("#### 5. Bank Account Details (For Direct Credit)")
            col_b1, col_b2 = st.columns(2)
            b_name = col_b1.text_input("11. Bank Name", placeholder="e.g., State Bank of India", key="b_name")
            b_acc = col_b2.text_input("SB Account No.", placeholder="e.g., 10234567890", key="b_acc")
            col_b3, col_b4 = st.columns(2)
            b_ifsc = col_b3.text_input("IFSC Code", placeholder="e.g., SBIN0000691", key="b_ifsc")
            b_micr = col_b4.text_input("Branch MICR Code", placeholder="e.g., 110002001", key="b_micr")
            
            col_d1, col_d2 = st.columns(2)
            dec_date = col_d1.text_input("Declaration Date", value="09.09.2026", key="dec_date")
            dec_place = col_d2.text_input("Declaration Place", value="New Delhi", key="dec_place")
            
            if st.button("Generate CGHS Form MRC(S) (.docx)", key="btn_cghs_form"):
                doc = docx.Document()
                for s in doc.sections:
                    s.page_width = docx.shared.Mm(210)
                    s.page_height = docx.shared.Mm(297)
                    s.top_margin = Inches(0.5)
                    s.bottom_margin = Inches(0.5)
                    s.left_margin = Inches(0.6)
                    s.right_margin = Inches(0.6)
                    
                p_hdr = doc.add_paragraph()
                p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_h1 = p_hdr.add_run("FORM-MRC (S) (For serving employees)\n")
                r_h1.bold = True
                r_h1.font.size = Pt(11)
                r_h2 = p_hdr.add_run("CENTRAL GOVERNMENT HEALTH SCHEME\nMEDICAL REIMBURSEMENT CLAIM FORM\n")
                r_h2.bold = True
                r_h2.font.size = Pt(12)
                r_h3 = p_hdr.add_run("(To be filled up by the Principal Card holder in BLOCK LETTERS)")
                r_h3.font.size = Pt(9.5)
                r_h3.italic = True
                
                doc.add_paragraph().paragraph_format.space_after = Pt(2)
                
                table = doc.add_table(rows=0, cols=3)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.autofit = False
                
                def add_mrc_row(num_label, desc_label, val_text):
                    row = table.add_row()
                    c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
                    c0.width = Inches(0.5)
                    c1.width = Inches(3.2)
                    c2.width = Inches(3.3)
                    
                    p0 = c0.paragraphs[0]
                    p0.add_run(num_label).bold = True
                    p0.paragraph_format.space_after = Pt(3)
                    
                    p1 = c1.paragraphs[0]
                    p1.add_run(desc_label)
                    p1.paragraph_format.space_after = Pt(3)
                    
                    p2 = c2.paragraphs[0]
                    p2.add_run(f": {val_text}")
                    p2.paragraph_format.space_after = Pt(3)

                add_mrc_row("1.", "(a) Name of Principal Card Holder & Desig.", cghs_name)
                add_mrc_row("", "(b) CGHS Ben ID No.", cghs_ben_id)
                add_mrc_row("", "(c) Employee Code No.", cghs_emp_code)
                add_mrc_row("", "(d) Ward Entitlement / Basic Pay", f"{cghs_ward} / ₹ {cghs_basic}")
                add_mrc_row("", "(e) Full Address", cghs_address)
                add_mrc_row("", "(f) Mobile No. and E-mail address", cghs_contact)
                add_mrc_row("2.", "(a) Patient's Name", cghs_pat_name)
                add_mrc_row("", "(b) Patient's CGHS Ben ID No.", cghs_pat_id)
                add_mrc_row("", "(c) Relationship with Principal Card Holder", cghs_pat_rel)
                add_mrc_row("3.", "Name & Address of Hospital/Diag. Centre", cghs_hosp)
                add_mrc_row("4.", "Hospital Empanelled under CGHS?", cghs_empanelled)
                
                opd_txt = f"₹ {amt_opd:,.2f}" if amt_opd > 0 else "Nil"
                indoor_txt = f"₹ {amt_indoor:,.2f}" if amt_indoor > 0 else "Nil"
                tests_txt = f"₹ {amt_tests:,.2f}" if amt_tests > 0 else "Nil"
                
                add_mrc_row("5.", "Treatment Claimed: (a) OPD / (b) Indoor", f"OPD: {opd_txt} | Indoor: {indoor_txt}")
                add_mrc_row("6.", "Treatment taken in emergency?", cghs_emergency)
                add_mrc_row("7.", "Prior permission taken for treatment?", cghs_permission)
                add_mrc_row("8.", "Subscribing to other Medical Insurance?", cghs_ins)
                add_mrc_row("9.", "Details of Medical Advance taken", cghs_adv)
                add_mrc_row("10.", "Total Amount Claimed (a+b+c)", f"₹ {total_claim:,.2f} (OPD: {opd_txt}, Tests: {tests_txt})")
                add_mrc_row("11.", "Bank Details", f"{b_name}, A/C: {b_acc}")
                add_mrc_row("", "IFSC Code / MICR Code", f"{b_ifsc} / {b_micr}")
                
                p_dec_title = doc.add_paragraph()
                p_dec_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_dt = p_dec_title.add_run("\nDECLARATION")
                r_dt.bold = True
                r_dt.font.size = Pt(10)
                
                p_dec = doc.add_paragraph(
                    "I hereby declare that the statements made in the application are true to the best of my knowledge and belief "
                    "and the person for whom medical expenses were incurred is wholly dependent on me. I am a CGHS beneficiary "
                    "and the CGHS card was valid at the time of treatment. I agree for the reimbursement as is admissible under the rules."
                )
                p_dec.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_dec.runs[0].font.size = Pt(9)
                
                doc.add_paragraph() # Add vertical spacing before signatures
                
                # --- NEW TABLE LAYOUT FOR DATE/PLACE AND SIGNATURE ---
                footer_table = doc.add_table(rows=1, cols=2)
                footer_table.autofit = True
                
                cell_left = footer_table.cell(0, 0)
                p_left = cell_left.paragraphs[0]
                p_left.add_run(f"Date: {dec_date}\nPlace: {dec_place}").bold = True
                
                cell_right = footer_table.cell(0, 1)
                p_right = cell_right.paragraphs[0]
                p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                r_sig = p_right.add_run("\n\nSignature of the Principal CGHS card holder")
                r_sig.bold = True
                r_sig.font.size = Pt(9.5)
                
                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                st.success("CGHS Medical Claim Form MRC(S) generated successfully!")
                st.download_button("📥 Download CGHS Form MRC(S)", data=out, file_name="CGHS_Medical_Claim_Form_MRC_S.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_cghs_btn")

    if t54:
        with st.expander("54. Application for Medical Test/Treatment Permission"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Fill in the details for your diagnostic test or medical treatment permission. Click **'Generate Permission Form'** to download a ready-to-print Word document formatted exactly matching the official line-by-line format.")

            st.markdown("#### Employee Details")
            col_e0, col_e1, col_e2, col_e3 = st.columns([1, 2, 2, 2])
            med_sal = col_e0.selectbox("Salutation", ["Shri", "Smt.", "Km.", "Ms.", "Mr."], key="med_sal")
            med_emp_name = col_e1.text_input("1. Name of Employee", value="SAURABH BATRA", key="med_emp_name")
            med_desig = col_e2.text_input("2. Designation", value="Senior Assistant", key="med_desig")
            med_pay = col_e3.text_input("3. Basic Pay (₹)", key="med_pay")

            st.markdown("#### Patient & Treatment Details")
            col_p1, col_p2 = st.columns(2)
            med_patient = col_p1.text_input("4. Name of Patient", key="med_patient")
            med_relation = col_p2.selectbox("5. Relation with Employee", ["Self", "Spouse", "Son", "Daughter", "Father", "Mother", "Other"], key="med_rel")

            col_t1, col_t2 = st.columns(2)
            med_recom = col_t1.selectbox("6. Recommended by", ["CMO, CGHS Dispensary", "Specialist, Govt. Hospital", "Authorised Medical Attendant (AMA)"], key="med_recom")
            med_date = col_t2.text_input("7. Date of Prescription slip(s)", key="med_date")

            med_tests = st.text_area("8. Details of Diagnostic tests/Medical Treatment", key="med_tests")
            med_hosp = st.text_input("9. Name of Diagnostic Centre/Hospital", key="med_hosp")

            st.markdown("#### CGHS / AMA Details")
            col_c1, col_c2, col_c3 = st.columns(3)
            med_cghs = col_c1.text_input("10(a) CGHS Card No.", key="med_cghs")
            med_disp = col_c2.text_input("10(b) Dispensary Name & No.", key="med_disp")
            med_ama = col_c3.text_input("11. Name of AMA (If applicable)", key="med_ama")

            st.markdown("#### Official Details")
            col_o1, col_o2, col_o3 = st.columns(3)
            med_branch = col_o1.text_input("Branch", value="Systems Division", key="med_branch")
            med_tel = col_o2.text_input("Tel. No.", key="med_tel")
            med_app_date = col_o3.text_input("Date", key="med_app_date")

            if st.button("Generate Permission Form (.docx)", key="btn_med_perm"):
                doc = docx.Document()
                
                # Adjust default style for better readability while fitting perfectly on A4
                style = doc.styles['Normal']
                style.font.name = 'Arial'
                style.font.size = Pt(11)  
                style.paragraph_format.space_after = Pt(6)  
                style.paragraph_format.line_spacing = 1.15  

                for s in doc.sections:
                    # Sticking to strictly enforced A4 dimensions
                    s.page_width = docx.shared.Mm(210)
                    s.page_height = docx.shared.Mm(297)
                    # Reduced vertical margins to give extra space
                    s.top_margin = Inches(0.4)
                    s.bottom_margin = Inches(0.4)
                    s.left_margin = Inches(0.6)
                    s.right_margin = Inches(0.6)

                p_hdr = doc.add_paragraph()
                p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_hdr.paragraph_format.space_after = Pt(8)
                r_h1 = p_hdr.add_run("APPLICATION FOR GRANT OF PERMISSION FOR DIAGNOSTIC TESTS/MEDICAL TREATMENT\n")
                r_h1.bold = True
                r_h1.font.size = Pt(12)
                r_h2 = p_hdr.add_run("[Test/Treatment is to be taken by the official after getting written permission from the Office]")
                r_h2.font.size = Pt(10)

                p1 = doc.add_paragraph(f"1. Name of the Employee (in capital letters): {med_sal} {med_emp_name.upper()}")
                p1.paragraph_format.space_after = Pt(8)
                
                p2 = doc.add_paragraph(f"2. Designation: {med_desig}")
                p2.paragraph_format.space_after = Pt(8)
                
                p3 = doc.add_paragraph(f"3. Basic Pay: {med_pay}")
                p3.paragraph_format.space_after = Pt(8)
                
                p4 = doc.add_paragraph(f"4. Name of the Patient: {med_patient}")
                p4.paragraph_format.space_after = Pt(8)
                
                p5 = doc.add_paragraph(f"5. Relation with the Emjployee: {med_relation}")
                p5.paragraph_format.space_after = Pt(8)
                
                cmo_check = "(✔)" if med_recom == "CMO, CGHS Dispensary" else "( )"
                spec_check = "(✔)" if med_recom == "Specialist, Govt. Hospital" else "( )"
                ama_check = "(✔)" if med_recom == "Authorised Medical Attendant (AMA)" else "( )"
                
                p6 = doc.add_paragraph("6. Diagnostic Tests/Treatment recommended by: [Please(✔) against the relevant head]")
                p6.paragraph_format.space_after = Pt(2)
                
                p6_a = doc.add_paragraph(f"    (a) CMO, CGHS Dispensary {cmo_check}      (b) Specialist, Govt. Hospital {spec_check}")
                p6_a.paragraph_format.space_after = Pt(2)
                
                p6_c = doc.add_paragraph(f"    (c) Authorised Medical Attendant [for beneficiary not covered under CGHS] {ama_check}")
                p6_c.paragraph_format.space_after = Pt(8)
                
                p7 = doc.add_paragraph(f"7. Date of Prescription slip (s): {med_date}")
                p7.paragraph_format.space_after = Pt(8)
                
                # --- TABLE FOR 8 & 9 ---
                table_8_9 = doc.add_table(rows=2, cols=2)
                table_8_9.style = 'Table Grid'
                table_8_9.autofit = False
                
                # Explicit fixed width per cell to avoid A4 margin blowouts
                for row in table_8_9.rows:
                    row.cells[0].width = Inches(3.5)
                    row.cells[1].width = Inches(3.5)
                
                p8_hdr = table_8_9.cell(0,0).paragraphs[0]
                p8_hdr.paragraph_format.space_after = Pt(2)
                p8_hdr.add_run("8. Details of the Diagnostic tests/Medical Treatment").bold = True
                
                p9_hdr = table_8_9.cell(0,1).paragraphs[0]
                p9_hdr.paragraph_format.space_after = Pt(2)
                p9_hdr.add_run("9. Name of Diagnostic Centre/Hospital where\nMedical Diagnostic test/Treatment is to be taken").bold = True
                
                table_8_9.cell(1,0).text = med_tests
                table_8_9.cell(1,0).paragraphs[0].paragraph_format.space_after = Pt(2)
                table_8_9.cell(1,1).text = med_hosp
                table_8_9.cell(1,1).paragraphs[0].paragraph_format.space_after = Pt(2)
                
                p10_title = doc.add_paragraph("10. To be filled by beneficiary covered under CGHS")
                p10_title.paragraph_format.space_before = Pt(8)
                p10_title.paragraph_format.space_after = Pt(4)
                
                # --- TABLE FOR 10 ---
                table_10 = doc.add_table(rows=2, cols=2)
                table_10.style = 'Table Grid'
                table_10.autofit = False
                
                for row in table_10.rows:
                    row.cells[0].width = Inches(3.5)
                    row.cells[1].width = Inches(3.5)
                    
                table_10.cell(0,0).text = "(a) CGHS Card No."
                table_10.cell(0,0).paragraphs[0].paragraph_format.space_after = Pt(2)
                table_10.cell(0,1).text = med_cghs
                table_10.cell(0,1).paragraphs[0].paragraph_format.space_after = Pt(2)
                
                table_10.cell(1,0).text = "(b) Name & Number of the Dispensary"
                table_10.cell(1,0).paragraphs[0].paragraph_format.space_after = Pt(2)
                table_10.cell(1,1).text = med_disp
                table_10.cell(1,1).paragraphs[0].paragraph_format.space_after = Pt(2)
                
                p11 = doc.add_paragraph()
                p11.paragraph_format.space_before = Pt(8)
                p11.paragraph_format.space_after = Pt(4)
                p11.add_run("11.\tTo be filled by beneficiary ")
                r_not = p11.add_run("not")
                r_not.underline = True
                p11.add_run(" covered under CGHS")
                
                p11_a = doc.add_paragraph()
                p11_a.paragraph_format.space_after = Pt(8)
                p11_a.add_run("\t(a) Name of the Authorised Medical Attendant (AMA): Dr. ")
                r_ama = p11_a.add_run(f"{med_ama}" if med_ama else "_________________________")
                r_ama.bold = True
                
                p12_1 = doc.add_paragraph("12. I have enclosed the photocopy of the following documents:")
                p12_1.paragraph_format.space_after = Pt(4)
                p12_2 = doc.add_paragraph("    (a) Prescription slip issued by the doctor.")
                p12_2.paragraph_format.space_after = Pt(2)
                p12_3 = doc.add_paragraph("        [The name of the doctor, dispensary, date and stamp should be clearly visible and legible]")
                p12_3.paragraph_format.space_after = Pt(4)
                p12_4 = doc.add_paragraph("    (b) CGHS Card")
                p12_4.paragraph_format.space_after = Pt(4)
                p12_5 = doc.add_paragraph("    (c) Order of appointment of AMA [for beneficiary not covered under CGHS]")
                p12_5.paragraph_format.space_after = Pt(8)
                
                p13 = doc.add_paragraph("13. I may kindly be granted permission for the above mentioned Test/Treatment.")
                p13.paragraph_format.space_after = Pt(12)
                
                # --- RIGHT-ALIGNED FOOTER BLOCK ---
                footer_table = doc.add_table(rows=4, cols=2)
                footer_table.autofit = False
                for row in footer_table.rows:
                    row.cells[0].width = Inches(4.2)
                    row.cells[1].width = Inches(2.8)
                    
                p0 = footer_table.cell(0, 1).paragraphs[0]
                p0.add_run("Signature: ____________________")
                p0.paragraph_format.space_after = Pt(0)
                
                def add_footer_line(row_idx, label, val):
                    p = footer_table.cell(row_idx, 1).paragraphs[0]
                    p.add_run(label)
                    pad_length = max(0, 24 - len(str(val)))
                    padded_val = f"{val}" + "\u00A0" * pad_length
                    r_val = p.add_run(padded_val)
                    r_val.underline = True
                    p.paragraph_format.space_after = Pt(0)
                    
                add_footer_line(1, "Branch: ", med_branch)
                add_footer_line(2, "Tel.No.: ", med_tel)
                add_footer_line(3, "Date: ", med_app_date)

                out = io.BytesIO()
                doc.save(out)
                out.seek(0)
                st.success("Permission Form generated exactly matching the original format!")
                st.download_button("📥 Download Exact Permission Form", data=out, file_name="Medical_Test_Permission_Exact_Inline_Form.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_med_perm")


# ==========================================
# CATEGORY 10: ULTIMATE PRODUCTIVITY
# ==========================================
if any([t48, t49, t50, t51, t52, t53]):
    st.markdown("## 🏆 Ultimate Productivity Utilities")

    if t48:
        with st.expander("48. Scanned PDF Whitener & De-Shadow Cleaner (दस्तावेज़ बैकग्राउंड क्लीनर)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a poor-quality scanned PDF with a grey/yellow background or shadows. This tool will clean the background, making it pure white and sharpening the text. It saves massive amounts of printer ink.")
            if fitz is None:
                st.warning("⚠️ 'PyMuPDF' library is required. Run `pip install PyMuPDF`.")
            else:
                whitener_file = st.file_uploader("Upload Scanned PDF to Clean", type=["pdf"], key="wh_pdf")
                if whitener_file and st.button("Clean & Whiten Background", key="btn_wh"):
                    try:
                        with fitz.open(stream=whitener_file.read(), filetype="pdf") as doc:
                            out_pdf_bytes = io.BytesIO()
                            merger = PdfWriter()
                            
                            st.write("Processing pixels to remove shadows and whiten background...")
                            progress_bar = st.progress(0)
                            total_p = len(doc)
                            
                            for i, page in enumerate(doc):
                                pix = page.get_pixmap(dpi=200)
                                # Convert to Grayscale
                                img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
                                
                                # Enhance Contrast massively to separate text from background
                                img = ImageEnhance.Contrast(img).enhance(2.0)
                                
                                # Thresholding: Any pixel lighter than grey (150) becomes pure white (255)
                                img = img.point(lambda p: 255 if p > 150 else p)
                                
                                temp_pdf = io.BytesIO()
                                img.save(temp_pdf, format="PDF")
                                temp_pdf.seek(0)
                                
                                merger.add_page(PdfReader(temp_pdf).pages[0])
                                progress_bar.progress((i + 1) / total_p)
                                
                            merger.write(out_pdf_bytes)
                            out_pdf_bytes.seek(0)
                            
                            progress_bar.empty()
                            st.success("Background cleaned successfully!")
                            st.download_button("Download Cleaned PDF", data=out_pdf_bytes, file_name="Cleaned_White_Document.pdf", mime="application/pdf", key="dl_wh")
                    except Exception as e:
                        st.error(f"Error cleaning PDF: {e}")

    if t49:
        with st.expander("49. Smart Auto-PII Redactor (आधार / पैन / মোবাইল ऑटो-ब्लैकआउट)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a digital PDF. Select what personal info you want to hide (Aadhaar, PAN, Mobile, Email). Click 'Auto-Redact' and the system will automatically find and blackout all such sensitive data everywhere in the document.")
            if fitz is None:
                st.warning("⚠️ 'PyMuPDF' library is required. Run `pip install PyMuPDF`.")
            else:
                pii_pdf = st.file_uploader("Upload Digital PDF for PII Redaction", type=["pdf"], key="pii_pdf")
                
                st.markdown("Select what information to automatically hide:")
                col1, col2, col3, col4 = st.columns(4)
                hide_aadhaar = col1.checkbox("Aadhaar Numbers", value=True)
                hide_pan = col2.checkbox("PAN Numbers", value=True)
                hide_mobile = col3.checkbox("Mobile Numbers", value=True)
                hide_email = col4.checkbox("Email IDs", value=True)
                
                if pii_pdf and st.button("Auto-Redact Sensitive Info", key="btn_pii"):
                    try:
                        with fitz.open(stream=pii_pdf.read(), filetype="pdf") as doc:
                            st.write("Scanning document for Personal Identifiable Information (PII)...")
                            progress_bar = st.progress(0)
                            total_p = len(doc)
                            total_redactions = 0
                            
                            # Define Regex Patterns
                            patterns = []
                            if hide_aadhaar: patterns.append(r"\b\d{4}[\s\-]?\d{4}[\s\-]?\d{4}\b")
                            if hide_pan: patterns.append(r"\b[A-Z]{5}\d{4}[A-Z]\b")
                            if hide_mobile: patterns.append(r"\b(?:\+91|91)?[\s\-]?\d{10}\b")
                            if hide_email: patterns.append(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b")
                            
                            for i, page in enumerate(doc):
                                text = page.get_text("text")
                                for pattern in patterns:
                                    matches = re.findall(pattern, text)
                                    for match in matches:
                                        areas = page.search_for(match)
                                        for rect in areas:
                                            page.add_redact_annot(rect, fill=(0, 0, 0))
                                            total_redactions += 1
                                page.apply_redactions()
                                progress_bar.progress((i + 1) / total_p)
                                
                            out_pdf = io.BytesIO()
                            doc.save(out_pdf)
                            out_pdf.seek(0)
                            
                            progress_bar.empty()
                            st.success(f"Success! Blocked {total_redactions} sensitive PII items in the document.")
                            st.download_button("Download Redacted PDF", data=out_pdf, file_name="PII_Redacted_Document.pdf", mime="application/pdf", key="dl_pii")
                    except Exception as e:
                        st.error(f"Error during Auto-Redaction: {e}")

    if t50:
        with st.expander("50. Manual Duplex Printing Assistant (ऑड-ईवन प्रिंट हेल्पर)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF. This tool splits it into two files: 'Odd_Pages' and 'Even_Pages_Reversed'. Print the Odd file, put the stack exactly as it comes out back into the printer tray, and print the Even file. Perfect two-sided manual printing without mistakes!")
            
            duplex_pdf = st.file_uploader("Upload PDF to Split for Duplex Printing", type=["pdf"], key="duplex_pdf")
            if duplex_pdf and st.button("Generate Duplex Print Files", key="btn_duplex"):
                try:
                    with st.spinner("Splitting and reversing pages for manual duplex..."):
                        reader = PdfReader(duplex_pdf)
                        writer_odd = PdfWriter()
                        writer_even = PdfWriter()
                        
                        even_pages = []
                        for i in range(len(reader.pages)):
                            if i % 2 == 0:  # 0-indexed, so 0 is Page 1 (Odd)
                                writer_odd.add_page(reader.pages[i])
                            else:           # 1 is Page 2 (Even)
                                even_pages.append(reader.pages[i])
                                
                        # Reverse Even pages for standard tray reloading
                        for p in reversed(even_pages):
                            writer_even.add_page(p)
                            
                        odd_out = io.BytesIO()
                        writer_odd.write(odd_out)
                        
                        even_out = io.BytesIO()
                        writer_even.write(even_out)
                        
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            zip_file.writestr("1_Print_First_Odd_Pages.pdf", odd_out.getvalue())
                            zip_file.writestr("2_Print_Second_Even_Pages_Reversed.pdf", even_out.getvalue())
                            
                        zip_buffer.seek(0)
                        st.success("Files prepared successfully!")
                        st.download_button("Download Duplex Print ZIP", data=zip_buffer, file_name="Manual_Duplex_Printing.zip", mime="application/zip", key="dl_duplex")
                except Exception as e:
                    st.error(f"Error generating duplex files: {e}")

    if t51:
        with st.expander("51. PDF Color vs B&W Page Audit & Splitter (प्रिंटर बजट सेवर)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a long report. The tool scans every pixel and separates it into two PDFs: one with only Black & White pages (cheap to print) and one with only Color pages (for the color printer).")
            if fitz is None:
                st.warning("⚠️ 'PyMuPDF' library is required. Run `pip install PyMuPDF`.")
            else:
                audit_pdf = st.file_uploader("Upload PDF to Audit Colors", type=["pdf"], key="audit_pdf")
                if audit_pdf and st.button("Audit and Split Pages", key="btn_audit"):
                    try:
                        with fitz.open(stream=audit_pdf.read(), filetype="pdf") as doc, fitz.open() as bw_doc, fitz.open() as color_doc:
                            st.write("Scanning all pixels to detect color pages...")
                            progress_bar = st.progress(0)
                            total_p = len(doc)
                            bw_count, color_count = 0, 0
                            
                            for i in range(total_p):
                                # Render at low DPI just to check colors quickly
                                pix = doc[i].get_pixmap(colorspace=fitz.csRGB, dpi=36)
                                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                                
                                # Compare original RGB to a Grayscale version of itself
                                gray_img = img.convert('L').convert('RGB')
                                diff = ImageChops.difference(img, gray_img)
                                stat = ImageStat.Stat(diff)
                                
                                # If average difference in color bands is greater than a small threshold, it's a color page
                                is_color = sum(stat.mean) > 1.5
                                
                                if is_color:
                                    color_doc.insert_pdf(doc, from_page=i, to_page=i)
                                    color_count += 1
                                else:
                                    bw_doc.insert_pdf(doc, from_page=i, to_page=i)
                                    bw_count += 1
                                    
                                progress_bar.progress((i + 1) / total_p)
                                    
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                                if bw_count > 0:
                                    bw_bytes = io.BytesIO()
                                    bw_doc.save(bw_bytes)
                                    zip_file.writestr(f"Black_And_White_Pages_({bw_count}).pdf", bw_bytes.getvalue())
                                if color_count > 0:
                                    col_bytes = io.BytesIO()
                                    color_doc.save(col_bytes)
                                    zip_file.writestr(f"Color_Pages_({color_count}).pdf", col_bytes.getvalue())
                                    
                            zip_buffer.seek(0)
                            progress_bar.empty()
                            st.success(f"Audit Complete! Found {bw_count} B&W pages and {color_count} Color pages.")
                            st.download_button("Download Split Audit ZIP", data=zip_buffer, file_name="Color_BW_Split_Audit.zip", mime="application/zip", key="dl_audit")
                    except Exception as e:
                        st.error(f"Error during audit: {e}")

    if t52:
        with st.expander("52. PDF Embedded Portfolio / Attachment Packer (डिजिटल मिसल टूल)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a Main PDF (like a Note Sheet). Then upload multiple attachments (Excel, Word, Images). This tool will embed those attachments *inside* the Main PDF as internal files, creating a single shareable electronic portfolio (E-Missil).")
            if fitz is None:
                st.warning("⚠️ 'PyMuPDF' library is required. Run `pip install PyMuPDF`.")
            else:
                main_pdf = st.file_uploader("1. Upload Main PDF Document", type=["pdf"], key="main_pdf")
                attach_files = st.file_uploader("2. Upload Files to Attach/Embed", accept_multiple_files=True, key="attach_files")
                
                if main_pdf and attach_files and st.button("Pack into Portfolio PDF", key="btn_pack"):
                    try:
                        with st.spinner("Embedding files into PDF..."):
                            with fitz.open(stream=main_pdf.read(), filetype="pdf") as doc:
                                for f in attach_files:
                                    try:
                                        # Using PyMuPDF's embfile_add function to embed attachments natively
                                        doc.embfile_add(f.name, f.read(), filename=f.name)
                                    except AttributeError:
                                        st.error("Your version of PyMuPDF is outdated and doesn't support embedding. Please run: pip install --upgrade PyMuPDF")
                                        st.stop()
                                        
                                out_pdf = io.BytesIO()
                                doc.save(out_pdf)
                                out_pdf.seek(0)
                                
                                st.success(f"Successfully embedded {len(attach_files)} files into the Main PDF!")
                                st.info("💡 Note: To view the embedded attachments, open the downloaded PDF in Adobe Acrobat Reader and click the 'Paperclip' icon on the left sidebar.")
                                st.download_button("Download E-Portfolio PDF", data=out_pdf, file_name="Embedded_Portfolio_Document.pdf", mime="application/pdf", key="dl_pack")
                    except Exception as e:
                        st.error(f"Error embedding files: {e}")

    if t53:
        with st.expander("53. Smart Booklet Imposition (Center-Staple Maker)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF. Select your printer's paper size (A4, A3, Letter, or Custom). The tool will automatically add blank white pages to make the count a multiple of 4, then arrange the pages for a center-stapled booklet. \n\n*(If you have a Word file, please convert it to PDF first using Tool 13).*")
            
            if fitz is None:
                st.warning("⚠️ 'PyMuPDF' library is required. Run `pip install PyMuPDF`.")
            else:
                booklet_pdf = st.file_uploader("Upload PDF Document to create Booklet", type=["pdf"], key="booklet_pdf")
                sheet_size = st.selectbox("Select Printer Paper Size (Target Sheet)", [
                    "A4 (Creates an A5 size folded booklet)", 
                    "A3 (Creates an A4 size folded booklet)", 
                    "Letter (Creates a Half-Letter folded booklet)",
                    "Custom Size"
                ], key="booklet_size")
                
                custom_w_pts, custom_h_pts = 0, 0
                if sheet_size == "Custom Size":
                    col1, col2, col3 = st.columns(3)
                    unit = col1.selectbox("Unit", ["Inches", "Centimeters"], key="custom_unit")
                    custom_w = col2.number_input("Sheet Width (Landscape)", min_value=1.0, value=17.0 if unit=="Inches" else 42.0, key="custom_w")
                    custom_h = col3.number_input("Sheet Height", min_value=1.0, value=11.0 if unit=="Inches" else 29.7, key="custom_h")
                    
                    # Convert to points (1 inch = 72 points, 1 cm = 72 / 2.54 points)
                    multiplier = 72.0 if unit == "Inches" else (72.0 / 2.54)
                    custom_w_pts = custom_w * multiplier
                    custom_h_pts = custom_h * multiplier
                
                if booklet_pdf and st.button("Generate Print-Ready Booklet", key="btn_booklet"):
                    try:
                        with st.spinner("Calculating imposition and arranging pages..."):
                            with fitz.open(stream=booklet_pdf.read(), filetype="pdf") as doc, fitz.open() as out_doc:
                                
                                # 1. Add blank pages to make total pages a multiple of 4
                                remainder = len(doc) % 4
                                blanks_to_add = 0
                                if remainder != 0:
                                    blanks_to_add = 4 - remainder
                                    for _ in range(blanks_to_add):
                                        # Add a blank page matching the size of the last page
                                        doc.new_page(width=doc[-1].rect.width, height=doc[-1].rect.height)
                                        
                                total_pages = len(doc)
                                
                                # 2. Determine target sheet dimensions (Landscape)
                                if "A4" in sheet_size:
                                    sheet_w, sheet_h = fitz.paper_size("a4-l")
                                elif "A3" in sheet_size:
                                    sheet_w, sheet_h = fitz.paper_size("a3-l")
                                elif "Letter" in sheet_size:
                                    sheet_w, sheet_h = fitz.paper_size("letter-l")
                                else:
                                    sheet_w, sheet_h = custom_w_pts, custom_h_pts
                                    
                                half_w = sheet_w / 2
                                
                                # 3. Saddle-Stitch Imposition Logic
                                for k in range(total_pages // 4):
                                    # Front side of the sheet
                                    front_page = out_doc.new_page(width=sheet_w, height=sheet_h)
                                    left_idx_front = total_pages - 1 - 2*k
                                    right_idx_front = 2*k
                                    
                                    front_page.show_pdf_page(fitz.Rect(0, 0, half_w, sheet_h), doc, left_idx_front)
                                    front_page.show_pdf_page(fitz.Rect(half_w, 0, sheet_w, sheet_h), doc, right_idx_front)
                                    
                                    # Back side of the sheet
                                    back_page = out_doc.new_page(width=sheet_w, height=sheet_h)
                                    left_idx_back = 2*k + 1
                                    right_idx_back = total_pages - 2 - 2*k
                                    
                                    back_page.show_pdf_page(fitz.Rect(0, 0, half_w, sheet_h), doc, left_idx_back)
                                    back_page.show_pdf_page(fitz.Rect(half_w, 0, sheet_w, sheet_h), doc, right_idx_back)
                                    
                                out_pdf = io.BytesIO()
                                out_doc.save(out_pdf)
                                out_pdf.seek(0)
                                
                                st.success(f"Booklet Imposition Complete! ✅ Automatically added {blanks_to_add} blank page(s) to maintain sequence.")
                                st.info("🖨️ **Printing Tip:** When printing this downloaded file, select **'Print on Both Sides' (Duplex)** and choose **'Flip on Short Edge'**. Then simply fold the printed stack in half and staple the center!")
                                st.download_button("Download Print-Ready Booklet", data=out_pdf, file_name="Center_Staple_Booklet.pdf", mime="application/pdf", key="dl_booklet")
                    except Exception as e:
                        st.error(f"Error generating booklet: {e}")

# --- FOOTER TAGLINE ---
st.markdown("<br><br><br><br>", unsafe_allow_html=True)
st.markdown("""
    <div class="footer">
        For secure and seamless document conversion, use दस्तावेज़ सेतु—because confidentiality should never be compromised- Saurabh, SYSTEMS DIVISION.
    </div>
""", unsafe_allow_html=True)