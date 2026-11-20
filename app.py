import io
import re
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- Formatting Constants ---
BOLD_TERMS = [
    r"diperakui dengan terma atau syarat",
    r"diperakui",
    r"tidak dapat dipertimbangkan",
    r"kepada"
]

CUSTOM_ENGLISH_WORDS = {
    "accelerator", "accident", "activities", "acuity", "adl", "admin", "administration", "administrative", "administrator", "adolescent", "affiliated",
    "all", "ambulatory", "amp", "anaesthesia", "anaesthesiology", "analog", "and", "angio", "angiografi", "area", "assessment", "association", "asthma", "at",
    "audiologist", "audiometric", "audiometry", "autoclave", "available", "avenue", "bank", "barrier", "based", "basement", "bassinet", "bassinets", "bay",
    "bays", "beam", "bed", "bedded", "bereavement", "billing", "bin", "block", "blood", "bobath", "body", "bone", "booth", "brachytherapy", "breastfeeding",
    "bridge", "bronchoscopy", "bulk", "bunker", "business", "c-arm", "cabinet", "cafeteria", "cardiac", "care", "cashier", "catheterization", "centre", "chair",
    "change", "changing", "chemo", "chemotherapy", "child", "children", "city", "clean", "cleaning", "clerk", "clinic", "clinical", "clinics", "closet", "cold",
    "collection", "colonoscopy", "community", "complex", "cone", "consultation", "consumable", "control", "corridor", "couch", "counselling", "counter",
    "counterpulsation", "crrt", "cssd", "cssu", "cycling", "cytotoxic", "day", "daycare", "dayward", "delivery", "densitometer", "densitometry", "dental",
    "department", "dermatology", "desk", "development", "developmental", "diagnostic", "dialysis", "dietitian", "digital", "dining", "dirty", "discharge",
    "discussion", "diseases", "dispensary", "disposal", "district", "doctor", "doctor's", "domestic", "donor", "double", "double-bedded", "drug", "east", "ecg", "echo",
    "eco", "education", "eeg", "eighth", "electrical", "electrotherapy", "elevator", "emergency", "endoscopy", "enhanced", "entrance", "equipment", "exam", "examination",
    "exec", "executive", "existing", "extended", "extension", "external", "eye", "facilities", "facillities", "family", "female", "fertility", "fetal", "fifth", "file",
    "filing", "first", "floor", "focused", "food", "for", "four", "four-bedded", "fourth", "function", "gamma", "gas", "gastroenterology", "gateway", "general", "grievance",
    "grill", "grilled", "ground", "gym", "gymnasium", "gynaecology", "haemodialysis", "hall", "hand", "hdu", "head", "health", "high", "hold", "holding", "home", "homemaker",
    "hospice", "hospital", "hot", "hybrid", "icu", "imaging", "in", "infectious", "inhouse", "inpatient", "intensity", "intensive", "intermediate", "international", "intra",
    "intraoperative", "isolation", "it", "janitor", "janitor's", "joint", "kitchen", "knife", "lab", "laboratory", "labour", "laser", "ldr", "level", "licensed", "lift", "linear",
    "linen", "link", "lobby", "local", "locker", "lounge", "lower", "lung", "machine", "machines", "major", "male", "mall", "mammogram", "mammography", "management", "manager",
    "maternal", "maternity", "matron", "maxillofacial", "medical", "medication", "medicine", "meeting", "mezzanine", "mineral", "minor", "mobile", "multidisiplin", "multipurpose",
    "negative", "nephrologist", "nephrology", "nerve", "neurology", "new", "ninth", "non-hepatitis", "nurse", "nursery", "nursing", "observation", "obstetrics", "occupational", "of",
    "off", "office", "officers", "oncology", "open", "operating", "operation", "ophthalmology", "optometrist", "optometry", "oral", "orthopaedic", "outpatient", "outsourced", "over",
    "packing", "paediatric", "paediatrician", "paediatrics", "paeds", "palliative", "pantry", "paraffin", "park", "pass", "patient", "pediatric", "pendant", "pharmacy", "phlebotomy",
    "physical", "physician's", "physio", "physiotherapy", "plastic", "play", "portable", "positive", "pre-function", "pre-operating", "prep", "preparation", "pressure", "private",
    "procedure", "processing", "psychiatry", "public", "quarantine", "radiation", "radiography", "radiologist", "radiology", "radiotherapy", "readily", "receiving", "reception",
    "reconstitution", "record", "recovery", "refraction", "refuge", "registration", "rehab", "rehabilitation", "relokasi", "reporting", "reprocessor", "reproductive", "respiratory",
    "respite", "resusitasi", "retail", "robotic", "room", "rooming-in", "rooms", "sample", "satellite", "scanner", "scientific", "scn", "scope", "screen", "screening", "second",
    "sensory", "serum", "server", "services", "settlement", "seventh", "silent", "simulator", "single", "single-bedded", "sister", "sixth", "sleep", "sluice", "society", "soiled",
    "special", "specialist", "specialists", "speech", "splinting", "sports", "staff", "stair", "station", "sterilization", "storage", "store", "stress", "stretcher", "study", "sub",
    "suite", "supervisor", "supply", "support", "surgery", "surgical", "table", "taking", "tea", "teaching", "technical", "tenth", "test", "theatre", "therapy", "third", "three",
    "three-bedded", "times", "to", "toilet", "topical", "tower", "traction", "training", "treadmill", "treadmills", "treatment", "triage", "trolley", "ultrasound", "unit", "uro-gynaecology",
    "uroflowmetry", "utility", "vacant", "vascular", "vertigo", "vestibular", "vision", "visiting", "visitor", "visual", "volunteer", "wait", "waiting", "ward", "waste", "wax", "wheelchair",
    "wing", "with", "women's", "work", "x-ray", "airlock", "auditory", "autologous", "baby", "basis", "brainstem", "breast", "cancers", "cell", "chamber", "clinically", "committee", "disabled",
    "electrophysiology", "endocrine", "enhancement", "ethical", "execute", "formula", "green", "hearing", "hyperbaric", "immune", "institutional", "main", "marketing", "multiplace", "nourishment",
    "only", "oxygen", "policies", "procedures", "progress", "protocol", "proven", "psychiatric", "ramp", "red", "response", "standard", "stem", "therapeutic", "trial", "value", "wardrobe", "yellow", "zone"
}

# --- Extractor ---
def extract_field_data(doc):
    """Extracts fields using positional table coordinates and tracks dynamic XML numbering."""
    data = {
        "nama_premis": "",
        "alamat_premis": "",
        "no_fail": "",
        "pegawai_proses": "",
        "isu_ulasan": ""
    }

    # --- Step 1: Build safe sequential text for Regex fields (No Fail & Pegawai) ---
    sequential_lines = []
    for element in doc.element.body.iterchildren():
        if isinstance(element, CT_P):
            p = Paragraph(element, doc)
            if p.text.strip():
                sequential_lines.append(p.text.strip())
        elif isinstance(element, CT_Tbl):
            table = Table(element, doc)
            for row in table.rows:
                row_cells = []
                seen_cells = set()
                for cell in row.cells:
                    if cell._tc not in seen_cells:
                        seen_cells.add(cell._tc)
                        if cell.text.strip():
                            row_cells.append(cell.text.strip())
                if row_cells:
                    sequential_lines.append(" | ".join(row_cells))
                    
    full_document_text = "\n".join(sequential_lines)

    # 2. No. Fail
    fail_match = re.search(r'(KKM[\s/]*87/A2/[A-Za-z0-9/.-]+(?:[\s]*JILID[\s]*\d+)?)', full_document_text, re.IGNORECASE)
    if fail_match:
        data["no_fail"] = fail_match.group(1).strip()

    # 3. Pegawai Proses
    pegawai_match = re.search(r'Pegawai\s*Proses[\s:\|]*([^\n\r\|]+)', full_document_text, re.IGNORECASE)
    if pegawai_match:
        data["pegawai_proses"] = pegawai_match.group(1).strip()

    # --- Step 2: Positional Extraction for Table Data ---
    for table in doc.tables:
        if len(table.rows) >= 2:
            row_2_cells = []
            seen_cells = set()
            for cell in table.rows[1].cells:
                if cell._tc not in seen_cells:
                    seen_cells.add(cell._tc)
                    row_2_cells.append(cell)
            
            if len(row_2_cells) >= 4:
                # --- 1. Nama Premis & Alamat Premis ---
                cell_2_2_text = row_2_cells[1].text.strip()
                premis_match = re.search(r'([\s\S]*?)(?=Pemohon:)', cell_2_2_text, re.IGNORECASE)
                raw_premis_alamat = premis_match.group(1).strip() if premis_match else cell_2_2_text
                
                lines = [line.strip() for line in raw_premis_alamat.split('\n') if line.strip()]
                if lines:
                    data["nama_premis"] = lines[0]
                    data["alamat_premis"] = " ".join(lines[1:])
                
                # --- 4. Isu / Ulasan (With Dynamic Number/Bullet Tracking) ---
                # The details that always need numbering are in the 4th column (index 3).
                cell_2_4 = row_2_cells[3]
                cell_text_parts = []
                
                list_counters = {}
                last_ilvl = -1
                
                for p in cell_2_4.paragraphs:
                    text = p.text.strip()
                    if not text:
                        continue
                        
                    is_bullet = False
                    is_number = False
                    ilvl = 0
                    
                    style_name = p.style.name if p.style else ""
                    if 'Bullet' in style_name:
                        is_bullet = True
                        
                    # Check XML for native MS Word list formatting properties
                    num_pr = p._element.xpath('.//w:numPr')
                    if num_pr:
                        ilvl_nodes = p._element.xpath('.//w:ilvl/@w:val')
                        if ilvl_nodes:
                            ilvl = int(ilvl_nodes[0])
                        if not is_bullet:
                            is_number = True
                    elif 'Number' in style_name or 'List' in style_name:
                        if not is_bullet:
                            is_number = True
                        if p.paragraph_format.left_indent:
                            # Estimate indent level if not explicitly in XML
                            ilvl = int(p.paragraph_format.left_indent.pt // 18)

                    indent_space = "    " * ilvl
                    
                    # Safety check: if text is manually typed as a list (e.g., "1. ", "a) ") avoid double numbering
                    if re.match(r'^([a-zA-Z0-9]+[\.\)]|[-•])\s', text):
                        # Reset tracking if user manually forced a new main numbered list
                        if re.match(r'^1[\.\)]\s', text):
                            list_counters = {}
                        cell_text_parts.append(f"{indent_space}{text}")
                        continue
                        
                    if is_bullet:
                        cell_text_parts.append(f"{indent_space}• {text}")
                    elif is_number:
                        # Reset deeper levels if we returned to a higher list level
                        if ilvl < last_ilvl:
                            for level in list(list_counters.keys()):
                                if level > ilvl:
                                    list_counters[level] = 0
                                    
                        # Increment current list level counter
                        list_counters[ilvl] = list_counters.get(ilvl, 0) + 1
                        last_ilvl = ilvl
                        val = list_counters[ilvl]
                        
                        # Apply appropriate numbering format based on indentation depth
                        if ilvl == 0:
                            prefix = f"{val}."
                        elif ilvl == 1:
                            prefix = f"{chr(96 + val) if val <= 26 else val}."
                        elif ilvl == 2:
                            roman = ["", "i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x"]
                            prefix = f"{roman[val]}." if val <= 10 else f"{val}."
                        else:
                            prefix = f"{val}."
                            
                        cell_text_parts.append(f"{indent_space}{prefix} {text}")
                    else:
                        cell_text_parts.append(text)
                            
                cell_2_4_text = "\n".join(cell_text_parts)
                
                # Regex capture block
                isu_match = re.search(
                    r'(Permohonan[\s\S]*?tarikh pindaan bermula[^\n\r]*)', 
                    cell_2_4_text, 
                    re.IGNORECASE
                )
                
                if isu_match:
                    data["isu_ulasan"] = isu_match.group(1).strip()
                else:
                    data["isu_ulasan"] = cell_2_4_text
                
                break 

    return data


# --- Builders & Formatters ---
def format_styled_text_run(paragraph, text):
    """Formats right-column text with Arial 9pt, selective bolding, and italics for English words."""
    if not text:
        return

    bold_pattern = f"({'|'.join(BOLD_TERMS)})"
    tokens = re.split(bold_pattern, text, flags=re.IGNORECASE)

    for token in tokens:
        if not token:
            continue
            
        is_bold = any(re.fullmatch(term, token, re.IGNORECASE) for term in BOLD_TERMS)
        sub_tokens = re.split(r'(\s+|[.,;:]+)', token)
        
        for sub in sub_tokens:
            if not sub:
                continue
            
            run = paragraph.add_run(sub)
            run.font.name = 'Arial'
            run.font.size = Pt(9)
            
            if is_bold:
                run.bold = True
            
            # Preserve hyphens and apostrophes during the lookup clean up
            clean_word = re.sub(r'[^\w\-\']', '', sub).lower()
            if clean_word in CUSTOM_ENGLISH_WORDS and not is_bold:
                run.italic = True


def build_summary_doc(records):
    """Creates in-memory Word Document formatted with Arial 9pt, selective bold labels, and styling."""
    doc = Document()
    
    # Title Section
    title_p = doc.add_paragraph()
    title_run = title_p.add_run('Laporan Semakan Permohonan - Summarized Table')
    title_run.font.name = 'Arial'
    title_run.font.size = Pt(12)
    title_run.bold = True

    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    
    # Header Row
    hdr_cells = table.rows[0].cells
    hdr_titles = ["Maklumat Premis & Pegawai", "Isu / Ulasan"]
    
    for i, title in enumerate(hdr_titles):
        p = hdr_cells[i].paragraphs[0]
        run = p.add_run(title)
        run.font.name = 'Arial'
        run.font.size = Pt(9)
        run.bold = True

    # Populate Data
    for data in records:
        row_cells = table.add_row().cells
        
        # --- Left Column ---
        p_left = row_cells[0].paragraphs[0]
        
        val_premis = p_left.add_run(f"{data.get('nama_premis', '')}\n")
        val_premis.font.name = 'Arial'
        val_premis.font.size = Pt(9)
        val_premis.bold = True

        val_alamat = p_left.add_run(f"{data.get('alamat_premis', '')}\n")
        val_alamat.font.name = 'Arial'
        val_alamat.font.size = Pt(9)
        val_alamat.bold = False

        lbl_fail = p_left.add_run("No. Fail: ")
        lbl_fail.font.name = 'Arial'
        lbl_fail.font.size = Pt(9)
        lbl_fail.bold = True
        
        val_fail = p_left.add_run(f"{data.get('no_fail', '')}\n")
        val_fail.font.name = 'Arial'
        val_fail.font.size = Pt(9)
        val_fail.bold = False

        lbl_pegawai = p_left.add_run("Pegawai Proses: ")
        lbl_pegawai.font.name = 'Arial'
        lbl_pegawai.font.size = Pt(9)
        lbl_pegawai.bold = True
        
        val_pegawai = p_left.add_run(f"{data.get('pegawai_proses', '')}")
        val_pegawai.font.name = 'Arial'
        val_pegawai.font.size = Pt(9)
        val_pegawai.bold = False

        # --- Right Column ---
        right_text = data.get('isu_ulasan', '')
        
        right_text = re.sub(r'\s*(Tarikh pindaan)', r'\n\n\1', right_text, flags=re.IGNORECASE)
        lines = right_text.split('\n')
        
        first_paragraph_used = False
        for line in lines:
            if not line.strip():
                if first_paragraph_used:
                    row_cells[1].add_paragraph()
                continue
            
            leading_spaces = len(line) - len(line.lstrip())
            indent_level = leading_spaces // 4
            clean_text = line.lstrip() 
            
            # Check for list prefix (e.g., "1. ", "a. ", "i. ", "• ", "- ")
            list_match = re.match(r'^([a-zA-Z0-9]+[\.\)]|[•\-])\s', clean_text)
            
            if list_match:
                bullet_str = list_match.group(1)
                content_str = clean_text[len(list_match.group(0)):]
                
                # Create a borderless nested table for perfect list alignment
                nested_table = row_cells[1].add_table(rows=1, cols=2)
                nested_table.autofit = False
                
                cell_bullet = nested_table.cell(0, 0)
                cell_content = nested_table.cell(0, 1)
                
                # Force widths to keep bullet compact and left-aligned
                cell_bullet.width = Pt(25 + (15 * indent_level))
                
                # Format the Number/Bullet cell
                p_bullet = cell_bullet.paragraphs[0]
                p_bullet.paragraph_format.left_indent = Pt(15 * indent_level)
                run_bullet = p_bullet.add_run(bullet_str)
                run_bullet.font.name = 'Arial'
                run_bullet.font.size = Pt(9)
                
                # Format the Content cell
                p_content = cell_content.paragraphs[0]
                p_content.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                format_styled_text_run(p_content, content_str)
                
                first_paragraph_used = True
            else:
                if not first_paragraph_used:
                    p_right = row_cells[1].paragraphs[0]
                    first_paragraph_used = True
                else:
                    p_right = row_cells[1].add_paragraph()
                
                p_right.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                if indent_level > 0:
                    p_right.paragraph_format.left_indent = Pt(18 * indent_level)
                
                format_styled_text_run(p_right, clean_text)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# --- Streamlit UI ---
st.set_page_config(page_title="Document Extractor", page_icon="📄", layout="centered")

st.title("📄 Pindaan to Lampiran 2 Extractor")
st.write("Upload your Word documents (.docx) below to extract data into a 2-column Word table.")

uploaded_files = st.file_uploader(
    "Choose Word files (.docx only)",
    type=["docx"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("Process Documents", type="primary"):
        with st.spinner("Extracting data from files..."):
            extracted_records = []
            
            for file in uploaded_files:
                doc = Document(file)
                record = extract_field_data(doc)
                extracted_records.append(record)

            output_buffer = build_summary_doc(extracted_records)

        st.success("Extraction Complete!")
        
        st.download_button(
            label="📥 Download Extracted Summary (.docx)",
            data=output_buffer,
            file_name="Lampiran_2_Pindaan_raw.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
