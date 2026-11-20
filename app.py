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
    """Extracts fields using positional table coordinates and targeted regex."""
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

    # 2. No. Fail (Matches KKM87/A2/... with or without Jilid)
    fail_match = re.search(r'(KKM[\s/]*87/A2/[A-Za-z0-9/.-]+(?:[\s]*JILID[\s]*\d+)?)', full_document_text, re.IGNORECASE)
    if fail_match:
        data["no_fail"] = fail_match.group(1).strip()

    # 3. Pegawai Proses (Captures name directly, retaining Dr./KJ prefixes)
    pegawai_match = re.search(r'Pegawai\s*Proses[\s:\|]*([^\n\r\|]+)', full_document_text, re.IGNORECASE)
    if pegawai_match:
        data["pegawai_proses"] = pegawai_match.group(1).strip()

    # --- Step 2: Positional Extraction for Table Data ---
    for table in doc.tables:
        if len(table.rows) >= 2:
            
            # Safely get actual columns of the 2nd row (index 1), skipping merged cell duplicates
            row_2_cells = []
            seen_cells = set()
            for cell in table.rows[1].cells:
                if cell._tc not in seen_cells:
                    seen_cells.add(cell._tc)
                    row_2_cells.append(cell)
            
            # Ensure it's the right table (at least 4 columns)
            if len(row_2_cells) >= 4:
                
                # --- 1. Nama Premis & Alamat Premis (2nd Row, 2nd Column) ---
                cell_2_2_text = row_2_cells[1].text.strip()
                premis_match = re.search(r'([\s\S]*?)(?=Pemohon:)', cell_2_2_text, re.IGNORECASE)
                raw_premis_alamat = premis_match.group(1).strip() if premis_match else cell_2_2_text
                
                # Split Nama (First Line) and Alamat (Remaining Lines)
                lines = [line.strip() for line in raw_premis_alamat.split('\n') if line.strip()]
                if lines:
                    data["nama_premis"] = lines[0]
                    data["alamat_premis"] = " ".join(lines[1:])
                
                # --- 4. Isu / Ulasan (Single Block Extraction) ---
                
                # Parse the 4th column (index 3) paragraph by paragraph to preserve list structure
                cell_2_4 = row_2_cells[3]
                cell_text_parts = []
                for p in cell_2_4.paragraphs:
                    text = p.text.strip()
                    if text:
                        style_name = p.style.name if p.style else ""
                        if 'Number' in style_name:
                            cell_text_parts.append(f"1. {text}")
                        elif 'List' in style_name or 'Bullet' in style_name:
                            cell_text_parts.append(f"• {text}")
                        else:
                            cell_text_parts.append(text)
                            
                cell_2_4_text = "\n".join(cell_text_parts)
                
                # Regex: Capture from "Permohonan" to the end of the line containing "tarikh pindaan bermula"
                isu_match = re.search(
                    r'(Permohonan[\s\S]*?tarikh pindaan bermula[^\n\r]*)', 
                    cell_2_4_text, 
                    re.IGNORECASE
                )
                
                if isu_match:
                    data["isu_ulasan"] = isu_match.group(1).strip()
                else:
                    # Fallback to the whole cell text if the exact phrase isn't found
                    data["isu_ulasan"] = cell_2_4_text
                
                break # Stop searching tables once the target is processed

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
            
            clean_word = re.sub(r'[^\w]', '', sub).lower()
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
        
        # 1. Nama Premis (Bold, no label)
        val_premis = p_left.add_run(f"{data.get('nama_premis', '')}\n")
        val_premis.font.name = 'Arial'
        val_premis.font.size = Pt(9)
        val_premis.bold = True

        # 2. Alamat Premis (Value only, no label)
        val_alamat = p_left.add_run(f"{data.get('alamat_premis', '')}\n")
        val_alamat.font.name = 'Arial'
        val_alamat.font.size = Pt(9)
        val_alamat.bold = False

        # 3. No. Fail (Label + Value)
        lbl_fail = p_left.add_run("No. Fail: ")
        lbl_fail.font.name = 'Arial'
        lbl_fail.font.size = Pt(9)
        lbl_fail.bold = True
        
        val_fail = p_left.add_run(f"{data.get('no_fail', '')}\n")
        val_fail.font.name = 'Arial'
        val_fail.font.size = Pt(9)
        val_fail.bold = False

        # 4. Pegawai Proses (Label + Value)
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
        
        # Force a double newline specifically before "Tarikh pindaan bermula"
        # \s* removes existing spacing so it doesn't accidentally create 3 or 4 spaces
        right_text = re.sub(r'\s*(Tarikh pindaan)', r'\n\n\1', right_text, flags=re.IGNORECASE)
        
        lines = right_text.split('\n')
        
        first_paragraph_used = False
        for line in lines:
            clean_line = line.strip()
            
            if not first_paragraph_used:
                p_right = row_cells[1].paragraphs[0]
                first_paragraph_used = True
            else:
                p_right = row_cells[1].add_paragraph()
            
            # If the line is empty, it just leaves the paragraph blank (like an Enter)
            if clean_line:
                p_right.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                format_styled_text_run(p_right, clean_line)

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
            file_name="Lampiran 2 Pindaan raw.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
