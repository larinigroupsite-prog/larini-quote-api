# app.py
import io, json
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

# --- ReportLab / layout ---
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.utils import ImageReader

# Optional parser (best-effort)
try:
    import pdfplumber
except Exception:
    pdfplumber = None
try:
    import docx
except Exception:
    docx = None

LARINI_RED = colors.Color(193/255.0, 18/255.0, 31/255.0)
MARGIN_PT = 42.52                # ~15 mm
BOTTOM_OFFSET_PT = 0.5 * cm      # 0.5 cm = 14.173 pt
RESERVED_GAP_ABOVE_FOOTER_PT = 10
VEHICLE_TOP_BOTTOM_GAP_PT = 28.35  # ~1 cm
MAX_VEHICLE_RATIO = 0.30           # max 30% area utile (prima pagina)

def build_styles():
    styles = getSampleStyleSheet()
    base = styles["Normal"]; base.fontName="Helvetica"; base.fontSize=11; base.leading=14; base.spaceAfter=6
    title = ParagraphStyle("TitleBold", parent=base, fontName="Helvetica-Bold", fontSize=11, leading=14, spaceBefore=8, spaceAfter=6)
    price = ParagraphStyle("Price", parent=base, alignment=TA_CENTER, fontName="Helvetica-Bold",
                           fontSize=24, leading=28, textColor=LARINI_RED, spaceBefore=8, spaceAfter=8)
    italic = ParagraphStyle("ItalicNote", parent=base, fontName="Helvetica-Oblique", alignment=TA_CENTER, fontSize=11, leading=14, spaceBefore=2, spaceAfter=8)
    return {"base": base, "title": title, "price": price, "italic": italic}

def read_image_size(path): return ImageReader(str(path)).getSize()
def scaled_height_for_full_width(wpx, hpx, target_wpt):
    if wpx <= 0 or hpx <= 0: return 0
    return hpx * (target_wpt/float(wpx))
def calc_footer_height(page_width_pt, footer_path):
    w,h = read_image_size(footer_path); return scaled_height_for_full_width(w, h, page_width_pt)
def calc_header_height(usable_width_pt, header_path):
    w,h = read_image_size(header_path); return scaled_height_for_full_width(w, h, usable_width_pt)

def parse_supplier_file(path: Path):
    data, text = {}, ""
    if path.suffix.lower()==".pdf" and pdfplumber is not None:
        try:
            with pdfplumber.open(str(path)) as pdf:
                for page in pdf.pages[:5]:
                    text += "\n" + (page.extract_text() or "")
        except Exception: pass
    elif path.suffix.lower() in [".docx",".doc"] and docx is not None:
        try:
            d = docx.Document(str(path))
            for para in d.paragraphs: text += "\n"+para.text
        except Exception: pass

    import re
    text = re.sub(r'[^\x00-\x7F]+',' ',text); text = re.sub(r'\s+',' ',text).strip()
    pats = {
        "cliente_nome": r"(?:Cliente|Ragione sociale|Nome)\s*[:\-]\s*([A-Za-z0-9\s\.\,\-]+)",
        "piva_cf": r"(?:P\.?\s*IVA|CF|Codice fiscale)\s*[:\-]\s*([A-Za-z0-9]+)",
        "sede": r"(?:Sede|Indirizzo)\s*[:\-]\s*([A-Za-z0-9\s\.\,\-]+)",
        "referente": r"(?:Referente|Contatto)\s*[:\-]\s*([A-Za-z0-9\s\.\,\-]+)",
        "marca_modello": r"(?:Marca\s*\/?\s*Modello|Veicolo|Modello)\s*[:\-]\s*([A-Za-z0-9\s\.\,\-]+)",
        "versione": r"(?:Versione|Allestimento)\s*[:\-]\s*([A-Za-z0-9\s\.\,\-]+)",
        "motore": r"(?:Motore|Alimentazione|Cambio|Potenza)\s*[:\-]\s*([A-Za-z0-9\s\.\,\-\/]+)",
        "neopatentati": r"(?:Neopatentati|Idoneo neopatentati)\s*[:\-]\s*(Si|No|N\.D\.)",
        "consegna": r"(?:Consegna|Consegna stimata|Lead time)\s*[:\-]\s*([A-Za-z0-9\s\.\,\-\/]+)",
        "durata": r"(?:Durata)\s*[:\-]\s*([0-9]{1,2})",
        "km_annui": r"(?:Km|Chilometraggio)\s*[:\-]\s*([0-9\.\,]+)",
        "anticipo": r"(?:Anticipo)\s*[:\-]\s*([0-9\.\,]+)",
        "canone": r"(?:Canone|Canone mensile)\s*[:\-]\s*([0-9\.\,]+)"
    }
    for k, pat in pats.items():
        m = re.search(pat, text, re.IGNORECASE)
        if m: data[k]=m.group(1).strip()
    return data

def build_story(data, vehicle_photo_path: Optional[Path], styles, first_page_usable_height):
    story = []
    if vehicle_photo_path and vehicle_photo_path.exists():
        img = Image(str(vehicle_photo_path)); img.hAlign="CENTER"
        story += [Spacer(1,VEHICLE_TOP_BOTTOM_GAP_PT), img, Spacer(1,VEHICLE_TOP_BOTTOM_GAP_PT)]
        img._restrictSize(10*cm, MAX_VEHICLE_RATIO*first_page_usable_height)

    story += [
        Paragraph("DATI CLIENTE", styles["title"]),
        Paragraph(f"Ragione sociale / Nome: {data.get('cliente_nome','N.D.')}", styles["base"]),
        Paragraph(f"P.IVA / CF: {data.get('piva_cf','N.D.')}", styles["base"]),
        Paragraph(f"Sede: {data.get('sede','N.D.')}", styles["base"]),
        Paragraph(f"Referente: {data.get('referente','N.D.')}", styles["base"]),
        Spacer(1,6),
        Paragraph("VEICOLO PROPOSTO", styles["title"]),
        Paragraph(f"Marca e Modello: {data.get('marca_modello','N.D.')}", styles["base"]),
        Paragraph(f"Versione / Allestimento: {data.get('versione','N.D.')}", styles["base"]),
        Paragraph(f"Motore / Alimentazione / Cambio / Potenza: {data.get('motore','N.D.')}", styles["base"]),
        Paragraph(f"Idoneo neopatentati: {data.get('neopatentati','N.D.')}", styles["base"]),
        Paragraph(f"Consegna stimata: {data.get('consegna','N.D.')}", styles["base"]),
        Spacer(1,6),
        Paragraph("CONDIZIONI ECONOMICHE", styles["title"]),
        Paragraph(f"Durata: {data.get('durata','N.D.')} mesi", styles["base"]),
        Paragraph(f"Chilometraggio: {data.get('km_annui','N.D.')} km/anno", styles["base"]),
        Paragraph(f"Anticipo: {data.get('anticipo','N.D.')} euro", styles["base"]),
    ]
    price = data.get("canone","N.D.")
    if price and not str(price).lower().startswith("n.d"):
        story += [Spacer(1,4), Paragraph(f"{price} euro + IVA al mese", styles["price"]), Paragraph("Tutto incluso - senza sorprese", styles["italic"])]
    else:
        story += [Paragraph("Canone mensile: N.D.", styles["base"])]

    story += [Spacer(1,6), Paragraph("SERVIZI INCLUSI", styles["title"])]
    for s in ["RCA","Kasko / Danni / Furto & Incendio","Manutenzione ordinaria e straordinaria","Pneumatici (premium o equivalenti)","Assistenza stradale 24h","Gestione sinistri e contravvenzioni","Veicolo sostitutivo (se previsto)","Immatricolazione e consegna"]:
        story.append(Paragraph(f"- {s}", styles["base"]))

    story += [
        Spacer(1,6), Paragraph("DOCUMENTAZIONE RICHIESTA", styles["title"]),
        Paragraph("SOCIETA (SRL, SAS, SPA, SRLS): documento; cod. fiscale; visura aggiornata a 6 mesi; ultimo bilancio depositato con ricevuta", styles["base"]),
        Paragraph("DITTA INDIVIDUALE / SNC / LIBERO PROFESSIONISTA: documento; cod. fiscale; visura aggiornata a 6 mesi; modello unico", styles["base"]),
        Paragraph("PRIVATI: documento; cod. fiscale; CUD anno precedente; ultime 2 buste paga", styles["base"]),
        Paragraph("PENSIONATI: documento; cod. fiscale; cedolini o estratto conto", styles["base"]),
        Spacer(1,6), Paragraph("TERMINI E CONDIZIONI", styles["title"]),
        Paragraph("Offerta soggetta ad approvazione della societa di noleggio; Immagini a scopo illustrativo; Canoni IVA esclusa salvo diversa indicazione; Disponibilita e canone variabili secondo valutazione creditizia", styles["base"]),
        Spacer(1,6), Paragraph("CONTATTI", styles["title"]),
        Paragraph("Larini Automotive Rent | Tel. 379 2114207 | noleggio@larini.it | www.larinirent.it", styles["base"])
    ]
    return story

def generate_pdf_bytes(data: dict, vehicle_photo: Optional[Path], header_path: Path, footer_path: Path) -> bytes:
    from reportlab.platypus import PageTemplate, BaseDocTemplate, Frame
    page_width, page_height = A4
    footer_h = calc_footer_height(page_width, footer_path)
    bottom_reserved = BOTTOM_OFFSET_PT + footer_h + RESERVED_GAP_ABOVE_FOOTER_PT
    usable_width = page_width - 2*MARGIN_PT
    header_h = calc_header_height(usable_width, header_path)

    first_top = page_height - MARGIN_PT - header_h
    later_top = page_height - MARGIN_PT
    h_first = max(first_top - bottom_reserved, 100)
    h_later = max(later_top - bottom_reserved, 100)

    import io
    buf = io.BytesIO()
    doc = BaseDocTemplate(buf, pagesize=A4, leftMargin=MARGIN_PT, rightMargin=MARGIN_PT, topMargin=MARGIN_PT, bottomMargin=MARGIN_PT)
    frame_first = Frame(MARGIN_PT, bottom_reserved, usable_width, h_first, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, showBoundary=0)
    frame_later = Frame(MARGIN_PT, bottom_reserved, usable_width, h_later, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, showBoundary=0)

    def draw_footer(c, footer_path, page_width, page_height):
        w,h = read_image_size(footer_path)
        th = scaled_height_for_full_width(w, h, page_width)
        c.drawImage(str(footer_path), 0, BOTTOM_OFFSET_PT, width=page_width, height=th, preserveAspectRatio=True, mask='auto')

    def draw_header_first_page(c, header_path, margin_left, page_width, page_height, usable_width):
        w,h = read_image_size(header_path)
        hh = scaled_height_for_full_width(w, h, usable_width)
        x = margin_left; y = page_height - hh - MARGIN_PT
        c.drawImage(str(header_path), x, y, width=usable_width, height=hh, preserveAspectRatio=True, mask='auto')

    def on_first_page(c, _doc):
        draw_header_first_page(c, header_path, MARGIN_PT, page_width, page_height, usable_width)
        draw_footer(c, footer_path, page_width, page_height)

    def on_later_pages(c, _doc):
        draw_footer(c, footer_path, page_width, page_height)

    doc.addPageTemplates([
        PageTemplate(id="First", frames=[frame_first], onPage=on_first_page),
        PageTemplate(id="Later", frames=[frame_later], onPage=on_later_pages),
    ])
    styles = build_styles()
    story = build_story(data, vehicle_photo, styles, h_first)
    doc.build(story)
    pdf = buf.getvalue(); buf.close(); return pdf

# ---------- FastAPI ----------
app = FastAPI(title="Larini Quote API", version="1.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

class HealthResp(BaseModel):
    status: str

@app.get("/health")
def health(): return HealthResp(status="ok")

@app.post("/generate")
async def generate(
    data_json: Optional[str] = Form(default=None),
    vehicle_photo: Optional[UploadFile] = File(default=None),
    supplier_file: Optional[UploadFile] = File(default=None),
    header_image: Optional[UploadFile] = File(default=None),
    footer_image: Optional[UploadFile] = File(default=None),
    output_name: Optional[str] = Form(default=None)
):
    assets = Path("assets"); assets.mkdir(exist_ok=True)
    default_header = assets / "Heater.jpg"
    default_footer = assets / "footer.jpg"
    if not default_header.exists() or not default_footer.exists():
        raise HTTPException(500, "Brand assets missing. Upload header_image/footer_image or include assets/Heater.jpg and assets/footer.jpg in the repo.")

    tmp = Path("tmp"); tmp.mkdir(exist_ok=True)

    # data
    data = {}
    if data_json:
        try:
            data = json.loads(data_json)
        except Exception as e:
            raise HTTPException(400, f"Invalid data_json: {e}")

    # optional supplier parse
    if supplier_file is not None:
        sp = tmp / supplier_file.filename
        with sp.open("wb") as f: f.write(await supplier_file.read())
        parsed = parse_supplier_file(sp)
        for k,v in parsed.items():
            data.setdefault(k,v)

    # defaults
    defaults = {"cliente_nome":"N.D.","piva_cf":"N.D.","sede":"N.D.","referente":"N.D.","marca_modello":"N.D.","versione":"N.D.","motore":"N.D.","neopatentati":"N.D.","consegna":"N.D.","durata":"N.D.","km_annui":"N.D.","anticipo":"N.D.","canone":"N.D."}
    for k,v in defaults.items():
        data.setdefault(k,v)

    # vehicle photo
    vpath = None
    if vehicle_photo is not None:
        vpath = tmp / vehicle_photo.filename
        with vpath.open("wb") as f: f.write(await vehicle_photo.read())

    # header/footer override
    header_path = default_header
    footer_path = default_footer
    if header_image is not None:
        header_path = tmp / f"header_{header_image.filename}"
        with header_path.open("wb") as f: f.write(await header_image.read())
    if footer_image is not None:
        footer_path = tmp / f"footer_{footer_image.filename}"
        with footer_path.open("wb") as f: f.write(await footer_image.read())

    # output filename
    if not output_name:
        model = data.get("marca_modello","Modello").replace(" ","_")
        output_name = f"Preventivo_Larini_{model}_{datetime.utcnow().date().isoformat()}.pdf"

    try:
        pdf = generate_pdf_bytes(data, vpath, header_path, footer_path)
    except Exception as e:
        raise HTTPException(500, f"PDF generation failed: {e}")

    return StreamingResponse(io.BytesIO(pdf), media_type="application/pdf",
                             headers={"Content-Disposition": f'attachment; filename=\"%s\"' % output_name})
