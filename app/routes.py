"""
GenoPilot – Rutas Flask
- '/' : formulario
- '/generate' : genera PDF (DOCX -> PDF) usando plantilla con {{TABLA_RESULTADO}}
"""
from __future__ import annotations
from flask import Blueprint, render_template, request, send_file
import os, json, datetime, re, tempfile
from docxtpl import DocxTemplate
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

bp = Blueprint("main", __name__)
BASE_DIR = os.path.dirname(os.path.dirname(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
REPORTS_DIR = os.path.join(BASE_DIR, "reports")
TPL_DATA_PATH = os.path.join(DATA_DIR, "GenoPilot_report_template.docx")
TPL_APP_PATH  = os.path.join(os.path.dirname(__file__), "templates", "report_template.docx")
os.makedirs(REPORTS_DIR, exist_ok=True)

# --------- Carga de datos (desde tu Excel ya volcados a JSON) ---------
with open(os.path.join(DATA_DIR, "markers.json"), "r", encoding="utf-8") as f:
    MARKERS = json.load(f)
with open(os.path.join(DATA_DIR, "cyp2d6_stars.json"), "r", encoding="utf-8") as f:
    CYP_STARS = json.load(f)
with open(os.path.join(DATA_DIR, "cyp2d6_pheno.json"), "r", encoding="utf-8") as f:
    CYP_PHENO = json.load(f)
with open(os.path.join(DATA_DIR, "recs.json"), "r", encoding="utf-8") as f:
    RECS = json.load(f)

# --------- Helpers de fenotipo / recomendación ---------
def dpyd_from_inputs(values: dict):
    stars_list, polym = [], []
    for m in MARKERS.get("DPYD", []):
        col, ref, var, star = m["column"], m["ref"], m["var"], m["star"]
        geno = values.get(col, "-/-")
        polym.append(f"{col} ({m['rsid']}): {geno}")
        alt = re.split(r"[\/,\s]+", var)[0] if var else ""
        if alt and alt != "-":
            if geno in (f"{ref}/{alt}", f"{alt}/{ref}"):
                stars_list.append(star)
            elif geno == f"{alt}/{alt}":
                stars_list.extend([star, star])
    if not stars_list:
        dipl, pheno = "*1/*1", "Metabolizador normal"
    elif len(stars_list) == 1:
        dipl, pheno = f"*1/{stars_list[0]}", "Metabolizador intermedio"
    else:
        dipl, pheno = f"{stars_list[0]}/{stars_list[1]}", "Metabolizador lento"
    rec_row = next((r for r in RECS if r["Gene"]=="DPYD" and r["Phenotype"].lower() in pheno.lower()), None)
    rec_text = rtext_short("DPYD", pheno, rec_row["RecText"] if rec_row else "—")
    return dipl, pheno, rec_text, polym

def ugt1a1_from_inputs(geno: str):
    polym = [f"UGT1A1*80 (rs887829): {geno}"]
    if geno == "C/C":
        dipl, pheno = "*1/*1", "Metabolizador normal"
    elif geno in ("C/T", "T/C"):
        dipl, pheno = "*1/*28", "Metabolizador intermedio"
    elif geno == "T/T":
        dipl, pheno = "*28/*28", "Metabolizador lento"
    else:
        dipl, pheno = "-/-", "Indeterminado"
    rec_row = next((r for r in RECS if r["Gene"]=="UGT1A1" and r["Phenotype"].lower() in pheno.lower()), None)
    rec_text = rtext_short("UGT1A1", pheno, rec_row["RecText"] if rec_row else "—")
    return dipl, pheno, rec_text, polym

def _cyp_lookup_pheno(dipl: str) -> str:
    row = next((r for r in CYP_PHENO if r["CYP2D6 Diplotype"]==dipl), None)
    if row:
        return row["Coded Diplotype/Phenotype Summary"]
    # Backoff mínimo por “activity score”
    def ascore(star: str) -> float:
        loss = {"*3","*4","*5","*6","*7","*14","*15","*19","*59"}
        reduced = {"*10","*17","*29","*41","*56B"}
        if star in loss: return 0.0
        if star in reduced: return 0.5
        return 1.0
    a1, a2 = dipl.split("/")
    s = ascore(a1) + ascore(a2)
    if s == 0: return "Poor Metabolizer"
    if s <= 1.0: return "Intermediate Metabolizer"
    if s <= 2.25: return "Normal Metabolizer"
    return "Ultrarapid Metabolizer"

def cyp2d6_from_stars(a1: str, a2: str):
    dipl = f"{a1}/{a2}"
    pheno = _cyp_lookup_pheno(dipl)
    # Normalizamos etiqueta a ES
    conv = {"Normal":"Metabolizador normal",
            "Intermediate":"Metabolizador intermedio",
            "Poor":"Metabolizador pobre",
            "Ultrarapid":"Metabolizador ultrarrápido"}
    for k,v in conv.items():
        if pheno.startswith(k):
            pheno = v
            break
    rec_row = next((r for r in RECS if r["Gene"]=="CYP2D6" and ((pheno.split()[1] in r.get("Phenotype","")) if " " in pheno else pheno in r.get("Phenotype",""))), None)
    rec_text = rtext_short("CYP2D6", pheno, (rec_row or {}).get("RecText","—"))
    polym = [f"CYP2D6 diplotipo manual: {dipl}"]
    return dipl, pheno, rec_text, polym

def cyp2d6_from_markers(values: dict):
    hits, polym = [], []
    for m in MARKERS.get("CYP2D6", []):
        col, ref, var, star = m["column"], m["ref"], m["var"], str(m["star"]).split()[0]
        geno = values.get(col, "-/-")
        polym.append(f"{col} ({m['rsid']}): {geno}")
        alt = re.split(r"[\/,\s]+", var)[0] if var else ""
        if not star or star == "-":
            continue
        if alt and alt != "-":
            if geno in (f"{ref}/{alt}", f"{alt}/{ref}"):
                hits.append(star)
            elif geno == f"{alt}/{alt}":
                hits.extend([star, star])
    if not hits:
        dipl = "*1/*1"
    elif len(hits) == 1:
        dipl = f"*1/{hits[0]}"
    else:
        dipl = f"{hits[0]}/{hits[1]}"
    pheno = _cyp_lookup_pheno(dipl)
    conv = {"Normal":"Metabolizador normal",
            "Intermediate":"Metabolizador intermedio",
            "Poor":"Metabolizador pobre",
            "Ultrarapid":"Metabolizador ultrarrápido"}
    for k,v in conv.items():
        if pheno.startswith(k):
            pheno = v
            break
    rec_row = next((r for r in RECS if r["Gene"]=="CYP2D6" and (pheno.split()[1] in r.get("Phenotype","") if " " in pheno else pheno in r.get("Phenotype",""))), None)
    rec_text = rtext_short("CYP2D6", pheno, (rec_row or {}).get("RecText","—"))
    return dipl, pheno, rec_text, polym

# --------- Recomendación corta (para que quepa en la tabla) ---------
SHORT_REC = {
    ("DPYD","metabolizador normal"):       "Dosis estándar según ficha técnica.",
    ("DPYD","metabolizador intermedio"):   "Reducir dosis inicial; titular y monitorizar.",
    ("DPYD","metabolizador lento"):        "Evitar fluoropirimidinas; valorar alternativas.",
    ("UGT1A1","metabolizador normal"):     "Dosis estándar según ficha técnica.",
    ("UGT1A1","metabolizador intermedio"): "Considerar reducción dosis; monitorizar neutropenia.",
    ("UGT1A1","metabolizador lento"):      "Reducir dosis inicial (30–50%); monitorización estrecha.",
    ("CYP2D6","metabolizador normal"):     "Dosis estándar.",
    ("CYP2D6","metabolizador intermedio"): "Considerar terapia hormonal alternativa.",
    ("CYP2D6","metabolizador pobre"):      "Evitar tamoxifeno; alternativa terapéutica.",
    ("CYP2D6","metabolizador ultrarrápido"): "Valorar alternativas según contexto."
}
def rtext_short(gene: str, pheno: str, long_text: str) -> str:
    key = (gene, pheno.lower())
    if key in SHORT_REC: return SHORT_REC[key]
    # fallback: limpiar URLs y citas para acortar
    t = re.sub(r'\[.*?\]', '', long_text)
    t = re.sub(r'https?://\S+', '', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

# --------- Tabla resultado como subdocumento (tamaño y anchos) ---------
def _shade(cell, hex_fill: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_fill)
    tcPr.append(shd)

def _set_fixed_layout(table):
    tblPr = table._tbl.tblPr
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'), 'fixed')
    tblPr.append(tblLayout)

def _set_col_widths(table, widths_cm):
    # desactiva autofit y fija anchos (en todas las filas)
    _set_fixed_layout(table)
    widths = [Cm(w) for w in widths_cm]
    for row in table.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = w

def _set_font_size(cell, pt):
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.size = Pt(pt)

def build_result_subdoc(tpl: DocxTemplate, summary: list):
    sub = tpl.new_subdoc()
    table = sub.add_table(rows=1, cols=4)
    table.style = 'Table Grid'

    # Cabecera
    hdr = table.rows[0].cells
    headers = ["Gen", "Fenotipo", "Fármaco de interés", "Recomendación terapéutica"]
    for i, t in enumerate(headers):
        hdr[i].text = t
        _shade(hdr[i], 'DDDDDD')
        _set_font_size(hdr[i], 10)

    # Filas (fuente 9.5; última columna 9)
    for row in summary:
        r = table.add_row().cells
        r[0].text = f"{row['gen']}\n{row['diplotipo']}"
        r[1].text = row['fenotipo']
        r[2].text = row['drug']
        r[3].text = row['rec']

        fen = row['fenotipo'].lower()
        if "normal" in fen:
            shade = 'E6F4E6'  # verde suave
        elif any(k in fen for k in ("intermedio","pobre","ultra")):
            shade = 'F6DEDE'  # rosa suave
        else:
            shade = 'FFFFFF'
        _shade(r[1], shade)
        _shade(r[3], shade)

        # tamaños
        for c in (r[0], r[1], r[2]):
            _set_font_size(c, 9.5)
        _set_font_size(r[3], 9)

    # Anchos (que “ajustan” la primera 3 y dejan aire a la última)
    # Gene(3.2cm), Fenotipo(5.2cm), Fármaco(5.2cm), Recomendación(8.4cm) ≈ A4 margen estándar
    _set_col_widths(table, [3.2, 5.2, 5.2, 8.4])

    return sub

# --------- Rutas ----------
@bp.route("/", methods=["GET", "POST"])
def index():
    dpyd_inputs = {m["column"]: m["options"] for m in MARKERS.get("DPYD", [])}
    ugt_inputs  = {m["column"]: m["options"] for m in MARKERS.get("UGT1A1", [])}
    cyp_inputs  = [(m["column"], m["options"]) for m in MARKERS.get("CYP2D6", [])]

    if request.method == "POST":
        patient = {
            "nombre": request.form.get("nombre",""),
            "apellidos": request.form.get("apellidos",""),
            "full_name": (request.form.get("nombre","") + " " + request.form.get("apellidos","")).strip(),
            "historia": request.form.get("historia",""),
            "sexo": request.form.get("sexo","-"),
            "fecha_nac": request.form.get("fecha_nac","")
        }
        clinical = {
            "enf_actual": request.form.get("enf_actual",""),
            "otras_pat": request.form.get("otras_pat",""),
            "tratamiento": request.form.get("tto",""),
        }

        dpyd_vals = {k: request.form.get(k,"-/-") for k in dpyd_inputs.keys()}
        dpyd_dipl, dpyd_pheno, dpyd_rec, dpyd_poly = dpyd_from_inputs(dpyd_vals)

        ugt_col = next(iter(ugt_inputs.keys()))
        ugt_geno = request.form.get(ugt_col, "-/-")
        ugt_dipl, ugt_pheno, ugt_rec, ugt_poly = ugt1a1_from_inputs(ugt_geno)

        cyp_mode = request.form.get("cyp_mode","diplotype")
        if cyp_mode == "markers":
            cyp_vals = {m[0]: request.form.get(m[0], "-/-") for m in cyp_inputs}
            cyp_dipl, cyp_pheno, cyp_rec, cyp_poly = cyp2d6_from_markers(cyp_vals)
        else:
            cyp_vals = {}
            cyp_a1 = request.form.get("cyp_a1","*1")
            cyp_a2 = request.form.get("cyp_a2","*1")
            cyp_dipl, cyp_pheno, cyp_rec, cyp_poly = cyp2d6_from_stars(cyp_a1, cyp_a2)

        preview = [
            {"gen":"DPYD",   "dipl":dpyd_dipl, "pheno":dpyd_pheno, "drug":"Fluorouracilo, capecitabina, tegafur", "rec":dpyd_rec},
            {"gen":"CYP2D6", "dipl":cyp_dipl,  "pheno":cyp_pheno,  "drug":"Tamoxifeno",                            "rec":cyp_rec},
            {"gen":"UGT1A1", "dipl":ugt_dipl,  "pheno":ugt_pheno,  "drug":"Irinotecán",                            "rec":ugt_rec},
        ]

        return render_template(
            "preview.html",
            patient=patient, clinical=clinical, preview=preview,
            dpyd_inputs=dpyd_inputs, ugt_inputs=ugt_inputs,
            dpyd_vals=dpyd_vals, ugt_geno=ugt_geno,
            cyp_mode=cyp_mode, cyp_inputs=cyp_inputs, cyp_vals=cyp_vals,
            cyp_a1=request.form.get("cyp_a1","*1"),
            cyp_a2=request.form.get("cyp_a2","*1"),
        )

    return render_template(
        "index.html",
        dpyd_inputs={m["column"]: m["options"] for m in MARKERS.get("DPYD", [])},
        ugt_inputs={m["column"]: m["options"] for m in MARKERS.get("UGT1A1", [])},
        cyp_inputs=[(m["column"], m["options"]) for m in MARKERS.get("CYP2D6", [])],
        cyp_stars=CYP_STARS,
    )

@bp.route("/generate", methods=["POST"])
def generate():
    # Recoger todo (idéntico a index POST)
    patient = {
        "nombre": request.form.get("nombre",""),
        "apellidos": request.form.get("apellidos",""),
        "full_name": (request.form.get("nombre","") + " " + request.form.get("apellidos","")).strip(),
        "historia": request.form.get("historia",""),
        "sexo": request.form.get("sexo","-"),
        "fecha_nac": request.form.get("fecha_nac","")
    }
    clinical = {
        "enf_actual": request.form.get("enf_actual",""),
        "otras_pat": request.form.get("otras_pat",""),
        "tratamiento": request.form.get("tto",""),
    }

    dpyd_inputs = {m["column"]: m["options"] for m in MARKERS.get("DPYD", [])}
    dpyd_vals = {k: request.form.get(k,"-/-") for k in dpyd_inputs.keys()}
    dpyd_dipl, dpyd_pheno, dpyd_rec, dpyd_poly = dpyd_from_inputs(dpyd_vals)

    ugt_inputs = {m["column"]: m["options"] for m in MARKERS.get("UGT1A1", [])}
    ugt_col = next(iter(ugt_inputs.keys()))
    ugt_geno = request.form.get(ugt_col, "-/-")
    ugt_dipl, ugt_pheno, ugt_rec, ugt_poly = ugt1a1_from_inputs(ugt_geno)

    cyp_mode = request.form.get("cyp_mode","diplotype")
    if cyp_mode == "markers":
        cyp_inputs = [(m["column"], m["options"]) for m in MARKERS.get("CYP2D6", [])]
        cyp_vals = {m[0]: request.form.get(m[0], "-/-") for m in cyp_inputs}
        cyp_dipl, cyp_pheno, cyp_rec, cyp_poly = cyp2d6_from_markers(cyp_vals)
    else:
        cyp_a1 = request.form.get("cyp_a1","*1")
        cyp_a2 = request.form.get("cyp_a2","*1")
        cyp_dipl, cyp_pheno, cyp_rec, cyp_poly = cyp2d6_from_stars(cyp_a1, cyp_a2)

    summary = [
        {"gen":"DPYD",   "diplotipo":dpyd_dipl, "fenotipo":dpyd_pheno, "drug":"Fluorouracilo, capecitabina, tegafur", "rec":dpyd_rec},
        {"gen":"CYP2D6", "diplotipo":cyp_dipl,  "fenotipo":cyp_pheno,  "drug":"Tamoxifeno",                            "rec":cyp_rec},
        {"gen":"UGT1A1", "diplotipo":ugt_dipl,  "fenotipo":ugt_pheno,  "drug":"Irinotecán",                            "rec":ugt_rec},
    ]

    tpl_path = TPL_DATA_PATH if os.path.exists(TPL_DATA_PATH) else TPL_APP_PATH
    tpl = DocxTemplate(tpl_path)
    tabla_subdoc = build_result_subdoc(tpl, summary)

    now = datetime.datetime.now()
    context = {
        "patient": patient,
        "clinical": clinical,
        "TABLA_RESULTADO": tabla_subdoc,
        "polymorphisms": "; ".join(dpyd_poly + cyp_poly + ugt_poly),
        "sources": {
            "cpic_url": "https://cpicpgx.org/guidelines/",
            "dpwgd_doi": "DOI:10.1038/s41431-022-01243-2",
        },
        "meta": {
            "sample_code": patient.get("historia") or "—",
            "request_date": now.strftime("%d/%m/%Y"),
            "report_date": now.strftime("%d/%m/%Y"),
            "generated_at": now.strftime("%Y-%m-%d %H:%M"),
            "last_update": "septiembre 2025",
            "version": "0.4.0",
        },
    }

    # 1) Render DOCX temporal
    tmp_docx = os.path.join(tempfile.gettempdir(), f"GenoPilot_tmp_{now.strftime('%H%M%S')}.docx")
    tpl.render(context)
    tpl.save(tmp_docx)

    # 2) Convertir a PDF (Windows + Word) y servir el PDF
    pdf_name = f"GenoPilot_{patient.get('full_name','Paciente')}_{now.strftime('%Y%m%d_%H%M')}.pdf"
    pdf_path = os.path.join(REPORTS_DIR, pdf_name)

    try:
        from docx2pdf import convert
        convert(tmp_docx, pdf_path)   # requiere MS Word instalado en Windows
        return send_file(pdf_path, as_attachment=True, download_name=pdf_name, mimetype="application/pdf")
    except Exception as e:
        # Fallback: si no hay Word/docx2pdf, devolvemos el DOCX para no bloquear la entrega
        fallback_name = pdf_name.replace(".pdf", ".docx")
        return send_file(tmp_docx, as_attachment=True,
                         download_name=fallback_name,
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
