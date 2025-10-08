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

# --------- Stars para DPYD / UGT1A1 (derivadas de MARKERS) ---------
def _collect_stars(gene: str, extra: list[str]|None=None) -> list[str]:
    stars = {"*1"}
    for m in MARKERS.get(gene, []):
        s = str(m.get("star","")).strip()
        if not s or s == "-": 
            continue
        # Estandariza: toma el primer token (p.ej. "HapB3" o "*2A")
        s = s.split()[0]
        stars.add(s)
    if extra:
        stars.update(extra)
    # asegura orden humano: *1, *2A, *13, HapB3...
    return sorted(stars, key=lambda x: (x.replace("*","0"), x))
DPYD_STARS  = _collect_stars("DPYD")   # → {"*1","*2A","*13","HapB3","c.2846A>T",...}
UGT1A1_STARS = _collect_stars("UGT1A1") # → {"*1","*28",...}

# --------- Helpers de fenotipo / recomendación ---------
def rtext_short(gene: str, pheno: str, long_text: str) -> str:
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
        ("CYP2D6","metabolizador ultrarrápido"): "Valorar alternativas según contexto.",
    }
    key = (gene, pheno.lower())
    if key in SHORT_REC: 
        return SHORT_REC[key]
    # fallback: limpia enlaces y corchetes
    t = re.sub(r'\[.*?\]', '', long_text or "")
    t = re.sub(r'https?://\S+', '', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t or "—"

# ---------- DPYD ----------
def dpyd_from_markers(values: dict):
    stars_list, polym = [], []
    for m in MARKERS.get("DPYD", []):
        col, ref, var, star = m["column"], m["ref"], m["var"], m["star"]
        geno = values.get(col, "-/-")
        polym.append(f"{col} ({m['rsid']}): {geno}")
        alt = re.split(r"[\/,\s]+", var)[0] if var else ""
        if alt and alt != "-":
            if geno in (f"{ref}/{alt}", f"{alt}/{ref}"):  # hetero
                stars_list.append(star)
            elif geno == f"{alt}/{alt}":                  # homo
                stars_list.extend([star, star])
    if not stars_list:
        dipl, pheno = "*1/*1", "Metabolizador normal"
    elif len(stars_list) == 1:
        dipl, pheno = f"*1/{stars_list[0]}", "Metabolizador intermedio"
    else:
        dipl, pheno = f"{stars_list[0]}/{stars_list[1]}", "Metabolizador lento"
    rec_row = next((r for r in RECS if r["Gene"]=="DPYD" and r["Phenotype"].lower() in pheno.lower()), None)
    return dipl, pheno, rtext_short("DPYD", pheno, (rec_row or {}).get("RecText")), polym

def dpyd_from_diplotype(a1: str, a2: str):
    # Reglas docentes: *1/*1=normal; uno no-*1 -> intermedio; dos no-*1 -> lento
    non1 = sum(1 for x in (a1, a2) if x != "*1")
    if non1 == 0:
        pheno = "Metabolizador normal"
    elif non1 == 1:
        pheno = "Metabolizador intermedio"
    else:
        pheno = "Metabolizador lento"
    dipl = f"{a1}/{a2}"
    rec_row = next((r for r in RECS if r["Gene"]=="DPYD" and r["Phenotype"].lower() in pheno.lower()), None)
    return dipl, pheno, rtext_short("DPYD", pheno, (rec_row or {}).get("RecText")), [f"DPYD diplotipo manual: {dipl}"]

# ---------- UGT1A1 ----------
def ugt1a1_from_markers(geno: str):
    # C/C=*1/*1; C/T=*1/*28; T/T=*28/*28
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
    return dipl, pheno, rtext_short("UGT1A1", pheno, (rec_row or {}).get("RecText")), polym

def ugt1a1_from_diplotype(a1: str, a2: str):
    pair = "/".join(sorted([a1, a2]))
    if pair == "*1/*1":
        pheno = "Metabolizador normal"
    elif pair in ("*1/*28",):
        pheno = "Metabolizador intermedio"
    elif pair == "*28/*28":
        pheno = "Metabolizador lento"
    else:
        pheno = "Indeterminado"
    rec_row = next((r for r in RECS if r["Gene"]=="UGT1A1" and r["Phenotype"].lower() in pheno.lower()), None)
    return pair, pheno, rtext_short("UGT1A1", pheno, (rec_row or {}).get("RecText")), [f"UGT1A1 diplotipo manual: {pair}"]

# ---------- CYP2D6 ----------
def _cyp_lookup_pheno(dipl: str) -> str:
    row = next((r for r in CYP_PHENO if r["CYP2D6 Diplotype"]==dipl), None)
    if row:
        return row["Coded Diplotype/Phenotype Summary"]
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
    conv = {"Normal":"Metabolizador normal","Intermediate":"Metabolizador intermedio","Poor":"Metabolizador pobre","Ultrarapid":"Metabolizador ultrarrápido"}
    for k,v in conv.items():
        if pheno.startswith(k): pheno = v
    rec_row = next(
        (
            r
            for r in RECS
            if r["Gene"] == "CYP2D6"
            and (
                (pheno.split()[1] in r.get("Phenotype", "")) if " " in pheno else (pheno in r.get("Phenotype", ""))
            )
        ),
        None
    )
    return dipl, pheno, rtext_short("CYP2D6", pheno, (rec_row or {}).get("RecText") or ""), [f"CYP2D6 diplotipo manual: {dipl}"]

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
            if geno in (f"{ref}/{alt}", f"{alt}/{ref}"): hits.append(star)
            elif geno == f"{alt}/{alt}": hits.extend([star, star])
    dipl = "*1/*1" if not hits else (f"*1/{hits[0]}" if len(hits)==1 else f"{hits[0]}/{hits[1]}")
    pheno = _cyp_lookup_pheno(dipl)
    conv = {"Normal":"Metabolizador normal","Intermediate":"Metabolizador intermedio","Poor":"Metabolizador pobre","Ultrarapid":"Metabolizador ultrarrápido"}
    for k,v in conv.items():
        if pheno.startswith(k): pheno = v
    rec_row = next(
        (
            r
            for r in RECS
            if r["Gene"] == "CYP2D6"
            and (
                (pheno.split()[1] in r.get("Phenotype", "")) if " " in pheno else (pheno in r.get("Phenotype", ""))
            )
        ),
        None
    )
    return dipl, pheno, rtext_short("CYP2D6", pheno, (rec_row or {}).get("RecText") or ""), polym

# --------- Tabla resultado como subdocumento (compacta) ---------
def _shade(cell, hex_fill: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hex_fill)
    tcPr.append(shd)
def _set_fixed_layout(table):
    tblPr = table._tbl.tblPr
    tblLayout = OxmlElement('w:tblLayout'); tblLayout.set(qn('w:type'), 'fixed'); tblPr.append(tblLayout)
def _set_col_widths(table, widths_cm):
    _set_fixed_layout(table); widths = [Cm(w) for w in widths_cm]
    for row in table.rows:
        for i, w in enumerate(widths): row.cells[i].width = w
def _set_font_size(cell, pt):
    for p in cell.paragraphs:
        for r in p.runs: r.font.size = Pt(pt)
def build_result_subdoc(tpl: DocxTemplate, summary: list):
    sub = tpl.new_subdoc()
    table = sub.add_table(rows=1, cols=4); table.style = 'Table Grid'
    hdr = table.rows[0].cells; headers = ["Gen", "Fenotipo", "Fármaco de interés", "Recomendación terapéutica"]
    for i, t in enumerate(headers): hdr[i].text = t; _shade(hdr[i], 'DDDDDD'); _set_font_size(hdr[i], 10)
    for row in summary:
        r = table.add_row().cells
        r[0].text = f"{row['gen']}\n{row['diplotipo']}"; r[1].text = row['fenotipo']; r[2].text = row['drug']; r[3].text = row['rec']
        fen = row['fenotipo'].lower()
        shade = 'E6F4E6' if "normal" in fen else ('F6DEDE' if any(k in fen for k in ("intermedio","pobre","ultra")) else 'FFFFFF')
        _shade(r[1], shade); _shade(r[3], shade)
        for c in (r[0], r[1], r[2]): _set_font_size(c, 9.5)
        _set_font_size(r[3], 9)
    _set_col_widths(table, [3.2, 5.2, 5.2, 8.4])
    return sub

# --------- Rutas ----------
@bp.route("/", methods=["GET", "POST"])
def index():
    # inputs por marcador
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

        # ----- DPYD -----
        dpyd_mode = request.form.get("dpyd_mode","markers")
        if dpyd_mode == "diplotype":
            dpyd_a1 = request.form.get("dpyd_a1","*1"); dpyd_a2 = request.form.get("dpyd_a2","*1")
            dpyd_dipl, dpyd_pheno, dpyd_rec, dpyd_poly = dpyd_from_diplotype(dpyd_a1, dpyd_a2)
            dpyd_vals = {}
        else:
            dpyd_vals = {k: request.form.get(k,"-/-") for k in dpyd_inputs.keys()}
            dpyd_dipl, dpyd_pheno, dpyd_rec, dpyd_poly = dpyd_from_markers(dpyd_vals)
            dpyd_a1 = dpyd_a2 = ""

        # ----- UGT1A1 -----
        ugt_mode = request.form.get("ugt_mode","markers")
        if ugt_mode == "diplotype":
            ugt_a1 = request.form.get("ugt_a1","*1"); ugt_a2 = request.form.get("ugt_a2","*1")
            ugt_dipl, ugt_pheno, ugt_rec, ugt_poly = ugt1a1_from_diplotype(ugt_a1, ugt_a2)
            ugt_geno = "-/-"
        else:
            ugt_col = next(iter(ugt_inputs.keys()))
            ugt_geno = request.form.get(ugt_col, "-/-")
            ugt_dipl, ugt_pheno, ugt_rec, ugt_poly = ugt1a1_from_markers(ugt_geno)
            ugt_a1 = ugt_a2 = ""

        # ----- CYP2D6 -----
        cyp_mode = request.form.get("cyp_mode","diplotype")
        if cyp_mode == "markers":
            cyp_vals = {m[0]: request.form.get(m[0], "-/-") for m in cyp_inputs}
            cyp_dipl, cyp_pheno, cyp_rec, cyp_poly = cyp2d6_from_markers(cyp_vals)
            cyp_a1 = cyp_a2 = ""
        else:
            cyp_a1 = request.form.get("cyp_a1","*1"); cyp_a2 = request.form.get("cyp_a2","*1")
            cyp_dipl, cyp_pheno, cyp_rec, cyp_poly = cyp2d6_from_stars(cyp_a1, cyp_a2)
            cyp_vals = {}

        preview = [
            {"gen":"DPYD",   "dipl":dpyd_dipl, "pheno":dpyd_pheno, "drug":"Fluorouracilo, capecitabina, tegafur", "rec":dpyd_rec},
            {"gen":"CYP2D6", "dipl":cyp_dipl,  "pheno":cyp_pheno,  "drug":"Tamoxifeno",                            "rec":cyp_rec},
            {"gen":"UGT1A1", "dipl":ugt_dipl,  "pheno":ugt_pheno,  "drug":"Irinotecán",                            "rec":ugt_rec},
        ]

        return render_template(
            "preview.html",
            patient=patient, clinical=clinical, preview=preview,

            # DPYD
            dpyd_mode=dpyd_mode, dpyd_inputs=dpyd_inputs, dpyd_vals=dpyd_vals,
            dpyd_a1=dpyd_a1, dpyd_a2=dpyd_a2, dpyd_stars=DPYD_STARS,

            # UGT1A1
            ugt_mode=ugt_mode, ugt_inputs=ugt_inputs, ugt_geno=ugt_geno,
            ugt_a1=ugt_a1, ugt_a2=ugt_a2, ugt_stars=UGT1A1_STARS,

            # CYP2D6
            cyp_mode=cyp_mode, cyp_inputs=cyp_inputs, cyp_vals=cyp_vals,
            cyp_a1=cyp_a1, cyp_a2=cyp_a2, cyp_stars=CYP_STARS,
        )

    return render_template(
        "index.html",
        # DPYD
        dpyd_inputs={m["column"]: m["options"] for m in MARKERS.get("DPYD", [])},
        dpyd_stars=DPYD_STARS,
        # UGT1A1
        ugt_inputs={m["column"]: m["options"] for m in MARKERS.get("UGT1A1", [])},
        ugt_stars=UGT1A1_STARS,
        # CYP2D6
        cyp_inputs=[(m["column"], m["options"]) for m in MARKERS.get("CYP2D6", [])],
        cyp_stars=CYP_STARS,
    )

@bp.route("/generate", methods=["POST"])
def generate():
    # --- Recoger todo igual que en index POST ---
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

    # ----- DPYD -----
    dpyd_inputs = {m["column"]: m["options"] for m in MARKERS.get("DPYD", [])}
    dpyd_mode = request.form.get("dpyd_mode","markers")
    if dpyd_mode == "diplotype":
        dpyd_a1 = request.form.get("dpyd_a1","*1"); dpyd_a2 = request.form.get("dpyd_a2","*1")
        dpyd_dipl, dpyd_pheno, dpyd_rec, dpyd_poly = dpyd_from_diplotype(dpyd_a1, dpyd_a2)
    else:
        dpyd_vals = {k: request.form.get(k,"-/-") for k in dpyd_inputs.keys()}
        dpyd_dipl, dpyd_pheno, dpyd_rec, dpyd_poly = dpyd_from_markers(dpyd_vals)

    # ----- UGT1A1 -----
    ugt_inputs = {m["column"]: m["options"] for m in MARKERS.get("UGT1A1", [])}
    ugt_mode = request.form.get("ugt_mode","markers")
    if ugt_mode == "diplotype":
        ugt_a1 = request.form.get("ugt_a1","*1"); ugt_a2 = request.form.get("ugt_a2","*1")
        ugt_dipl, ugt_pheno, ugt_rec, ugt_poly = ugt1a1_from_diplotype(ugt_a1, ugt_a2)
    else:
        ugt_col = next(iter(ugt_inputs.keys()))
        ugt_geno = request.form.get(ugt_col, "-/-")
        ugt_dipl, ugt_pheno, ugt_rec, ugt_poly = ugt1a1_from_markers(ugt_geno)

    # ----- CYP2D6 -----
    cyp_inputs = [(m["column"], m["options"]) for m in MARKERS.get("CYP2D6", [])]
    cyp_mode = request.form.get("cyp_mode","diplotype")
    if cyp_mode == "markers":
        cyp_vals = {m[0]: request.form.get(m[0], "-/-") for m in cyp_inputs}
        cyp_dipl, cyp_pheno, cyp_rec, cyp_poly = cyp2d6_from_markers(cyp_vals)
    else:
        cyp_a1 = request.form.get("cyp_a1","*1"); cyp_a2 = request.form.get("cyp_a2","*1")
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
        "sources": {"cpic_url": "https://cpicpgx.org/guidelines/", "dpwgd_doi": "DOI:10.1038/s41431-022-01243-2"},
        "meta": {"sample_code": patient.get("historia") or "—", "request_date": now.strftime("%d/%m/%Y"),
                 "report_date": now.strftime("%d/%m/%Y"), "generated_at": now.strftime("%Y-%m-%d %H:%M"),
                 "last_update": "septiembre 2025", "version": "0.5.0"},
    }

    # Render DOCX temporal
    tmp_docx = os.path.join(tempfile.gettempdir(), f"GenoPilot_tmp_{now.strftime('%H%M%S')}.docx")
    tpl.render(context); tpl.save(tmp_docx)

    # Convertir a PDF si es posible
    pdf_name = f"GenoPilot_{patient.get('full_name','Paciente')}_{now.strftime('%Y%m%d_%H%M')}.pdf"
    pdf_path = os.path.join(REPORTS_DIR, pdf_name)
    try:
        from docx2pdf import convert
        convert(tmp_docx, pdf_path)   # requiere MS Word en Windows
        return send_file(pdf_path, as_attachment=True, download_name=pdf_name, mimetype="application/pdf")
    except Exception:
        # fallback DOCX
        return send_file(tmp_docx, as_attachment=True,
                         download_name=pdf_name.replace(".pdf",".docx"),
                         mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
