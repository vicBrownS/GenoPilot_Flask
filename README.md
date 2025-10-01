# GenoPilot — Flask (demo docente)

**Objetivo**: App Flask con **entrada manual** de DPYD, UGT1A1 y CYP2D6 que genera un **informe DOCX** con recomendaciones (ClinPGX).  
Los **datos de mapeo** se extraen de `PharmPrev_macro_base_v2.xlsm` (sheets: Raw/Maps/CYP2D6_Pheno/Recs).


## Estructura
```
PharmPrev_Flask/
  app/
    __init__.py
    routes.py
    templates/
      base.html
      index.html
      preview.html
      report_template.docx
  data/
    markers.json
    cyp2d6_stars.json
    cyp2d6_pheno.json
    recs.json
  reports/
  run.py
  requirements.txt
```

## Ejecutar
```bash
python -m venv .venv && .venv\Scripts\activate
pip install -r requirements.txt
set FLASK_APP=run.py
flask run
# o: python run.py
```

## Generación de informe
- El botón **Generar DOCX** descarga el informe en `reports/`.
- Si quieres PDF: instala Word en Windows y usa `docx2pdf` (añadir paquete y conversión).

## Notas
- **DPYD** y **UGT1A1** se introducen por **marcador** (genotipo) siguiendo `Maps`.
- **CYP2D6** de momento se introduce **por diplotipo** (par de estrellas). Si quieres, ampliamos a entrada por marcadores.
