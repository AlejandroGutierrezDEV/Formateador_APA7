# -*- coding: utf-8 -*-
"""
Formateador automático de documentos Word (.docx) según la norma APA 7.ª edición.

Ejecución:  python -m streamlit run app.py
Requisitos: pip install python-docx streamlit
"""

import io
import re
import unicodedata
from datetime import date

import streamlit as st
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run

# ---------------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------------
FUENTE = "Times New Roman"
TAMANO = 12
MARGEN = Cm(2.54)
SANGRIA = Cm(1.27)
TEXTO_TOC = "Tabla de contenido"
NIVELES_TOC = "1-3"  # Los niveles 4 y 5 son títulos "en línea" (van dentro del párrafo)

# APA 7: "Figura 1"/"Tabla 1" en NEGRITA y el título en CURSIVA en la línea siguiente.
# Pon True si prefieres la etiqueta ("Figura X") también en cursiva.
ETIQUETAS_EN_CURSIVA = False

ALIN_IZQ = WD_ALIGN_PARAGRAPH.LEFT
ALIN_CEN = WD_ALIGN_PARAGRAPH.CENTER
ALIN_DER = WD_ALIGN_PARAGRAPH.RIGHT

W_P = qn("w:p")
W_TBL = qn("w:tbl")

RE_ESTILO_TITULO = re.compile(r"^(?:heading|t[ií]tulo|titulo)\s*([1-9])$", re.I)
RE_REFERENCIAS = re.compile(
    r"^\s*(?:referencias(?:\s+bibliogr[aá]ficas)?|bibliograf[ií]a|references|bibliography)\s*:?\s*$",
    re.I,
)
RE_ETIQUETA = {
    "Figura": re.compile(r"^\s*(?:figura|figure|fig\.)\s*\d+\s*(?:[.:\-–—]\s*|$)", re.I),
    "Tabla": re.compile(r"^\s*(?:tabla|table)\s*\d+\s*(?:[.:\-–—]\s*|$)", re.I),
}
RE_ETIQUETA_LAX = {
    "Figura": re.compile(r"^\s*(?:figura|figure|fig\.)\s*\d+\s*[.:\-–—]?\s*", re.I),
    "Tabla": re.compile(r"^\s*(?:tabla|table)\s*\d+\s*[.:\-–—]?\s*", re.I),
}
RE_ETIQUETA_SOLA = re.compile(r"^\s*(?:figura|figure|fig\.|tabla|table)\s*\d+\s*[.:]?\s*$", re.I)
RE_NOTA = re.compile(r"^\s*(?:nota|note)\s*[.:]\s*", re.I)

# Citas parentéticas básicas: (Autor 2020) -> (Autor, 2020) ; (Autor, 2020: 15) -> (Autor, 2020, p. 15)
RE_CITA_SIN_COMA = re.compile(
    r"\(([^()\d;]*?[A-Za-zÁÉÍÓÚÜÑáéíóúüñ.])\s+((?:19|20)\d{2}[a-z]?)(?=[,;)])"
)
RE_CITA_PAGINA = re.compile(
    r"((?:19|20)\d{2}[a-z]?)\s*:\s*(\d+)(?:\s*[-–]\s*(\d+))?(?=\s*\))"
)

MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]


# ---------------------------------------------------------------------------
# UTILIDADES GENERALES
# ---------------------------------------------------------------------------
def fecha_es(d):
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def es_parrafo(el):
    return el is not None and el.tag == W_P


def texto_el(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def nombre_estilo(p):
    try:
        return (p.style.name or "") if p.style is not None else ""
    except Exception:
        return ""


def nivel_titulo(p):
    m = RE_ESTILO_TITULO.match(nombre_estilo(p).strip())
    return min(int(m.group(1)), 5) if m else 0


def tiene_imagen(p_el):
    return next(p_el.iter(qn("w:drawing"), qn("w:pict")), None) is not None


def es_vacio(p_el):
    """Párrafo sin texto ni contenido especial (imágenes, saltos, campos, ecuaciones...)."""
    if texto_el(p_el).strip():
        return False
    especiales = (
        qn("w:drawing"), qn("w:pict"), qn("w:br"), qn("w:fldChar"), qn("w:fldSimple"),
        qn("w:sectPr"), qn("w:object"), qn("w:bookmarkStart"), qn("m:oMath"), qn("m:oMathPara"),
    )
    return next(p_el.iter(*especiales), None) is None


def runs_de(p):
    return [Run(r, p) for r in p._p.iter(qn("w:r"))]


def limpiar_contenido(p_el):
    for hijo in list(p_el):
        if hijo.tag != qn("w:pPr"):
            p_el.remove(hijo)


def poner_en_orden(padre, nuevo, sucesores):
    """Inserta 'nuevo' respetando el orden del esquema XML de Word."""
    for etiqueta in sucesores:
        ref = padre.find(qn(etiqueta))
        if ref is not None:
            ref.addprevious(nuevo)
            return
    padre.append(nuevo)


def nuevo_parrafo(doc, ref_el, antes=True):
    nuevo = OxmlElement("w:p")
    if antes:
        ref_el.addprevious(nuevo)
    else:
        ref_el.addnext(nuevo)
    return Paragraph(nuevo, doc._body)


# ---------------------------------------------------------------------------
# FORMATO DE FUENTE Y PÁRRAFO
# ---------------------------------------------------------------------------
def fijar_fuente_rpr(rpr):
    """Fuerza Times New Roman (y elimina las fuentes de tema) en un elemento rPr."""
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rstyle = rpr.find(qn("w:rStyle"))
        if rstyle is not None:
            rstyle.addnext(rf)
        else:
            rpr.insert(0, rf)
    for atributo in list(rf.attrib):
        del rf.attrib[atributo]
    for atributo in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(atributo), FUENTE)


def formatear_runs(p, negro=False):
    for run in runs_de(p):
        fijar_fuente_rpr(run._r.get_or_add_rPr())
        run.font.size = Pt(TAMANO)
        if negro:
            run.font.color.rgb = RGBColor(0, 0, 0)


def formato_parrafo(p, alin=None, primera=None, izq=None, interlineado=2.0, mantener=None):
    """Interlineado y espaciado APA. Los parámetros en None no se modifican."""
    pf = p.paragraph_format
    if alin is not None:
        p.alignment = alin
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = interlineado
    pf.right_indent = Cm(0)
    pf.widow_control = True
    if primera is not None:
        pf.first_line_indent = primera
    if izq is not None:
        pf.left_indent = izq
    if mantener is not None:
        pf.keep_with_next = mantener

    ppr = p._p.pPr
    if ppr is not None:
        sp = ppr.find(qn("w:spacing"))
        if sp is not None:
            for a in ("w:beforeLines", "w:afterLines", "w:beforeAutospacing", "w:afterAutospacing"):
                sp.attrib.pop(qn(a), None)
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            for a in ("w:firstLineChars", "w:hangingChars", "w:leftChars", "w:rightChars",
                      "w:startChars", "w:endChars"):
                ind.attrib.pop(qn(a), None)
            if izq is not None:
                for a in ("w:start", "w:end"):
                    ind.attrib.pop(qn(a), None)


def agregar_campo(p, instruccion, resultado=""):
    """Añade un campo nativo de Word (PAGE, TOC...) a un párrafo."""
    def run_con(elemento):
        r = p.add_run()
        r._r.append(elemento)

    inicio = OxmlElement("w:fldChar")
    inicio.set(qn("w:fldCharType"), "begin")
    run_con(inicio)

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruccion} "
    run_con(instr)

    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    run_con(sep)

    p.add_run(resultado)

    fin = OxmlElement("w:fldChar")
    fin.set(qn("w:fldCharType"), "end")
    run_con(fin)


# ---------------------------------------------------------------------------
# ESTILOS, PÁGINA Y ENCABEZADO
# ---------------------------------------------------------------------------
def asegurar_estilos_titulo(doc):
    """Garantiza que existan los estilos Heading 1-5 (necesarios para el índice)."""
    for n in range(1, 6):
        nombre = f"Heading {n}"
        try:
            doc.styles[nombre]
        except KeyError:
            estilo = doc.styles.add_style(nombre, WD_STYLE_TYPE.PARAGRAPH)
            estilo.base_style = doc.styles["Normal"]
            ppr = estilo.element.get_or_add_pPr()
            nivel = OxmlElement("w:outlineLvl")
            nivel.set(qn("w:val"), str(n - 1))
            ppr.append(nivel)


def normalizar_estilos(doc):
    # Valores por defecto del documento
    dd = doc.styles.element.find(qn("w:docDefaults"))
    if dd is not None:
        rpr = dd.find(f"{qn('w:rPrDefault')}/{qn('w:rPr')}")
        if rpr is not None:
            fijar_fuente_rpr(rpr)
            rf = rpr.find(qn("w:rFonts"))
            anterior = rf
            for etiqueta in ("w:sz", "w:szCs"):
                el = rpr.find(qn(etiqueta))
                if el is None:
                    el = OxmlElement(etiqueta)
                    anterior.addnext(el)
                el.set(qn("w:val"), str(TAMANO * 2))
                anterior = el

    for estilo in doc.styles:
        try:
            if estilo.type not in (WD_STYLE_TYPE.PARAGRAPH, WD_STYLE_TYPE.CHARACTER):
                continue
            fijar_fuente_rpr(estilo.element.get_or_add_rPr())
            if estilo.type == WD_STYLE_TYPE.PARAGRAPH:
                estilo.font.size = Pt(TAMANO)
                nombre = (estilo.name or "").lower()
                if nombre.startswith("heading") or nombre in ("title", "subtitle"):
                    estilo.font.color.rgb = RGBColor(0, 0, 0)
                    estilo.font.bold = False   # la negrita/cursiva se aplica directamente
                    estilo.font.italic = False
                    estilo.paragraph_format.space_before = Pt(0)
                    estilo.paragraph_format.space_after = Pt(0)
                    estilo.paragraph_format.line_spacing = 2.0
        except Exception:
            continue  # estilos latentes o poco habituales

    normal = doc.styles["Normal"]
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 2.0


def configurar_paginas(doc):
    """Márgenes de 2,54 cm y número de página a la derecha del encabezado, en todas las páginas."""
    doc.settings.odd_and_even_pages_header_footer = False
    for sec in doc.sections:
        sec.top_margin = MARGEN
        sec.bottom_margin = MARGEN
        sec.left_margin = MARGEN
        sec.right_margin = MARGEN
        sec.gutter = Cm(0)
        sec.header_distance = Cm(1.25)
        sec.different_first_page_header_footer = False

        encabezado = sec.header
        encabezado.is_linked_to_previous = False
        for extra in encabezado.paragraphs[1:]:
            extra._p.getparent().remove(extra._p)
        p = encabezado.paragraphs[0]
        limpiar_contenido(p._p)
        formato_parrafo(p, ALIN_DER, primera=Cm(0), izq=Cm(0), interlineado=1.0)
        agregar_campo(p, "PAGE", "1")
        formatear_runs(p)


def pedir_actualizar_campos(doc):
    """Word ofrecerá actualizar los campos (índice, paginación) al abrir el archivo."""
    ajustes = doc.settings.element
    if ajustes.find(qn("w:updateFields")) is None:
        el = OxmlElement("w:updateFields")
        el.set(qn("w:val"), "true")
        ajustes.append(el)


# ---------------------------------------------------------------------------
# LIMPIEZA PREVIA
# ---------------------------------------------------------------------------
def extraer_titulo_documento(doc):
    """Si el documento tiene un párrafo de estilo 'Title', se usa en la portada y se retira."""
    for p in doc.paragraphs[:15]:
        if nombre_estilo(p).lower() in ("title", "título", "titulo") and texto_el(p._p).strip():
            titulo = texto_el(p._p).strip()
            p._p.getparent().remove(p._p)
            return titulo
    return ""


def quitar_tabla_contenido_previa(doc):
    """Elimina un índice previo para no duplicarlo."""
    body = doc.element.body
    for sdt in list(body.findall(qn("w:sdt"))):
        galeria = sdt.find(".//" + qn("w:docPartGallery"))
        if galeria is not None and "table of contents" in (galeria.get(qn("w:val")) or "").lower():
            body.remove(sdt)
    for p in list(doc.paragraphs):
        nombre = nombre_estilo(p).lower().strip()
        if re.match(r"^toc\s*\d", nombre) or nombre in ("toc heading", "table of contents"):
            p._p.getparent().remove(p._p)


# ---------------------------------------------------------------------------
# REFERENCIAS Y CITAS
# ---------------------------------------------------------------------------
def clave_orden(texto):
    t = unicodedata.normalize("NFKD", texto.lower())
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"^[^a-z0-9]+", "", t)


def ordenar_referencias(doc, informe):
    cabecera = None
    for p in reversed(list(doc.paragraphs)):
        if RE_REFERENCIAS.match(texto_el(p._p)):
            cabecera = p
            break
    if cabecera is None:
        informe["avisos"].append(
            "No se encontró una sección titulada «Referencias» (o «Bibliografía»): "
            "no se aplicó la sangría francesa ni el orden alfabético."
        )
        return

    hijos = []
    el = cabecera._p.getnext()
    while el is not None and el.tag == W_P:
        if nivel_titulo(Paragraph(el, doc._body)):
            break
        hijos.append(el)
        el = el.getnext()

    con_contenido = [h for h in hijos if not es_vacio(h)]
    for h in hijos:
        h.getparent().remove(h)
    con_contenido.sort(key=lambda h: clave_orden(texto_el(h)))
    ancla = cabecera._p
    for h in con_contenido:
        ancla.addnext(h)
        ancla = h


def normalizar_citas(p, informe):
    def repl_pagina(m):
        if m.group(3):
            return f"{m.group(1)}, pp. {m.group(2)}–{m.group(3)}"
        return f"{m.group(1)}, p. {m.group(2)}"

    for t in p._p.iter(qn("w:t")):
        original = t.text or ""
        nuevo = RE_CITA_SIN_COMA.sub(r"(\1, \2", original)
        nuevo = RE_CITA_PAGINA.sub(repl_pagina, nuevo)
        if nuevo != original:
            t.text = nuevo
            informe["citas"] += 1


# ---------------------------------------------------------------------------
# FIGURAS Y TABLAS: ETIQUETAS Y NOTAS
# ---------------------------------------------------------------------------
def _titulo_si_etiqueta(doc, el, tipo):
    """Si 'el' es una etiqueta de figura/tabla devuelve su título (puede ser ''), si no None."""
    if not es_parrafo(el):
        return None
    t = texto_el(el).strip()
    m = RE_ETIQUETA[tipo].match(t)
    if m:
        return t[m.end():].strip()
    if nombre_estilo(Paragraph(el, doc._body)).lower() == "caption":
        m = RE_ETIQUETA_LAX[tipo].match(t)
        if m:
            return t[m.end():].strip()
    return None


def _es_objeto(el):
    return el is not None and (el.tag == W_TBL or (el.tag == W_P and tiene_imagen(el)))


def buscar_etiqueta_existente(doc, el, tipo):
    ant = el.getprevious()
    titulo = _titulo_si_etiqueta(doc, ant, tipo)
    if titulo is not None:
        return [ant], titulo
    if es_parrafo(ant):  # patrón "Figura 1" + "Título" encima del objeto
        ant2 = ant.getprevious()
        if (es_parrafo(ant2) and RE_ETIQUETA_SOLA.match(texto_el(ant2))
                and _titulo_si_etiqueta(doc, ant2, tipo) is not None):
            return [ant2, ant], texto_el(ant).strip()
    sig = el.getnext()
    titulo = _titulo_si_etiqueta(doc, sig, tipo)
    if (titulo is not None and not RE_ETIQUETA_SOLA.match(texto_el(sig))
            and not _es_objeto(sig.getnext())):
        return [sig], titulo
    return [], ""


def escribir_nota(p, resto):
    limpiar_contenido(p._p)
    r = p.add_run("Nota.")
    r.italic = True
    if resto:
        p.add_run(" " + resto)


def etiquetar_objetos(doc, insertar_faltantes, informe):
    body = doc.element.body
    objetos = []
    for el in body.iterchildren():
        if el.tag == W_P and tiene_imagen(el):
            objetos.append((el, "Figura"))
        elif el.tag == W_TBL:
            objetos.append((el, "Tabla"))

    contador = {"Figura": 0, "Tabla": 0}
    for el, tipo in objetos:
        contador[tipo] += 1
        n = contador[tipo]
        informe["figuras" if tipo == "Figura" else "tablas"] += 1

        quitados, titulo = buscar_etiqueta_existente(doc, el, tipo)
        for q in quitados:
            q.getparent().remove(q)

        if quitados or insertar_faltantes:
            etiqueta = nuevo_parrafo(doc, el, antes=True)
            run = etiqueta.add_run(f"{tipo} {n}")
            if ETIQUETAS_EN_CURSIVA:
                run.italic = True
            else:
                run.bold = True
            nombre = nuevo_parrafo(doc, el, antes=True)
            nombre.add_run(titulo or f"[Título de la {tipo.lower()}]").italic = True

        sig = el.getnext()
        if es_parrafo(sig) and RE_NOTA.match(texto_el(sig)):
            resto = RE_NOTA.sub("", texto_el(sig), count=1).strip()
            escribir_nota(Paragraph(sig, doc._body), resto)
        elif insertar_faltantes:
            escribir_nota(nuevo_parrafo(doc, el, antes=False),
                          "[Indique la fuente o añada una aclaración.]")


def formatear_tablas(doc):
    sucesores_tbl = ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
                     "w:tblCaption", "w:tblDescription")
    sucesores_tc = ("w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
                    "w:tcFitText", "w:vAlign", "w:hideMark")

    def borde(lado, visible):
        el = OxmlElement(f"w:{lado}")
        if visible:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "4")      # 0,5 pt
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")
        else:
            el.set(qn("w:val"), "nil")
        return el

    for tabla in doc.tables:
        tbl_pr = tabla._tbl.tblPr
        for etiqueta in ("w:tblStyle", "w:tblBorders"):
            previo = tbl_pr.find(qn(etiqueta))
            if previo is not None:
                tbl_pr.remove(previo)

        bordes = OxmlElement("w:tblBorders")
        for lado, visible in (("top", True), ("left", False), ("bottom", True),
                              ("right", False), ("insideH", False), ("insideV", False)):
            bordes.append(borde(lado, visible))
        poner_en_orden(tbl_pr, bordes, sucesores_tbl)
        tabla.alignment = WD_TABLE_ALIGNMENT.CENTER

        for i_fila, fila in enumerate(tabla.rows):
            for celda in fila.cells:
                tc_pr = celda._tc.get_or_add_tcPr()
                for etiqueta in ("w:tcBorders", "w:shd"):
                    previo = tc_pr.find(qn(etiqueta))
                    if previo is not None:
                        tc_pr.remove(previo)
                if i_fila == 0:  # línea bajo los encabezados de columna
                    tc_bordes = OxmlElement("w:tcBorders")
                    tc_bordes.append(borde("bottom", True))
                    poner_en_orden(tc_pr, tc_bordes, sucesores_tc)
                for par in celda.paragraphs:
                    formato_parrafo(par, None, primera=Cm(0), interlineado=1.0)
                    formatear_runs(par)
                    if i_fila == 0:
                        for run in runs_de(par):
                            run.bold = True


# ---------------------------------------------------------------------------
# CUERPO DEL DOCUMENTO
# ---------------------------------------------------------------------------
def formatear_cuerpo(doc, cfg, informe):
    parrafos = list(doc.paragraphs)
    en_refs = False
    prev_etiqueta = False
    i = 0
    while i < len(parrafos):
        p = parrafos[i]
        if p._p.getparent() is None:
            i += 1
            continue

        era_etiqueta = prev_etiqueta
        prev_etiqueta = False

        if es_vacio(p._p):
            p._p.getparent().remove(p._p)
            i += 1
            continue

        texto = texto_el(p._p).strip()
        nivel = nivel_titulo(p)

        # "Referencias" sin estilo de título -> se trata como título de nivel 1
        if RE_REFERENCIAS.match(texto) and nivel in (0, 1):
            if nivel == 0:
                p.style = doc.styles["Heading 1"]
            nivel = 1

        # --- Títulos ---------------------------------------------------------
        if nivel:
            informe["titulos"][nivel] += 1
            if nivel <= 1:
                en_refs = bool(RE_REFERENCIAS.match(texto))
            negrita, cursiva = True, nivel in (3, 5)
            formato_parrafo(
                p,
                ALIN_CEN if nivel == 1 else ALIN_IZQ,
                primera=SANGRIA if nivel >= 4 else Cm(0),
                izq=Cm(0),
                mantener=nivel <= 3,
            )
            if nivel >= 4:  # punto final obligatorio
                ultimos = list(p._p.iter(qn("w:t")))
                if ultimos:
                    t = ultimos[-1]
                    base = (t.text or "").rstrip()
                    if base and base[-1] not in ".?!":
                        t.text = base + "."
            for run in runs_de(p):
                run.bold = negrita
                run.italic = cursiva
            formatear_runs(p, negro=True)

            # Niveles 4 y 5: el texto del párrafo siguiente continúa en la misma línea
            if nivel >= 4 and i + 1 < len(parrafos):
                sig = parrafos[i + 1]
                st_sig = texto_el(sig._p).strip()
                if (p._p.getnext() is sig._p and st_sig and not nivel_titulo(sig)
                        and not tiene_imagen(sig._p) and not RE_REFERENCIAS.match(st_sig)):
                    espacio = p.add_run(" ")
                    espacio.bold = False
                    espacio.italic = False
                    for hijo in list(sig._p):
                        if hijo.tag != qn("w:pPr"):
                            p._p.append(hijo)
                    sig._p.getparent().remove(sig._p)
                    formatear_runs(p)
                    informe["fusionados"] += 1
                    if cfg["citas"]:
                        normalizar_citas(p, informe)
                    i += 2
                    continue
            i += 1
            continue

        # --- Sección de referencias -----------------------------------------
        if en_refs:
            formato_parrafo(p, ALIN_IZQ, primera=Cm(-1.27), izq=SANGRIA)
            formatear_runs(p)
            informe["referencias"] += 1
            i += 1
            continue

        # --- Imágenes: centradas ---------------------------------------------
        if tiene_imagen(p._p):
            # Interlineado sencillo: con doble, Word duplicaría la altura de la línea de la imagen
            formato_parrafo(p, ALIN_CEN, primera=Cm(0), izq=Cm(0), interlineado=1.0, mantener=True)
            i += 1
            continue

        # --- Etiquetas, títulos de figura/tabla y notas ----------------------
        es_etiqueta = bool(RE_ETIQUETA_SOLA.match(texto))
        if (es_etiqueta or era_etiqueta or RE_NOTA.match(texto)
                or nombre_estilo(p).lower() == "caption"):
            formato_parrafo(p, ALIN_IZQ, primera=Cm(0), izq=Cm(0),
                            mantener=es_etiqueta or era_etiqueta)
            formatear_runs(p)
            prev_etiqueta = es_etiqueta
            i += 1
            continue

        # --- Citas en bloque (40 palabras o más) -----------------------------
        if (nombre_estilo(p).lower() in ("quote", "intense quote", "block text")
                and len(texto.split()) >= 40):
            formato_parrafo(p, ALIN_IZQ, primera=Cm(0), izq=SANGRIA)
            formatear_runs(p)
            i += 1
            continue

        # --- Listas: se respeta su sangría propia ----------------------------
        ppr = p._p.pPr
        if (ppr is not None and ppr.numPr is not None) or nombre_estilo(p).lower().startswith("list"):
            formato_parrafo(p, ALIN_IZQ)
            formatear_runs(p)
            if cfg["citas"]:
                normalizar_citas(p, informe)
            i += 1
            continue

        # --- Párrafo normal del cuerpo ---------------------------------------
        formato_parrafo(p, ALIN_IZQ, primera=SANGRIA, izq=Cm(0))
        formatear_runs(p)
        if cfg["citas"]:
            normalizar_citas(p, informe)
        i += 1


# ---------------------------------------------------------------------------
# PORTADA E ÍNDICE
# ---------------------------------------------------------------------------
def salto_antes(doc, elemento):
    if elemento.tag == W_P:
        Paragraph(elemento, doc._body).paragraph_format.page_break_before = True
    else:
        p = nuevo_parrafo(doc, elemento, antes=True)
        p.paragraph_format.page_break_before = True


def crear_portada(doc, ancla, datos):
    def linea(texto="", negrita=False):
        p = doc.add_paragraph()
        if texto:
            run = p.add_run(texto)
            run.bold = negrita
        formato_parrafo(p, ALIN_CEN, primera=Cm(0), izq=Cm(0))
        formatear_runs(p)
        ancla.addprevious(p._p)

    for _ in range(3):
        linea()
    linea(datos["titulo"], negrita=True)
    linea()
    for clave in ("autor", "institucion", "asignatura", "docente", "fecha"):
        if datos.get(clave):
            linea(datos[clave])


def crear_toc(doc, ancla, con_salto):
    titulo = doc.add_paragraph()
    run = titulo.add_run(TEXTO_TOC)
    run.bold = True
    formato_parrafo(titulo, ALIN_CEN, primera=Cm(0), izq=Cm(0), mantener=True)
    formatear_runs(titulo)
    if con_salto:
        titulo.paragraph_format.page_break_before = True
    ancla.addprevious(titulo._p)

    campo = doc.add_paragraph()
    agregar_campo(
        campo,
        f'TOC \\o "{NIVELES_TOC}" \\h \\z \\u',
        "Haz clic derecho aquí y elige «Actualizar campo» (o pulsa F9) para generar el índice.",
    )
    formato_parrafo(campo, ALIN_IZQ, primera=Cm(0), izq=Cm(0))
    formatear_runs(campo)
    ancla.addprevious(campo._p)


def crear_titulo_cuerpo(doc, ancla, titulo):
    p = doc.add_paragraph()
    p.add_run(titulo).bold = True
    formato_parrafo(p, ALIN_CEN, primera=Cm(0), izq=Cm(0), mantener=True)
    formatear_runs(p)
    p.paragraph_format.page_break_before = True
    ancla.addprevious(p._p)


# ---------------------------------------------------------------------------
# PROCESO COMPLETO
# ---------------------------------------------------------------------------
def formatear_documento(contenido, cfg):
    doc = Document(io.BytesIO(contenido))
    informe = {
        "titulos": {n: 0 for n in range(1, 6)},
        "fusionados": 0, "tablas": 0, "figuras": 0,
        "referencias": 0, "citas": 0, "avisos": [],
    }

    asegurar_estilos_titulo(doc)
    normalizar_estilos(doc)
    configurar_paginas(doc)

    titulo_doc = extraer_titulo_documento(doc)
    quitar_tabla_contenido_previa(doc)
    if cfg["referencias"]:
        ordenar_referencias(doc, informe)
    etiquetar_objetos(doc, cfg["etiquetas"], informe)
    formatear_tablas(doc)
    formatear_cuerpo(doc, cfg, informe)

    if sum(informe["titulos"].values()) == 0:
        informe["avisos"].append(
            "No se detectaron títulos. Usa los estilos «Título 1–5» de Word para que se "
            "reconozcan los niveles y se genere el índice."
        )

    # Portada, índice y título repetido al inicio del cuerpo
    titulo = (cfg["titulo"] or titulo_doc).strip()
    primer = doc.element.body[0]
    hay_portada = cfg["portada"]
    hay_frontal = False

    if hay_portada:
        if not titulo:
            informe["avisos"].append("Portada sin título: escribe uno en el formulario.")
            titulo = "[Título del trabajo]"
        datos = dict(cfg, titulo=titulo)
        crear_portada(doc, primer, datos)
        hay_frontal = True
    if cfg["toc"]:
        crear_toc(doc, primer, con_salto=hay_portada)
        hay_frontal = True
    if hay_portada:
        crear_titulo_cuerpo(doc, primer, titulo)
    elif hay_frontal:
        salto_antes(doc, primer)

    if cfg["actualizar"]:
        pedir_actualizar_campos(doc)

    salida = io.BytesIO()
    doc.save(salida)
    return salida.getvalue(), informe


# ---------------------------------------------------------------------------
# INTERFAZ STREAMLIT
# ---------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Formateador APA 7", page_icon="📄", layout="centered")
    st.title("📄 Formateador automático APA 7")
    st.caption("Sube un documento Word, completa los datos de la portada y descarga el resultado. "
               "Todo se procesa en tu equipo.")

    if "resultado" not in st.session_state:
        st.session_state.resultado = None

    archivo = st.file_uploader("Documento Word (.docx)", type=["docx"])

    st.subheader("Datos de la portada")
    c1, c2 = st.columns(2)
    titulo = c1.text_input("Título del trabajo")
    autor = c2.text_input("Autor/a")
    institucion = c1.text_input("Institución (facultad, universidad)")
    asignatura = c2.text_input("Asignatura (código y nombre)")
    docente = c1.text_input("Docente (opcional)")
    fecha = c2.date_input("Fecha de entrega", value=date.today())

    with st.expander("Opciones"):
        portada = st.checkbox("Generar portada", value=True)
        toc = st.checkbox("Insertar tabla de contenido (campo TOC)", value=True)
        referencias = st.checkbox("Ordenar referencias y aplicar sangría francesa", value=True)
        etiquetas = st.checkbox("Añadir etiqueta, título y nota a figuras y tablas si faltan", value=True)
        citas = st.checkbox("Normalizar citas autor-fecha básicas: (Autor 2020) → (Autor, 2020)", value=True)
        actualizar = st.checkbox("Pedir a Word que actualice el índice al abrir el archivo", value=True)

    if st.button("Formatear documento", type="primary", disabled=archivo is None):
        cfg = {
            "titulo": titulo, "autor": autor, "institucion": institucion,
            "asignatura": asignatura, "docente": docente, "fecha": fecha_es(fecha),
            "portada": portada, "toc": toc, "referencias": referencias,
            "etiquetas": etiquetas, "citas": citas, "actualizar": actualizar,
        }
        try:
            with st.spinner("Aplicando normas APA 7..."):
                datos, informe = formatear_documento(archivo.getvalue(), cfg)
            st.session_state.resultado = {
                "origen": archivo.name, "datos": datos, "informe": informe,
            }
        except Exception as e:
            st.session_state.resultado = None
            st.error("No se pudo procesar el archivo. ¿Es un .docx válido y no está protegido?")
            st.exception(e)

    res = st.session_state.resultado
    if res and archivo is not None and res["origen"] == archivo.name:
        informe = res["informe"]
        st.success("Documento formateado correctamente.")
        nombre_salida = archivo.name.rsplit(".", 1)[0] + "_APA7.docx"
        st.download_button(
            "⬇️ Descargar documento APA 7",
            data=res["datos"],
            file_name=nombre_salida,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        niveles = ", ".join(f"N{n}: {c}" for n, c in informe["titulos"].items() if c) or "ninguno"
        st.markdown(
            f"- **Títulos detectados:** {niveles} (títulos en línea fusionados: {informe['fusionados']})\n"
            f"- **Tablas:** {informe['tablas']} · **Figuras:** {informe['figuras']}\n"
            f"- **Referencias formateadas:** {informe['referencias']}\n"
            f"- **Citas corregidas:** {informe['citas']}"
        )
        for aviso in informe["avisos"]:
            st.warning(aviso)
        st.info("Al abrir el archivo en Word, acepta la actualización de campos o haz clic derecho "
                "en el índice → «Actualizar campo» → «Actualizar toda la tabla».")


if __name__ == "__main__":
    main()