import streamlit as st
import os
import re
import pandas as pd
import requests
import pdfplumber
import plotly.graph_objects as go
from io import BytesIO
from PIL import Image

# ============================================================
# FUENTE DE DATOS — Google Sheets
# ============================================================
SHEET_ID  = "1AhCKmAg7mak8qzNsI4FyKxVIL-Amiahc"
XLSX_URL  = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
CSV_URL   = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
os.makedirs(ASSETS_DIR, exist_ok=True)

# ============================================================
# MAPEO DE CATEGORÍAS
# ============================================================
TIPO_A_CATEGORIA = {
    "tanque":          "Tanques",
    "bomba":           "Bombas",
    "intercambiador":  "Intercambiadores",
    "filtro":          "Filtros",
    "compresor":       "Compresores",
}
ICONOS_CATEGORIA = {
    "Tanques":          "🛢️ Tanques",
    "Bombas":           "⚙️ Bombas",
    "Intercambiadores": "🔥 Intercambiadores de Calor",
    "Filtros":          "🔵 Filtros",
    "Compresores":      "💨 Compresores",
}
TODAS_CATEGORIAS = list(ICONOS_CATEGORIA.keys())

# ============================================================
# CARGA DE DATOS DESDE GOOGLE SHEETS (todas las pestañas)
# ============================================================

# Mapeo flexible de nombres de pestaña → categoría interna
_NOMBRE_A_CAT: dict[str, str] = {
    "tanque": "Tanques", "tanques": "Tanques",
    "bomba":  "Bombas",  "bombas":  "Bombas",
    "intercambiador":  "Intercambiadores",
    "intercambiadores":"Intercambiadores",
    "filtro": "Filtros", "filtros": "Filtros",
    "compresor":  "Compresores",
    "compresores":"Compresores",
}

def _limpiar_df(df: pd.DataFrame) -> pd.DataFrame:
    """Limpia un DataFrame crudo: strip de columnas y valores, elimina filas vacías."""
    df = df.copy().fillna("").astype(str)
    df.columns = [str(c).strip() for c in df.columns]
    for col in df.columns:
        df[col] = df[col].str.strip()
    df = df[df.apply(lambda r: any(v not in ("", "nan") for v in r), axis=1)]
    return df.reset_index(drop=True)

def _asignar_a_categorias(df: pd.DataFrame, hint_cat: str | None = None) -> dict[str, pd.DataFrame]:
    """Distribuye filas de un DataFrame a categorías según columna 'Equipo' o hint."""
    datos: dict[str, pd.DataFrame] = {cat: pd.DataFrame() for cat in TODAS_CATEGORIAS}
    if "Equipo" in df.columns:
        for tipo, cat in TIPO_A_CATEGORIA.items():
            mask = df["Equipo"].str.lower().str.strip() == tipo
            sub  = df[mask].reset_index(drop=True)
            if not sub.empty:
                datos[cat] = pd.concat([datos[cat], sub], ignore_index=True) if not datos[cat].empty else sub
    elif hint_cat and hint_cat in TODAS_CATEGORIAS:
        datos[hint_cat] = df
    else:
        datos["Tanques"] = df
    return datos

@st.cache_data(ttl=300)
def cargar_datos():
    datos: dict[str, pd.DataFrame] = {cat: pd.DataFrame() for cat in TODAS_CATEGORIAS}
    mensajes: list[str] = []

    # ── Intento 1: XLSX (todas las pestañas) ──────────────────
    try:
        resp = requests.get(XLSX_URL, timeout=30)
        resp.raise_for_status()
        hojas = pd.read_excel(BytesIO(resp.content), sheet_name=None, dtype=str)
        for nombre_hoja, df_raw in hojas.items():
            df = _limpiar_df(df_raw)
            if df.empty:
                continue
            nombre_lower = nombre_hoja.lower().strip()
            hint_cat = _NOMBRE_A_CAT.get(nombre_lower)
            parcial  = _asignar_a_categorias(df, hint_cat)
            for cat, sub in parcial.items():
                if not sub.empty:
                    datos[cat] = pd.concat([datos[cat], sub], ignore_index=True) if not datos[cat].empty else sub
        # Deduplicar por Tag dentro de cada categoría
        for cat in datos:
            if not datos[cat].empty and "Tag" in datos[cat].columns:
                datos[cat] = datos[cat].drop_duplicates(subset=["Tag"]).reset_index(drop=True)
        total = sum(len(v) for v in datos.values())
        mensajes.append(f"✅ {total} equipos cargados desde {len(hojas)} hoja(s)")
        return datos, " | ".join(mensajes)
    except Exception as e_xlsx:
        mensajes.append(f"⚠️ XLSX falló ({e_xlsx}), intentando CSV…")

    # ── Intento 2: CSV (solo primera hoja) ────────────────────
    try:
        resp = requests.get(CSV_URL, timeout=20)
        resp.raise_for_status()
        df_raw = pd.read_csv(BytesIO(resp.content), dtype=str)
        df = _limpiar_df(df_raw)
        parcial = _asignar_a_categorias(df)
        for cat, sub in parcial.items():
            if not sub.empty:
                datos[cat] = sub
        total = sum(len(v) for v in datos.values())
        mensajes.append(f"✅ {total} equipos cargados (CSV, 1 hoja)")
        return datos, " | ".join(mensajes)
    except Exception as e_csv:
        mensajes.append(f"⚠️ CSV también falló: {e_csv}")

    return datos, " | ".join(mensajes)


# ============================================================
# EXTRACCIÓN DE PDF — texto + tablas
# ============================================================
def _get_drive_download_url(url: str) -> str | None:
    if not url or url == "nan":
        return None
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
    return f"https://drive.google.com/uc?export=download&id={match.group(1)}" if match else None

@st.cache_data(ttl=600, show_spinner=False)
def extraer_contenido_pdf(plano_url: str) -> dict:
    dl_url = _get_drive_download_url(plano_url)
    if not dl_url:
        return {"texto": "", "tablas": []}
    try:
        resp = requests.get(dl_url, headers={"User-Agent": "Mozilla/5.0"},
                            timeout=30, allow_redirects=True)
        resp.raise_for_status()
        texto_paginas, tablas = [], []
        with pdfplumber.open(BytesIO(resp.content)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    texto_paginas.append(t)
                for tbl in page.extract_tables():
                    if tbl:
                        tablas.append(tbl)
        return {"texto": "\n".join(texto_paginas), "tablas": tablas}
    except Exception:
        return {"texto": "", "tablas": []}


# ============================================================
# MOTOR DE BÚSQUEDA LOCAL EN PDF
# ============================================================
STOPWORDS = {
    "de","del","la","el","los","las","un","una","unos","unas","que","en","a",
    "y","o","es","se","no","con","por","para","al","lo","le","su","sus","me",
    "mi","te","si","más","pero","ya","hay","como","este","esta","ese","esa",
    "fue","son","ser","cual","cuál","qué","cuanto","dime","dame",
    "informacion","información","sobre","acerca","tanque","equipo","bomba",
    "filtro","compresor","intercambiador","puedes","decirme","indicarme",
    "necesito","quiero","busco","muestra","muestrame","cuales",
    "valor","valores","dato","datos","tiene","tienes",
}
RUIDO_PDF = {
    "preparado por","revisado por","aprobado por","aceptado por",
    "fecha:","página ","www.","pág.","revist","all rights",
}

def extraer_keywords(pregunta: str) -> list[str]:
    tokens = re.findall(r"[a-záéíóúüñA-ZÁÉÍÓÚÜÑ0-9\-]+", pregunta.lower())
    return [t for t in tokens if len(t) > 2 and t not in STOPWORDS]

def _es_ruido(linea: str) -> bool:
    l = linea.lower()
    return any(r in l for r in RUIDO_PDF)

def _es_encabezado(texto: str) -> bool:
    t = texto.strip()
    if len(t) < 20:
        return True
    if re.match(r"^\d+[\.\d]*\s+[A-ZÁÉÍÓÚÜÑ\s]+$", t):
        return True
    if t == t.upper() and len(t) < 60:
        return True
    return False

def _limpiar(txt: str) -> str:
    return " ".join(txt.split())

def buscar_en_tablas(tablas: list, keywords: list[str], max_filas: int = 10) -> str | None:
    mejores = []
    for tabla in tablas:
        for fila in tabla:
            celdas = [str(c or "").strip() for c in fila]
            fila_texto = " ".join(celdas).lower()
            score = sum(1 for k in keywords if k in fila_texto)
            if score > 0 and len(celdas) >= 2:
                pares = [c for c in celdas if c and c != "nan"]
                if len(pares) >= 2:
                    mejores.append((score, " | ".join(pares)))
    if not mejores:
        return None
    mejores.sort(reverse=True)
    vistos, resultado = set(), []
    for _, texto in mejores:
        clave = texto[:40].lower()
        if clave not in vistos and not _es_ruido(texto):
            vistos.add(clave)
            resultado.append(texto)
        if len(resultado) >= max_filas:
            break
    return "\n".join(resultado) if resultado else None

def buscar_en_texto(texto: str, keywords: list[str],
                    incluir_encabezados: bool = False,
                    max_oraciones: int = 6) -> str | None:
    """
    Busca oraciones relevantes. Prioriza oraciones descriptivas (largas).
    Si incluir_encabezados=True, acepta también títulos de sección como contexto.
    """
    if not texto or not keywords:
        return None

    oraciones_raw = re.split(r"(?<=[.;])\s+|\n", texto)
    oraciones = [o.strip() for o in oraciones_raw if o.strip() and len(o.strip()) > 8]

    def puntuar(lista):
        scored = []
        for oracion in lista:
            if _es_ruido(oracion):
                continue
            o_lower = oracion.lower()
            score = sum(1 for k in keywords if k in o_lower)
            if score > 0:
                bonus = min(score, 3) * 0.1
                scored.append((score + bonus, oracion))
        return scored

    # Primero buscar en oraciones descriptivas (sin encabezados)
    scored = puntuar([o for o in oraciones if not _es_encabezado(o)])

    # Fallback: si no hay resultados descriptivos, aceptar encabezados también
    if not scored and incluir_encabezados:
        scored = puntuar(oraciones)

    if not scored:
        return None

    # Preferir oraciones más largas (más descriptivas) cuando el score es igual
    scored.sort(key=lambda x: (-x[0], -len(x[1])))

    vistos, resultado = set(), []
    for _, oracion in scored:
        clave = oracion[:40].lower()
        if clave in vistos:
            continue
        vistos.add(clave)
        resultado.append(_limpiar(oracion))
        if len(resultado) >= max_oraciones:
            break
    return "\n".join(resultado) if resultado else None

def _resumen_completo_pdf(contenido: dict) -> str:
    """Extrae TODA la información estructurada del PDF para mostrar ficha completa."""
    partes = []

    # Tablas: todas las filas con par clave-valor
    if contenido["tablas"]:
        filas_tabla = []
        vistos = set()
        for tabla in contenido["tablas"]:
            for fila in tabla:
                celdas = [str(c or "").strip() for c in fila if str(c or "").strip() and str(c or "").strip() != "nan"]
                if len(celdas) >= 2:
                    texto = " | ".join(celdas)
                    clave = texto[:50].lower()
                    if clave not in vistos and not _es_ruido(texto):
                        vistos.add(clave)
                        filas_tabla.append(texto)
        if filas_tabla:
            partes.append("**📊 Datos del plano técnico:**\n" + "\n".join(filas_tabla[:30]))

    # Texto: párrafos descriptivos no-ruido
    if contenido["texto"]:
        oraciones = [o.strip() for o in re.split(r"(?<=[.;])\s+|\n", contenido["texto"])
                     if o.strip() and len(o.strip()) > 20
                     and not _es_ruido(o) and not _es_encabezado(o)]
        if oraciones:
            partes.append("**📝 Descripción técnica:**\n" + "\n".join(oraciones[:10]))

    return "\n\n".join(partes) if partes else ""

def consultar_pdf(plano_url: str, pregunta: str, tag: str,
                  resumen_completo: bool = False) -> str | None:
    if not plano_url or plano_url == "nan":
        return None
    with st.spinner(f"🔍 Consultando plano técnico de {tag}..."):
        contenido = extraer_contenido_pdf(plano_url)

    # Modo resumen completo (sin keywords)
    if resumen_completo:
        return _resumen_completo_pdf(contenido) or None

    keywords = extraer_keywords(pregunta)
    if not keywords:
        return _resumen_completo_pdf(contenido) or None

    # 1. Tablas (más estructuradas y precisas) — hasta 10 filas relevantes
    resp_tablas = buscar_en_tablas(contenido["tablas"], keywords, max_filas=10)

    # 2. Texto — oraciones descriptivas; si no hay, acepta encabezados como guía
    resp_texto = buscar_en_texto(contenido["texto"], keywords,
                                 incluir_encabezados=True, max_oraciones=6)

    partes = []
    if resp_tablas:
        partes.append("**📊 Datos técnicos encontrados:**\n" + resp_tablas)
    if resp_texto:
        partes.append("**📝 Texto relevante:**\n" + resp_texto)

    return "\n\n".join(partes) if partes else None


# ============================================================
# UTILIDADES DE PRESENTACIÓN
# ============================================================
def gdrive_preview(url: str) -> str:
    if not url:
        return ""
    url = url.strip()
    if "/view" in url:
        url = url.split("/view")[0] + "/preview"
    return url

def estado_icon(estado: str) -> str:
    s = estado.lower()
    return "🟢" if ("servicio" in s or "operativ" in s) else "🔴"

def card_equipo(tag: str, info: dict) -> str:
    icon      = estado_icon(info.get("Estado", ""))
    plano_url = info.get("Plano", "").strip()
    prev      = gdrive_preview(plano_url)
    plano_txt = f"\n\n📄 [Ver plano completo]({prev})" if prev else ""
    excluir   = {"Tag", "Plano", "X_Plano", "Y_Plano"}
    filas     = ""
    for col, val in info.items():
        if col in excluir or not val or val == "nan":
            continue
        filas += f"| **{col}** | {val} |\n"
    return (
        f"### {tag} {icon}\n\n"
        f"| Parámetro | Detalle |\n|---|---|\n"
        f"{filas}"
        f"{plano_txt}\n\n"
        f"¿Necesitas más información?"
    )


# ============================================================
# CONFIGURACIÓN DE PÁGINA
# ============================================================
st.set_page_config(
    page_title="Asistente Técnico de Planta",
    page_icon="🤖",
    layout="centered"
)

logo_path   = os.path.join(ASSETS_DIR, "logo.png")
planta_path = os.path.join(ASSETS_DIR, "planta.jpg")

col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    if os.path.exists(logo_path):
        st.image(logo_path, width=80)
with col_titulo:
    st.title("Asistente Técnico de Planta")

datos, estado_carga = cargar_datos()

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.header("📋 Equipos Registrados")
    st.caption(estado_carga)
    st.markdown("---")
    categoria_label = st.selectbox(
        "Selecciona una categoría:",
        options=list(ICONOS_CATEGORIA.values())
    )
    categoria_key = [k for k, v in ICONOS_CATEGORIA.items() if v == categoria_label][0]
    st.markdown("---")
    df_cat = datos.get(categoria_key, pd.DataFrame())
    if not df_cat.empty:
        for idx, row in df_cat.iterrows():
            tag    = str(row.get("Tag", "")).strip()
            estado = str(row.get("Estado", ""))
            icono  = estado_icon(estado)
            if tag and tag != "nan":
                if st.button(f"{icono} {tag}", key=f"btn_{categoria_key}_{tag}_{idx}",
                             use_container_width=True):
                    st.session_state.pending_query = (categoria_key, tag)
    else:
        st.info("Sin equipos en esta categoría.")

# ============================================================
# ESTADO DE SESIÓN
# ============================================================
for key, default in [("messages", []), ("equipo_actual", None), ("pending_query", None)]:
    if key not in st.session_state:
        st.session_state[key] = default

# ============================================================
# TABS PRINCIPALES
# ============================================================
tab_chat, tab_mapa = st.tabs(["💬 Asistente", "🗺️ Mapa de Planta"])

# ============================================================
# TAB 1 — CHAT
# ============================================================
with tab_chat:
    st.write("Pregúntame sobre cualquier equipo de planta.")
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # ── Búsqueda de equipo ─────────────────────────────────────
    def buscar_equipo(query: str):
        q = query.lower().strip()
        for cat in TODAS_CATEGORIAS:
            df = datos.get(cat, pd.DataFrame())
            for _, row in df.iterrows():
                tag = str(row.get("Tag", "")).strip()
                if not tag or tag == "nan":
                    continue
                if tag.lower() == q or tag.lower() in q or q in tag.lower():
                    return cat, tag, row.to_dict()
        return None

    def respuesta_sidebar(categoria: str, tag: str) -> str:
        df  = datos.get(categoria, pd.DataFrame())
        row = df[df["Tag"] == tag]
        if row.empty:
            return f"No encontré datos para **{tag}**."
        st.session_state.equipo_actual = (categoria, tag)
        return card_equipo(tag, row.iloc[0].to_dict())

    # Columnas de la hoja con palabras clave de detección
    # NOTA: "fabricación"/"fabricacion" se quitaron del trigger de "Año de construcción"
    # porque son demasiado genéricas y colisionan con preguntas de proceso.
    KEYWORD_A_COLUMNA: list[tuple[list[str], str]] = [
        (["altura", "alto", "mide", "dimension", "dimensión"], "Altura"),
        (["diámetro", "diametro", "ancho", "diám"], "Diámetro"),
        (["volumen", "capacidad", "litros", "m3"], "Volumen"),
        (["fabricante", "proveedor", "quien hizo", "quien fabrico"], "Fabricante"),
        (["año", "cuando fue", "fecha construccion", "fecha de construccion"], "Año de construcción"),
        (["estado", "condición", "condicion", "operativo"], "Estado"),
        (["servicio", "almacena", "contiene", "químico", "quimico", "producto"], "Servicio"),
        (["zona", "ubica", "sector", "área", "area"], "Zona"),
    ]

    def detectar_columna_hoja(q: str) -> str | None:
        for palabras, columna in KEYWORD_A_COLUMNA:
            if any(p in q for p in palabras):
                return columna
        return None

    PALABRAS_DETALLE_PDF = [
        "detalle", "detalles", "todo", "completo", "completa", "tecnico",
        "técnico", "especificacion", "especificaciones", "ficha", "hoja",
        "caracteristica", "características", "mas informacion", "más información",
        "dame todo", "dime todo", "que dice", "qué dice", "lee el", "leer",
    ]

    def respuesta_chat(pregunta: str) -> str:
        q = pregunta.lower()

        # ── Enlace/plano ───────────────────────────────────────
        pide_enlace = any(p in q for p in ["link", "enlace", "abrir", "ver plano", "abrir plano"])
        pide_pdf    = any(p in q for p in ["pdf", "plano", "documento"])
        pide_detalle_pdf = any(p in q for p in PALABRAS_DETALLE_PDF)

        if pide_enlace and not pide_detalle_pdf:
            ctx = st.session_state.equipo_actual
            if ctx:
                cat, tag = ctx
                df  = datos.get(cat, pd.DataFrame())
                row = df[df["Tag"] == tag]
                if not row.empty:
                    plano = str(row.iloc[0].get("Plano", "")).strip()
                    prev  = gdrive_preview(plano)
                    if prev:
                        return f"📄 Plano de **{tag}**:\n\n[👉 Abrir en Google Drive]({prev})"
                    return f"No hay plano registrado para **{tag}**."
            return "Primero indícame de qué equipo necesitas el plano."

        # ── Identificar equipo ─────────────────────────────────
        resultado = buscar_equipo(pregunta)
        if resultado is None and st.session_state.equipo_actual:
            cat, tag = st.session_state.equipo_actual
            df  = datos.get(cat, pd.DataFrame())
            row = df[df["Tag"] == tag]
            if not row.empty:
                resultado = (cat, tag, row.iloc[0].to_dict())

        if resultado is None:
            return (
                "No encontré ese equipo. Prueba con el tag exacto "
                "(ej. **T-131**, **T-706A**) o selecciónalo desde el menú lateral."
            )

        cat, tag, info = resultado
        st.session_state.equipo_actual = (cat, tag)
        plano_url = str(info.get("Plano", "")).strip()
        prev_link  = gdrive_preview(plano_url)

        # ── Detalle completo del PDF (ficha técnica completa) ──
        if (pide_detalle_pdf or pide_pdf) and plano_url and plano_url != "nan":
            ficha_hoja = card_equipo(tag, info)
            resp_pdf   = consultar_pdf(plano_url, pregunta, tag, resumen_completo=pide_detalle_pdf)
            if resp_pdf:
                partes = [ficha_hoja, f"\n---\n### 📋 Plano Técnico de {tag}\n\n{resp_pdf}"]
                if prev_link:
                    partes.append(f"\n📄 [Ver plano completo en Drive]({prev_link})")
                return "\n".join(partes)
            return ficha_hoja

        # ── Columna conocida de la hoja ────────────────────────
        columna = detectar_columna_hoja(q)
        if columna:
            valor = str(info.get(columna, "")).strip()
            if valor and valor != "nan":
                # También buscar en PDF para complementar
                resp_pdf = consultar_pdf(plano_url, pregunta, tag) if plano_url and plano_url != "nan" else None
                respuesta = f"**{columna}** de **{tag}**: {valor}"
                if resp_pdf:
                    respuesta += f"\n\n📋 **Información adicional del plano:**\n{resp_pdf}"
                if prev_link:
                    respuesta += f"\n\n📄 [Ver plano completo]({prev_link})"
                return respuesta
            # No está en la hoja — buscar en PDF
            resp_pdf = consultar_pdf(plano_url, pregunta, tag) if plano_url and plano_url != "nan" else None
            if resp_pdf:
                r = f"📋 **{tag}** — encontrado en el plano técnico:\n\n{resp_pdf}"
                if prev_link:
                    r += f"\n\n📄 [Ver plano completo]({prev_link})"
                return r
            return f"No encontré **{columna}** para **{tag}** ni en la ficha ni en el plano."

        # ── ¿Pregunta genérica o tema específico? ──────────────
        keywords   = extraer_keywords(pregunta)
        tag_tokens = set(re.findall(r"[a-z0-9]+", tag.lower()))
        kw_tema    = [k for k in keywords if k not in tag_tokens]

        if not kw_tema:
            # Ficha básica + resumen de PDF si lo hay
            ficha = card_equipo(tag, info)
            if plano_url and plano_url != "nan":
                resp_pdf = consultar_pdf(plano_url, pregunta, tag, resumen_completo=True)
                if resp_pdf:
                    ficha += f"\n\n---\n### 📋 Plano Técnico\n{resp_pdf}"
                    if prev_link:
                        ficha += f"\n\n📄 [Ver plano completo]({prev_link})"
            return ficha

        # Buscar tema en valores de la hoja
        coincidencias_hoja = []
        for col, val in info.items():
            if col in {"Tag", "Plano", "Equipo", "X_Plano", "Y_Plano"}:
                continue
            val_str = str(val).lower()
            if val_str and val_str != "nan" and any(k in val_str for k in kw_tema):
                coincidencias_hoja.append(f"**{col}**: {val}")

        # Buscar en PDF
        resp_pdf = None
        if plano_url and plano_url != "nan":
            resp_pdf = consultar_pdf(plano_url, pregunta, tag)

        if coincidencias_hoja or resp_pdf:
            partes = []
            if coincidencias_hoja:
                partes.append("**📊 Datos de la ficha:**\n" + "\n".join(coincidencias_hoja))
            if resp_pdf:
                partes.append(f"**📋 Del plano técnico de {tag}:**\n{resp_pdf}")
            if prev_link:
                partes.append(f"📄 [Ver plano completo]({prev_link})")
            return "\n\n".join(partes)

        return card_equipo(tag, info)

    def agregar_mensaje(role: str, content: str):
        st.session_state.messages.append({"role": role, "content": content})
        with st.chat_message(role):
            st.markdown(content)

    # Procesar clic del sidebar
    if st.session_state.pending_query:
        categoria, tag = st.session_state.pending_query
        st.session_state.pending_query = None
        agregar_mensaje("user",      f"Información de {tag}")
        agregar_mensaje("assistant", respuesta_sidebar(categoria, tag))

    if pregunta := st.chat_input("¿En qué te puedo ayudar hoy?"):
        agregar_mensaje("user",      pregunta)
        agregar_mensaje("assistant", respuesta_chat(pregunta))


# ============================================================
# TAB 2 — MAPA DE PLANTA
# ============================================================
with tab_mapa:
    st.subheader("🗺️ Ubicación de Equipos en Planta")

    if not os.path.exists(planta_path):
        st.warning("No se encontró la imagen del plano de planta (assets/planta.jpg).")
    else:
        img_planta = Image.open(planta_path)
        img_w, img_h = img_planta.size

        # Recopilar todos los equipos con coordenadas X_Plano / Y_Plano
        equipos_con_coords = []
        for cat in TODAS_CATEGORIAS:
            df = datos.get(cat, pd.DataFrame())
            if df.empty:
                continue
            for _, row in df.iterrows():
                tag = str(row.get("Tag", "")).strip()
                if not tag or tag == "nan":
                    continue
                x_raw = str(row.get("X_Plano", "")).strip()
                y_raw = str(row.get("Y_Plano", "")).strip()
                try:
                    x = float(x_raw)
                    y = float(y_raw)
                    estado = str(row.get("Estado", ""))
                    servicio = str(row.get("Servicio", ""))
                    zona = str(row.get("Zona", ""))
                    equipos_con_coords.append({
                        "tag": tag, "cat": cat,
                        "x": x, "y": y,
                        "estado": estado,
                        "servicio": servicio,
                        "zona": zona,
                        "color": "#22c55e" if ("servicio" in estado.lower() or
                                               "operativ" in estado.lower()) else "#ef4444"
                    })
                except (ValueError, TypeError):
                    continue  # Equipo sin coordenadas aún

        # Construir figura Plotly
        fig = go.Figure()

        # Imagen de fondo
        fig.add_layout_image(
            dict(
                source=img_planta,
                x=0, y=0,
                xref="x", yref="y",
                sizex=img_w, sizey=img_h,
                xanchor="left", yanchor="bottom",
                sizing="stretch",
                layer="below"
            )
        )

        if equipos_con_coords:
            # Separar por estado para la leyenda
            for estado_label, filtro, color_marker in [
                ("En Servicio 🟢", lambda e: "servicio" in e["estado"].lower() or
                                              "operativ" in e["estado"].lower(), "#22c55e"),
                ("Fuera de Servicio 🔴", lambda e: not ("servicio" in e["estado"].lower() or
                                                         "operativ" in e["estado"].lower()), "#ef4444"),
            ]:
                subset = [e for e in equipos_con_coords if filtro(e)]
                if not subset:
                    continue
                fig.add_trace(go.Scatter(
                    x=[e["x"] for e in subset],
                    y=[e["y"] for e in subset],
                    mode="markers+text",
                    marker=dict(
                        size=18,
                        color=color_marker,
                        line=dict(color="white", width=2),
                        symbol="circle",
                    ),
                    text=[e["tag"] for e in subset],
                    textposition="top center",
                    textfont=dict(size=11, color="white",
                                  family="Arial Black"),
                    name=estado_label,
                    customdata=[
                        [e["cat"], e["estado"], e["servicio"], e["zona"]]
                        for e in subset
                    ],
                    hovertemplate=(
                        "<b>%{text}</b><br>"
                        "Categoría: %{customdata[0]}<br>"
                        "Estado: %{customdata[1]}<br>"
                        "Servicio: %{customdata[2]}<br>"
                        "Zona: %{customdata[3]}"
                        "<extra></extra>"
                    ),
                ))
        else:
            # Sin coordenadas: mostrar la imagen igual con aviso
            fig.add_annotation(
                x=img_w / 2, y=img_h / 2,
                text="Agrega columnas <b>X_Plano</b> e <b>Y_Plano</b> al Google Sheet<br>"
                     "con las coordenadas en píxeles de cada equipo.",
                showarrow=False,
                font=dict(size=16, color="white"),
                bgcolor="rgba(0,0,0,0.6)",
                bordercolor="white",
                borderwidth=1,
                borderpad=10,
                align="center",
            )

        fig.update_layout(
            xaxis=dict(range=[0, img_w], showgrid=False,
                       zeroline=False, showticklabels=False),
            yaxis=dict(range=[0, img_h], showgrid=False,
                       zeroline=False, showticklabels=False,
                       scaleanchor="x"),
            margin=dict(l=0, r=0, t=0, b=0),
            height=600,
            paper_bgcolor="black",
            plot_bgcolor="black",
            legend=dict(
                orientation="h",
                yanchor="bottom", y=1.01,
                xanchor="left", x=0,
                bgcolor="rgba(0,0,0,0.5)",
                font=dict(color="white"),
            ),
            dragmode="pan",
        )
        fig.update_layout(
            modebar_add=["zoom", "pan", "resetScale2d"],
        )

        st.plotly_chart(fig, use_container_width=True, config={
            "scrollZoom": True,
            "displaylogo": False,
            "modeBarButtonsToRemove": ["select2d", "lasso2d"],
        })

        # Instrucciones para añadir coordenadas
        with st.expander("ℹ️ ¿Cómo ubicar equipos en el mapa?"):
            st.markdown("""
**Para mostrar un equipo en el mapa**, agrega dos columnas a tu Google Sheet:

| Columna | Descripción |
|---|---|
| `X_Plano` | Posición horizontal en píxeles desde el borde izquierdo de la imagen |
| `Y_Plano` | Posición vertical en píxeles desde el borde **inferior** de la imagen |

El tamaño de la imagen del plano es **{w} × {h} píxeles**.

Puedes abrir la imagen en cualquier editor de imágenes (Paint, Preview, etc.)
y pasar el cursor sobre la ubicación del equipo para leer las coordenadas X e Y.

> **Nota:** El eje Y de la imagen va de 0 (abajo) a {h} (arriba).
> Si tu editor muestra Y desde arriba, calcula: `Y_Plano = {h} − Y_editor`
""".format(w=img_w, h=img_h))

        # Tabla resumen de equipos sin coordenadas
        sin_coords = []
        for cat in TODAS_CATEGORIAS:
            df = datos.get(cat, pd.DataFrame())
            if df.empty:
                continue
            for _, row in df.iterrows():
                tag = str(row.get("Tag", "")).strip()
                if not tag or tag == "nan":
                    continue
                x_raw = str(row.get("X_Plano", "")).strip()
                y_raw = str(row.get("Y_Plano", "")).strip()
                tiene_coords = False
                try:
                    float(x_raw); float(y_raw)
                    tiene_coords = True
                except (ValueError, TypeError):
                    pass
                if not tiene_coords:
                    sin_coords.append({"Tag": tag, "Categoría": cat})

        if sin_coords:
            st.caption(f"⚠️ {len(sin_coords)} equipo(s) sin coordenadas registradas:")
            st.dataframe(pd.DataFrame(sin_coords), use_container_width=True, hide_index=True)
