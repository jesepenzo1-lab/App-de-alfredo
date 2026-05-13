import streamlit as st
import google.generativeai as genai
import io
import time
from docx import Document

# ----------------- CONFIGURACIÓN -----------------
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=GEMINI_API_KEY)

MODEL_NAME = "gemini-2.5-flash"

# ----------------- PROMPT LEGAL PREDEFINIDO -----------------
PROMPT_PLANTILLA = """
Rol del Asistente: Actúa como un asistente legal experto en análisis procesal, redacción de sentencias y corrección lingüística especializada en español jurídico
... (completo igual) ...
"""

# ----------------- FUNCIONES DE EXTRACCIÓN -----------------

def extraer_texto_docx(file_bytes):
    try:
        doc = Document(io.BytesIO(file_bytes))
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        raise RuntimeError(f"Error al procesar Word: {e}")

def extraer_texto_varios_archivos(archivos_subidos, modelo):
    partes_solicitud = []
    lista_docx = []
    mapeo_indices = []

    for archivo in archivos_subidos:
        nombre = archivo.name
        mime = archivo.type
        bytes_data = archivo.read()

        if mime == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            lista_docx.append((nombre, bytes_data))
            continue

        if mime not in ["application/pdf", "image/jpeg", "image/png", "image/jpg"]:
            st.warning(f"Tipo de archivo no soportado: {nombre} ({mime})")
            continue

        partes_solicitud.append({
            "mime_type": mime,
            "data": bytes_data
        })
        mapeo_indices.append(nombre)

    if not partes_solicitud:
        textos = []
        for nombre, data in lista_docx:
            try:
                txt = extraer_texto_docx(data)
                textos.append((nombre, txt))
            except Exception as e:
                st.warning(f"Error con {nombre}: {e}")
                textos.append((nombre, "[Error al extraer]"))
        return textos

    prompt_instrucciones = (
        "Extrae TODO el texto de cada uno de los siguientes documentos, conservando fechas, nombres y números tal cual aparecen. "
        "Devuelve el resultado en este formato EXACTO:\n\n"
        "--- DOCUMENTO: nombre_real_del_archivo ---\n"
        "[texto extraído]\n\n"
        "Repite esa estructura para cada documento, en el mismo orden en que se proporcionan."
    )

    contenido = [prompt_instrucciones] + partes_solicitud

    for intento in range(2):
        try:
            respuesta = modelo.generate_content(contenido)
            texto_completo = respuesta.text
            break
        except Exception as e:
            if "429" in str(e) and intento == 0:
                st.warning("Límite de cuota alcanzado. Esperando 60 segundos...")
                time.sleep(60)
            else:
                st.error(f"Error al extraer texto con Gemini: {e}")
                textos = []
                for nombre, data in lista_docx:
                    try:
                        txt = extraer_texto_docx(data)
                        textos.append((nombre, txt))
                    except:
                        textos.append((nombre, "[Error]"))
                return textos

    # Parseo robusto
    textos_resultantes = []
    bloques = texto_completo.split("--- DOCUMENTO: ")
    if not bloques:
        st.error("Gemini no devolvió el formato esperado. Revisa la respuesta manualmente.")
        return []

    # Quitar bloque intro vacío
    if bloques[0].strip() == "":
        bloques = bloques[1:]

    for bloque in bloques:
        if "---" in bloque:
            partes = bloque.split("---", 1)
            nombre = partes[0].strip()
            texto = partes[1].strip() if len(partes) > 1 else ""
            textos_resultantes.append((nombre, texto))
        else:
            textos_resultantes.append(("Desconocido", bloque.strip()))

    for nombre, data in lista_docx:
        try:
            txt = extraer_texto_docx(data)
            textos_resultantes.append((nombre, txt))
        except Exception as e:
            st.warning(f"Error con Word {nombre}: {e}")
            textos_resultantes.append((nombre, "[Error al extraer]"))

    return textos_resultantes

def generar_texto_legal(fuentes, texto_base, modelo):
    prompt = PROMPT_PLANTILLA.format(fuentes=fuentes, texto_base=texto_base)
    for intento in range(2):
        try:
            respuesta = modelo.generate_content(prompt)
            return respuesta.text
        except Exception as e:
            if "429" in str(e) and intento == 0:
                st.warning("Límite alcanzado. Reintentando en 60 segundos...")
                time.sleep(60)
            else:
                st.error(f"Error al generar texto legal: {e}")
                return None

# ----------------- INTERFAZ STREAMLIT -----------------
st.set_page_config(page_title="Asistente Legal Procesal (Cloud)", layout="wide")
st.title("⚖️ Asistente Legal Procesal (Nube)")
st.markdown("Extrae texto de documentos (PDF, imágenes, Word) usando Gemini y genera un texto legal corregido según tus instrucciones.")

col1, col2 = st.columns(2)
with col1:
    st.header("1. Documentos fuente (verdad procesal)")
    archivos_fuente = st.file_uploader(
        "Sube PDFs, imágenes (JPG, PNG) o Word (DOCX)",
        type=["pdf", "jpg", "jpeg", "png", "docx"],
        accept_multiple_files=True,
        key="fuentes"
    )
    st.header("2. Texto base (plantilla)")
    texto_base = st.text_area("Pega aquí el texto base o plantilla:", height=300, key="texto_base")

with col2:
    st.header("Resultado")
    if st.button("Procesar", type="primary", disabled=not (archivos_fuente and texto_base.strip())):
        with st.spinner("Extrayendo texto de todos los documentos con Gemini (una sola llamada)..."):
            model = genai.GenerativeModel(MODEL_NAME)
            lista_textos = extraer_texto_varios_archivos(archivos_fuente, model)

            fuentes_completo = ""
            for nombre, txt in lista_textos:
                fuentes_completo += f"--- DOCUMENTO: {nombre} ---\n{txt}\n\n"

        with st.spinner("Generando texto legal..."):
            resultado = generar_texto_legal(fuentes_completo, texto_base, model)

        if resultado:
            st.success("Procesamiento completado.")
            if "Validación con Fuentes:" in resultado:
                partes = resultado.split("Validación con Fuentes:", 1)
                texto_generado = partes[0].replace("Texto Generado:", "").strip()
                validacion = partes[1].strip()
            else:
                texto_generado = resultado
                validacion = "No se encontró la sección de validación."
            st.subheader("Texto Generado")
            st.markdown(texto_generado)
            st.subheader("Validación con Fuentes")
            st.markdown(validacion)
