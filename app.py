import streamlit as st
import google.generativeai as genai
import io
import os
from docx import Document

# ----------------- CONFIGURACIÓN -----------------
# La API Key de Gemini se toma de los secretos de Streamlit Cloud
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=GEMINI_API_KEY)

# Modelo a utilizar (flash es rápido y económico)
MODEL_NAME = "gemini-2.5-flash"

# ----------------- PROMPT LEGAL PREDEFINIDO -----------------
PROMPT_PLANTILLA = """
Rol del Asistente: Actúa como un asistente legal experto en análisis procesal, redacción de sentencias y corrección lingüística especializada en español jurídico
. Instrucciones de Procesamiento: Debes procesar el [TEXTO BASE] que te proporcionaré al final de este mensaje, comparándolo con los documentos del expediente suministrado
. Sigue estrictamente estos tres pasos en el orden indicado
:
PASO 1: Aplicación de Instrucciones "Texto Base" (Sustitución y Reglas Procesales)
 Procesamiento interno: Sustitución de datos: Identifica fechas, números de acto, nombres, plazos o números de solicitud en el [TEXTO BASE] y sustitúyelos por la información real y verídica del expediente
. Prioriza siempre la etapa procesal actual y el documento más reciente
. Si el documento más reciente carece de un comprobante de recepción físico visible, utiliza la fecha de suscripción del documento
. Si el usuario te proporciona un número de solicitud o dato específico en su instrucción, asúmelo como correcto e intégralo; de lo contrario, indica [Dato no visible]
. NUNCA utilices números de solicitud, recibos o fechas de etapas procesales ya precluidas o pasadas (ej. solicitudes de años anteriores) para rellenar los vacíos de una actuación reciente
. Las conclusiones deben ser transcripciones fieles
. Fechas y Números: Utiliza el formato día/mes/año sin ceros iniciales (ejemplo: 1/5/2025)
. Sustituye "No.", "número" o "#" por la abreviatura "núm."
. Mayúsculas y Tribunales: Ninguna palabra puede tener más de una letra escrita en mayúscula, a excepción de siglas como RD$, USD, RNC o S.A.
. No abrevies SCJ, escribe siempre Suprema Corte de Justicia
. Los nombres de los tribunales se escriben con iniciales mayúsculas (ejemplo: Cuarta Sala de la Cámara Civil y Comercial del Juzgado de Primera Instancia del Distrito Nacional)
. Ortografía Estructural: La primera letra de un párrafo se escribe en mayúscula y todos los párrafos concluyen con punto (.)
. Asegura que no haya más de un espacio entre palabras
. Los ordinales numéricos se escriben: Primero, Segundo, Tercero
. Después de un ordinal o de dos puntos (:), la primera palabra inicia con letra mayúscula
.
PASO 2: Aplicación del "Corrector ABC" (Ajustes Lingüísticos y Presentación)
 Procesamiento interno antes de mostrar el resultado: Fidelidad al texto y Verdad Material: Corrige errores evidentes de forma y aplica las normas de la RAE (ortografía, tildes, puntuación y concordancia) sin sustituir términos jurídicos por sinónimos
. Como regla general, mantén la misma estructura del [TEXTO BASE], SIN EMBARGO, tienes autorización para flexibilizar esta regla, reformular o agregar literales nuevos (ej. un literal e) si detectas que el [TEXTO BASE] ha omitido un argumento de defensa vital o una pretensión fundamental (ej. alegatos de fuerza mayor) que sí constan claramente en los documentos aportados por las partes. Tu prioridad absoluta es ajustar la respuesta atendiendo a la realidad procesal solicitada por las partes en las fuentes. Excepción Específica (Minúsculas Obligatorias): Escribe SIEMPRE en minúscula las siguientes palabras, incluso al inicio de una frase (salvo que formen parte de un nombre propio oficial o vayan justo después de dos puntos): ciudad, provincia, municipio, calle, avenida, condominio, plaza, abogado, abogada, doctor, doctora, licenciada, licenciado, notario, acto, ministerial, alguacil, contrato, alquiler, local, comercial, artículo, entidad, certificado, título, matrícula, inmueble, sociedad, domicilio, demanda, demandada, demandante, parte, convenciones, ley, rescindir, resciliación, dispositivo, provisional, señora, señor, señores, número, interpuesta, propietaria, propietario, fiador, fiadora, solidario, solidaria, kilómetro, esquina, desalojo, desaloja
. (Aclaración: El nombre propio que las acompaña sí conserva su mayúscula)
. Expresiones fijas: Escribe siempre la expresión "vía el Centro de Servicio Presencial" exactamente así
. Acción de Salida: Una vez aplicados estos filtros silenciosamente, devuelve el texto corregido bajo el encabezado "Texto Generado:", manteniendo la misma distribución de párrafos y líneas que el texto original, sin explicaciones adicionales en esta parte
.
PASO 3: Validación con Fuentes (Veredicto de Veracidad)
 Inmediatamente después de presentar el párrafo, incluye un encabezado llamado "Validación con Fuentes:"
. Evalúa el texto que acabas de generar en relación con los datos de las fuentes suministradas
. Escribe la palabra VERDADERO (si todos los nombres, números, fechas y partes procesales corresponden exactamente a las fuentes) o FALSO (si el texto contiene alguna discrepancia o alucinación)
. Acompaña este resultado con una justificación puntual indicando de cuáles documentos extrajiste la verdad procesal
. Si utilizaste un número de solicitud u otro dato aportado por el usuario que no consta físicamente en las imágenes del expediente, aclara en tu justificación que este dato fue asumido por instrucción directa del usuario para la etapa procesal actual
.

FUENTES (VERDAD PROCESAL):
{fuentes}

[TEXTO BASE]:
{texto_base}
"""

# ----------------- FUNCIÓN DE EXTRACCIÓN CON GEMINI -----------------

def extraer_texto_con_gemini(file_bytes, mime_type):
    """Envía el archivo directamente a Gemini para que extraiga el texto."""
    model = genai.GenerativeModel(MODEL_NAME)
    try:
        response = model.generate_content([
            {"mime_type": mime_type, "data": file_bytes},
            "Extrae todo el texto de este documento, conservando las fechas, nombres y números tal cual aparecen."
        ])
        return response.text
    except Exception as e:
        st.error(f"Error al leer el archivo con Gemini: {e}")
        return "[Error al extraer texto]"

def extraer_texto_docx(file_bytes):
    """Extrae texto de un Word utilizando python-docx (en el servidor)."""
    try:
        doc = Document(io.BytesIO(file_bytes))
        texto = "\n".join([para.text for para in doc.paragraphs])
        return texto
    except Exception as e:
        st.error(f"Error al procesar Word: {e}")
        return "[Error al extraer texto]"

# ----------------- LLAMADA PRINCIPAL A GEMINI -----------------

def generar_texto_legal(fuentes, texto_base):
    prompt = PROMPT_PLANTILLA.format(fuentes=fuentes, texto_base=texto_base)
    model = genai.GenerativeModel(MODEL_NAME)
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        st.error(f"Error al generar el texto legal: {e}")
        return None

# ----------------- INTERFAZ STREAMLIT -----------------

st.set_page_config(page_title="Asistente Legal Procesal (Cloud)", layout="wide")
st.title("⚖️ Asistente Legal Procesal (Nube)")
st.markdown("Sustituye datos en plantillas legales según la verdad procesal extraída directamente por Gemini. **Sin instalaciones locales.**")

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
        # Extraer texto de todos los archivos
        with st.spinner("Extrayendo texto de los documentos con Gemini..."):
            textos_fuentes = []
            for archivo in archivos_fuente:
                bytes_data = archivo.read()
                mime = archivo.type
                if mime == "application/pdf":
                    txt = extraer_texto_con_gemini(bytes_data, "application/pdf")
                elif mime in ["image/jpeg", "image/png", "image/jpg"]:
                    txt = extraer_texto_con_gemini(bytes_data, mime)
                elif mime == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                    txt = extraer_texto_docx(bytes_data)  # Word se procesa localmente (python-docx en servidor)
                else:
                    st.warning(f"Tipo no soportado: {mime}")
                    continue
                textos_fuentes.append(f"--- DOCUMENTO: {archivo.name} ---\n{txt}")
            fuentes_completo = "\n\n".join(textos_fuentes)

        with st.spinner("Generando texto legal (esto puede tomar unos segundos)..."):
            resultado = generar_texto_legal(fuentes_completo, texto_base)

        if resultado:
            st.success("Procesamiento completado.")
            # Separar secciones esperadas
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
