import streamlit as st
import google.generativeai as genai
import io
import time
from docx import Document

# ----------------- CONFIGURACIÓN -----------------
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=GEMINI_API_KEY)

# Modelo a utilizar (flash tiene buena relación calidad/cupo gratuito)
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

# ----------------- FUNCIONES DE EXTRACCIÓN -----------------

def extraer_texto_docx(file_bytes):
    """Extrae texto de un Word (.docx) con python-docx."""
    try:
        doc = Document(io.BytesIO(file_bytes))
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        raise RuntimeError(f"Error al procesar Word: {e}")

def extraer_texto_varios_archivos(archivos_subidos, modelo):
    """
    Envía todos los archivos (PDF, imágenes) en una sola solicitud a Gemini
    para extraer el texto. Así se evita exceder el límite gratuito de 5 RPM.
    Retorna una lista de tuplas (nombre_archivo, texto_extraido).
    Si un archivo es DOCX, se extrae localmente.
    """
    partes_solicitud = []
    lista_docx = []  # (nombre, bytes) para procesar después
    mapeo_indices = []  # para saber a qué nombre corresponde cada parte

    for archivo in archivos_subidos:
        nombre = archivo.name
        mime = archivo.type
        bytes_data = archivo.read()

        if mime == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            lista_docx.append((nombre, bytes_data))
            continue  # se procesa aparte

        # Solo PDF e imágenes van a Gemini
        if mime not in ["application/pdf", "image/jpeg", "image/png", "image/jpg"]:
            st.warning(f"Tipo de archivo no soportado: {nombre} ({mime})")
            continue

        try:
            # Añadir parte con el archivo y una etiqueta para identificarlo
            partes_solicitud.append({
                "mime_type": mime,
                "data": bytes_data
            })
            mapeo_indices.append(nombre)
        except Exception as e:
            st.warning(f"No se pudo leer el archivo {nombre}: {e}")

    # Si no hay nada que enviar a Gemini, devolvemos solo los DOCX
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

    # Preparar mensaje para Gemini con todos los archivos
    # Incluimos una instrucción clara para que devuelva el texto en un formato que podamos separar por documento
    prompt_instrucciones = (
        "Extrae TODO el texto de cada uno de los siguientes documentos, conservando fechas, nombres y números tal cual aparecen. "
        "Devuelve el resultado en este formato EXACTO:\n\n"
        "--- DOCUMENTO: nombre_real_del_archivo ---\n"
        "[texto extraído]\n\n"
        "Repite esa estructura para cada documento, en el mismo orden en que se proporcionan."
    )

    # Construir partes completas
    contenido = [prompt_instrucciones] + partes_solicitud

    try:
        respuesta = modelo.generate_content(contenido)
        texto_completo = respuesta.text
    except Exception as e:
        error_str = str(e)
        if "429" in error_str:
            st.warning("Límite de cuota alcanzado incluso en la solicitud combinada. Esperando 20 segundos...")
            time.sleep(20)
            try:
                respuesta = modelo.generate_content(contenido)
                texto_completo = respuesta.text
            except Exception as e2:
                st.error(f"Error definitivo al extraer texto con Gemini: {e2}")
                # Devolver lo que tengamos de DOCX
                textos = []
                for nombre, data in lista_docx:
                    try:
                        txt = extraer_texto_docx(data)
                        textos.append((nombre, txt))
                    except:
                        textos.append((nombre, "[Error]"))
                return textos
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

    # Parsear la respuesta para separar por documentos
    # Buscamos las marcas "--- DOCUMENTO: ... ---"
    textos_resultantes = []
    bloques = texto_completo.split("--- DOCUMENTO: ")
    # El primer bloque puede estar vacío o contener texto antes del primer documento
    if bloques and bloques[0].strip() == "":
        bloques = bloques[1:]  # ignorar prefacio vacío

    for bloque in bloques:
        # Cada bloque debería tener el formato "nombre ---\ntexto"
        if "---" in bloque:
            partes = bloque.split("---", 1)
            nombre = partes[0].strip()
            texto = partes[1].strip() if len(partes) > 1 else ""
            textos_resultantes.append((nombre, texto))
        else:
            # Si no coincide, lo tomamos como documento sin nombre claro
            textos_resultantes.append(("Desconocido", bloque.strip()))

    # Agregar los documentos DOCX procesados localmente
    for nombre, data in lista_docx:
        try:
            txt = extraer_texto_docx(data)
            textos_resultantes.append((nombre, txt))
        except Exception as e:
            st.warning(f"Error con Word {nombre}: {e}")
            textos_resultantes.append((nombre, "[Error al extraer]"))

    return textos_resultantes

# ----------------- LLAMADA PRINCIPAL -----------------

def generar_texto_legal(fuentes, texto_base, modelo):
    prompt = PROMPT_PLANTILLA.format(fuentes=fuentes, texto_base=texto_base)
    try:
        respuesta = modelo.generate_content(prompt)
        return respuesta.text
    except Exception as e:
        if "429" in str(e):
            st.warning("Límite alcanzado. Reintentando en 20 segundos...")
            time.sleep(20)
            respuesta = modelo.generate_content(prompt)
            return respuesta.text
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
            # Extraer texto de todos los archivos en una sola solicitud
            lista_textos = extraer_texto_varios_archivos(archivos_fuente, model)

            # Construir el texto combinado para las fuentes
            fuentes_completo = ""
            for nombre, txt in lista_textos:
                fuentes_completo += f"--- DOCUMENTO: {nombre} ---\n{txt}\n\n"

        with st.spinner("Generando texto legal (esto puede tomar unos segundos)..."):
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
