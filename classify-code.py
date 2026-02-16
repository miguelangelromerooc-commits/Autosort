# AutoSort AI for Google Drive


from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import os
import io
import openai
import nltk
import openpyxl
from docx import Document
import fitz  # librería para interpretar docs en pdf.
import google.generativeai as genai

from collections import Counter
from nltk.corpus import stopwords
from collections import defaultdict
from openai import ChatCompletion
import requests
from googleapiclient.discovery import build as build_sheet
from datetime import datetime

# importaciones para la clasificación por embeddings
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np




SPREADSHEET_ID = 'Colocar aquí ID de la hoja de cálculo donde se hará informe' # ID de la hoja de cálculo de Google Sheets donde se hará el informe.
SHEET_NAME = 'Informe' # nombre de la hoja dentro del spreadsheet donde se hará el informe (debe coincidir con el nombre de la hoja en Google Sheets) 

nltk.download("punkt")
nltk.download("stopwords")

# Enlace entre el modelo de IA a utilizar y Google Drive
OPENAI_API_KEY = 'Coloca aquí tu clave de API de OpenAI'  # Coloca aquí tu clave de API del modelo de IA a utilizar
genai.configure(api_key=OPENAI_API_KEY)

# SCOPES define el nivel de acceso requerido, para este caso usamos el acceso a DRIVE donde se contendrán todos los documentos y 
# a SPREADSHEETS para el llenado automático de la hoja de cálculo del informe final sobre la clasificación.
SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]

# ID de la carpeta principal donde están los archivos (antes de clasificar)
MAIN_FOLDER_ID = "Colocar aquí el ID de la carpeta principal en Google Drive"

# Umbral mínimo de coincidencias (estas son las palabras clave que encuentre en los documentos analizados para tomar como punto de referencia)
CONFIDENCE_THRESHOLD = 3

# Modelo de embeddings utilizado 

EMB_MODEL_NAME = "sentence-transformers/all-MiniLM-L6-v2"
emb_model = SentenceTransformer(EMB_MODEL_NAME) # libreria para instanciar el modelo de embeddings, transformando texto en vectores numéricos.

# Umbral mínimo para similitud coseno, si es menor a 35 % nos quedaría como "sin clasificar"
EMB_SIM_THRESHOLD = 0.35


# método para diseñar el informa de clasificación en Google Sheets

def append_to_sheet(creds, data):
    try:
        sheet_service = build_sheet('sheets', 'v4', credentials=creds)
        values = [data]
        body = {'values': values}
        sheet_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!A1",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body=body
        ).execute()
    except Exception as e:
        print(f"Error al escribir en la hoja de cálculo: {e}")


# módulo para la autenticación en Google Drive

def authenticate_google_drive():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('drive', 'v3', credentials=creds)


# Módulo para descargar los archivos almacenados en Google Drive

def download_file(service, file_id, file_name, mime_type=None):
    file_path = os.path.join('temp', file_name)
    os.makedirs('temp', exist_ok=True)

    try:
        if mime_type and mime_type.startswith('application/vnd.google-apps'): 
            if mime_type == 'application/vnd.google-apps.document':
                export_mime_type = 'application/pdf'
                file_path = file_path.replace(".doc", ".pdf")
            elif mime_type == 'application/vnd.google-apps.spreadsheet':
                export_mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                file_path = file_path.replace(".xls", ".xlsx")
            elif mime_type == 'application/vnd.google-apps.presentation':
                export_mime_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                file_path = file_path.replace(".ppt", ".pptx")
            else:
                print(f"No se soporta la exportación para este tipo: {mime_type}")
                return None

            request = service.files().export_media(fileId=file_id, mimeType=export_mime_type)
        else:
            request = service.files().get_media(fileId=file_id)

        with io.FileIO(file_path, 'wb') as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
        return file_path

    except Exception as e:
        print(f"Error al descargar el archivo {file_name}: {e}")
        return None



# módulo para extraer texto de los archivos descargados.

def extract_text(file_path):
    if file_path.endswith('.docx'):
        try:
            doc = Document(file_path)
            text = '\n'.join([para.text for para in doc.paragraphs])
            return text
        except Exception as e:
            print(f"Error al leer .docx: {file_path} - {e}")
            return ""

    elif file_path.endswith('.xlsx'):
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            text = ""
            for sheet in wb.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value:
                            text += str(cell.value) + " "
            return text
        except Exception as e:
            print(f"Error al leer .xlsx: {file_path} - {e}")
            return ""

    elif file_path.endswith('.pdf'):
        try:
            text = ""
            with fitz.open(file_path) as doc:
                for page in doc:
                    text += page.get_text()
            return text
        except Exception as e:
            print(f"Error al leer .pdf: {file_path} - {e}")
            return ""

    else:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='ISO-8859-1') as f:
                return f.read()


# Módulo para clasificar por palabras clave. Se indica el score según las coincidencias. 

def classify_document_with_score(content, keywords):
    words = nltk.word_tokenize((content or "").lower())
    filtered_words = [word for word in words if word.isalnum()]
    word_freq = Counter(filtered_words)

    category_scores = defaultdict(int)
    for category, category_keywords in keywords.items():
        for keyword in category_keywords:
            category_scores[category] += word_freq.get(str(keyword).lower(), 0)

    if not category_scores:
        return "Sin clasificar", 0

    best_category = max(category_scores, key=category_scores.get, default=None)
    best_score = category_scores.get(best_category, 0)

    if best_score < CONFIDENCE_THRESHOLD:
        return "Sin clasificar", best_score

    return best_category, best_score


# Módulo de clasificación por LLM, en este caso se usó GPT-4.1-mini. Este módulo es el que debe ajustarse según el modelo que se desee utilizar.

def classify_with_gpt(content, keywords):
    openai.api_key = OPENAI_API_KEY
    #prompt enviado al modelo, este puede ajustarse según las necesidades, incluso cambiar el idioma. 
    prompt = f"""
    Clasifica el siguiente contenido en una de las categorías disponibles. 
    Categorías: {', '.join(keywords.keys())}
    Contenido:
    {content}

    Responde únicamente con el nombre de la categoría.
    """

    try:
        response = openai.ChatCompletion.create(
            model="gpt-4.1-mini",
            messages=[
                {"role": "system", "content": "Eres un asistente experto en clasificación de contenido educativo."},
                {"role": "user", "content": prompt}
            ]
        )
        category = response['choices'][0]['message']['content'].strip()
        return category if category in keywords else "Sin clasificar"
    except Exception as e:
        print(f"Error al usar GPT para clasificar: {e}")
        return "Sin clasificar"



# módulo de clasificación por embeddings
#  las instruccines pueden ser ajustadas según las necesidades, incluso cambiar el idioma debido a que el modelo trabaja con múltiples idiomas.
def build_category_prototypes(categories: dict) -> dict:
    """
    Construye un texto prototipo por categoría usando: nombre + keywords. 
    """
    prototypes = {}
    for cat, kws in categories.items():
        # Quita duplicados manteniendo orden y limita para evitar textos muy largos
        cleaned = []
        seen = set()
        for k in kws:
            kl = str(k).lower().strip()
            if kl and kl not in seen:
                seen.add(kl)
                cleaned.append(kl)
        cleaned = cleaned[:60]
        prototypes[cat] = f"Categoría: {cat}. Términos relacionados: " + ", ".join(cleaned)
    return prototypes

# módulo para normalizar el texto antes de generar embeddings
def normalize_for_embedding(text: str, max_chars: int = 6000) -> str:
    text = (text or "").strip()
    if len(text) > max_chars:
        text = text[:max_chars]
    return text

# módulo para preparar los embeddings de las categorías
def prepare_category_embeddings(prototypes: dict):
    cat_names = list(prototypes.keys())
    cat_texts = [prototypes[c] for c in cat_names]
    cat_embs = emb_model.encode(cat_texts, normalize_embeddings=True)
    return cat_names, cat_embs

# módulo para clasificar un documento usando embeddings
def classify_with_embeddings(content: str, cat_names, cat_embs):
    text = normalize_for_embedding(content)
    if not text:
        return "Sin clasificar", 0.0

    doc_emb = emb_model.encode([text], normalize_embeddings=True)
    sims = cosine_similarity(doc_emb, cat_embs)[0]
    best_idx = int(np.argmax(sims))
    best_score = float(sims[best_idx])
    best_cat = cat_names[best_idx]

    if best_score < EMB_SIM_THRESHOLD:
        return "Sin clasificar", best_score

    return best_cat, best_score


# módulo que determiana la categoría final según los resultados de los tres métodos.

def decide_final_category(kw_cat, llm_cat, emb_cat, emb_score, kw_ok: bool):
    def valid(cat):
        return cat and cat != "Sin clasificar"

    votes = [c for c in [kw_cat, llm_cat, emb_cat] if valid(c)]
    if votes:
        counts = Counter(votes)
        top_cat, top_n = counts.most_common(1)[0]
        if top_n >= 2:
            return top_cat

    # cuando los tres clasificadores difieren la prioridad es:  1) LLM, 2) y 3) Keywords
    if valid(llm_cat):
        return llm_cat
    if valid(emb_cat) and emb_score >= EMB_SIM_THRESHOLD:
        return emb_cat
    if valid(kw_cat) and kw_ok:
        return kw_cat

    return "Sin clasificar"


# módulo para mover archivos entre carpetas en Google Drive (al método seleccionado)


def move_file(service, file_id, folder_id):
    file = service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents', []))

    service.files().update(
        fileId=file_id,
        addParents=folder_id,
        removeParents=previous_parents,
        fields='id, parents'
    ).execute()



# módulo que obtiene la carpera en donde se haya decidido clasificar el archivo, o la crea si no existe.

def get_or_create_folder(service, folder_name, parent_id):
    query = f"'{parent_id}' in parents and name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder'"
    results = service.files().list(q=query, fields="files(id, name, trashed)").execute()
    items = results.get('files', [])

    for item in items:
        if not item.get("trashed", False):
            return item["id"]

    folder_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    folder = service.files().create(body=folder_metadata, fields='id').execute()
    return folder['id']


# módulo que contiene el método principal 

def main():
    service = authenticate_google_drive()
    #Categorías creadas manualmente y su set de palabras clave, estas pueden modificarse según el entorno en que vaya a aplicarse.
    categories = {
        "Seguridad en la PC": [
            "antivirus", "seguridad", "firewall", "ciberseguridad", "contraseñas", "Avast", "Panda", "Eset Online Scanner",
            "privacidad", "Norton", "malware", "phishing", "protección", "seguridad informática", "virus", "antispyware",
            "antivirus/antispyware"
        ],

        "Búsqueda de información en la web": [
            "buscadores especializados", "bibliotecas digitales", "colecciones digitales",
            "fuentes confiables", "artículos especializados", "tesis", "revistas",
            "trabajos académicos", "información confiable", "investigación", "información veraz", "evaluar información",
            "confiabilidad", "publicaciones en línea", "compartir información",
            "búsqueda efectiva", "búsquedas en línea", "fuentes de información",
            "calidad de la información", "evaluación de publicaciones", "fuentes académicas",
            "relevancia", "precisión", "autoridad", "actualidad", "sitios confiables",
            "fake news", "veracidad de contenido", "información en Internet", "calidad"
        ],

        "Creación de audio": [
            "audio", "grabación de sonido", "pista musical", "calidad de grabación",
            "efectos de sonido", "edición de audio", "duración del audio",
            "reproductor de audio", "licencia libre", "música de fondo", "producción de audio",
            "podcast", "mezcla de audio", "mp3", "archivo mp3"
        ],

        "Navegadores web": [
            "navegador", "Google Chrome", "Firefox", "Microsoft Edge", "Safari",
            "extensiones del navegador", "extensiones", "historial de navegación", "cookies",
            "modo incógnito", "pestañas", "descargas en navegadores", "marcadores", "descargas", "historial"
        ],

        "Buscadores web": [
            "Google", "Yahoo", "Bing", "motor de búsqueda", "consultas de búsqueda",
            "resultados de búsqueda", "keywords", "buscadores académicos", "palabras clave"
        ],

        "Ética en el contexto digital": [
            "ética", "moral", "piratería", "propiedad intelectual", "plagio",
            "citas", "referencias", "Creative Commons", "normas APA", "normas Chicago",
            "licencias", "derechos de autor", "comportamiento ético",
            "uso ético", "Internet", "contenido digital", "reconocimiento del autor",
            "plagio académico", "fuentes confiables", "Google Académico",
            "manual de normas", "citación", "responsabilidad", "acceso a información",
            "estilo APA", "licencias digitales", "Openverse"
        ],

        "Creación de imágenes digitales": [
            "edición de imágenes", "diseño gráfico", "Photopea", "herramientas de diseño",
            "diseño de imágenes", "capas", "filtros", "degradados", "Photoshop",
            "diseño visual", "imágenes digitales", "composición de imágenes",
            "texto en imágenes", "pinceles", "estilos de capas", "formato PSD",
            "formato PNG", "Adobe", "rueda de colores", "imágenes para sitios web"
        ],

        "Uso básico de la computadora": [
            "computadora", "teclado", "ratón", "pantalla", "hardware", "software",
            "dispositivos", "Windows", "explorador de archivos", "escritorio",
            "atajos de teclado", "panel de control", "cuentas de usuario",
            "configuración del sistema", "almacenamiento",
            "byte", "kilobyte", "megabyte", "gigabyte", "terabyte", "petabyte",
            "sincronizar servicios", "archivos", "carpetas"
        ],

        "Uso de la hoja de cálculo": [
            "hoja de cálculo", "Excel", "Google Sheets", "funciones", "gráficos",
            "fórmulas", "ordenar datos", "celdas", "formato condicional", "referencia de celda",
            "análisis de datos", "porcentajes", "gráficas"
        ],

        "Creación de infografías": [
            "infografía", "diseño visual", "síntesis", "información visual", "presentación visual",
            "herramientas de diseño", "Canva", "visualización de datos", "diagramas", "comunicación visual",
            "texto e imágenes", "fotografías", "presentar información", "tema específico", "contenido en forma de síntesis"
        ],

        "Creación de presentaciones digitales": [
            "presentaciones", "PowerPoint", "transiciones", "plantillas",
            "gráficos", "presentaciones dinámicas", "Slides", "exposición", "diseño de diapositivas"
        ],

        "Creación de páginas web": [
            "sitio web", "páginas web", "Google Sites", "desarrollo web", "diseño de sitios", "hipervínculos",
            "diseño web", "WordPress", "Wix", "contenido multimedia", "sitio", "publicación",
            "dominio", "enlaces externos", "navegación sencilla", "páginas"
        ],

        "Sistemas operativos": [
            "Windows", "Linux", "MacOS", "Ubuntu", "gestión de archivos",
            "configuración de sistema", "terminal", "comandos", "explorador de archivos"
        ],

        "Trabajo colaborativo": [
            "colaboración", "equipo", "herramientas colaborativas", "Google Drive",
            "proyectos compartidos", "trabajo en equipo", "comunicación digital",
            "documentos colaborativos", "Slack", "Microsoft Teams"
        ],

        "Creación de vídeos digitales": [
            "video", "vídeos digitales", "animación", "PowToon", "producción de video",
            "edición de video", "transiciones", "pista musical", "subir a YouTube",
            "publicar video", "plataforma de video", "código HTML", "narrar historias",
            "audio en video", "crear animaciones", "formato MP4", "formato AVI",
            "formato WMV", "streaming de video", "plataformas de video", "Vimeo",
            "DailyMotion", "visualización de video", "contenido multimedia"
        ],

        "Navegadores y buscadores web": [
            "navegadores", "buscadores", "web", "Google", "Yahoo",
            "Bing", "herramientas de búsqueda", "consultas web", "resultados web", "Chrome"
        ]
    }

    # prepara los embeddings de las categorías que se tienen (una sola vez)
    prototypes = build_category_prototypes(categories)
    cat_names, cat_embs = prepare_category_embeddings(prototypes)

    # se lleva a cabo la creación de carpetas para cada categoría o recupera su ID si ya existen
    folder_ids = {}
    for category in categories.keys():
        folder_id = get_or_create_folder(service, category, MAIN_FOLDER_ID)
        folder_ids[category] = folder_id

    # Crear carpeta para 'Sin clasificar'
    folder_ids['Sin clasificar'] = get_or_create_folder(service, 'Sin clasificar', MAIN_FOLDER_ID)

    # Procesar archivos en la carpeta principal
    results = service.files().list(
        q=f"'{MAIN_FOLDER_ID}' in parents",
        pageSize=100,
        fields="files(id, name, mimeType)"
    ).execute()
    items = results.get('files', [])

    if not items:
        print("No se encontraron archivos en la carpeta principal.")
        return

    for item in items:
        file_id, file_name, mime_type = item['id'], item['name'], item.get('mimeType', '')

        if mime_type.startswith("application/vnd.google-apps.folder"):
            continue

        file_path = download_file(service, file_id, file_name, mime_type)
        # 
        if file_path:
            content = extract_text(file_path)
            #aplicando los tres métodos de clasificación para cada archivo, obteniendo el resultado de cada uno y su score o nivel de confianza.
            # 1) Keywords
            manual_category, kw_score = classify_document_with_score(content, categories)
            kw_ok = kw_score >= CONFIDENCE_THRESHOLD

            # 2) Embeddings
            emb_category, emb_score = classify_with_embeddings(content, cat_names, cat_embs)

            # 3) LLM
            try:
                gpt_category = classify_with_gpt(content, categories)
                print(f"Analyzing the file: {file_name}")
            except Exception as e:
                print(f"Error modelo IA: {e}")
                gpt_category = "Sin clasificar"

            print(
                f"LLM: {gpt_category} | KW: {manual_category} (score={kw_score}) | "
                f"EMB: {emb_category} (sim={emb_score:.3f})"
            )

            # Decisión final de clasificación según los resultados obtenidos por los tres métodos.
            final_category = decide_final_category(
                kw_cat=manual_category,
                llm_cat=gpt_category,
                emb_cat=emb_category,
                emb_score=emb_score,
                kw_ok=kw_ok
            )

            # Mover archivo a la carpeta correspondiente según la categoría final decidida.
            move_file(service, file_id, folder_ids[final_category])
            os.remove(file_path)
            print(f"'{file_name}' clasificado como '{final_category}'.")

            # Datos que van y se guardan en el informe de clasificación en Google Sheets
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            append_to_sheet(service._http.credentials, [
                file_name,
                manual_category,
                kw_score,
                emb_category,
                round(emb_score, 4),
                gpt_category,
                final_category,
                now
            ])


if __name__ == '__main__':
    main()
