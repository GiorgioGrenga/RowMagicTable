"""Gemini-powered extraction of tables from images and PDF documents."""

from __future__ import annotations

import io
import json
import logging
import math
from typing import Any

import pymupdf
from google import genai
from google.genai import errors as genai_errors
from google.genai import types
from PIL import Image, UnidentifiedImageError

MAX_DOCUMENT_BYTES = 15 * 1024 * 1024
MAX_PDF_PAGES = 40
SUPPORTED_IMAGES = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}
MODEL_NAMES = ("gemini-3.8-flash", "gemini-2.5-flash")
logger = logging.getLogger(__name__)


class TableExtractionError(Exception):
    """Raised when a file cannot be converted into a valid table."""

    def __init__(self, message: str, http_status: int = 400) -> None:
        super().__init__(message)
        self.http_status = http_status


def _api_error_message(error: genai_errors.APIError) -> str:
    status = str(error.status or "").upper()
    if status in {"API_KEY_INVALID", "UNAUTHENTICATED"} or error.code == 401:
        return (
            "Gemini ha rifiutato la chiave API. Verifica che GEMINI_API_KEY sia "
            "configurata correttamente nelle variabili del servizio e ridistribuisci."
        )
    if status in {"PERMISSION_DENIED", "PERMISSIONDENIED"} or error.code == 403:
        return (
            "La chiave Gemini non ha accesso al modello. Controlla progetto, "
            "restrizioni della chiave e disponibilità dell'API Gemini."
        )
    if status in {"RESOURCE_EXHAUSTED", "QUOTA_EXCEEDED"} or error.code == 429:
        return (
            "La quota o il limite di richieste Gemini è stato raggiunto. Controlla "
            "la quota del progetto in Google AI Studio e riprova più tardi."
        )
    if status in {"NOT_FOUND", "MODEL_NOT_FOUND"} or error.code == 404:
        return (
            "Il modello Gemini non è disponibile per questa richiesta. Controlla "
            "che il modello configurato sia abilitato per la tua chiave."
        )
    if status in {"INVALID_ARGUMENT", "FAILED_PRECONDITION"} or error.code == 400:
        return (
            "Gemini ha rifiutato il documento o il formato della richiesta. Prova "
            "un'immagine più piccola o un PDF con meno pagine."
        )
    if error.code is not None and error.code >= 500:
        return "Il servizio Gemini è temporaneamente indisponibile. Riprova tra poco."
    return (
        "Gemini ha rifiutato la richiesta. Verifica la configurazione della chiave "
        "e la disponibilità del servizio, poi riprova."
    )

def _validate_document(file_bytes: bytes, filename: str) -> str:
    if not file_bytes:
        raise TableExtractionError("Il file caricato è vuoto.")
    if len(file_bytes) > MAX_DOCUMENT_BYTES:
        raise TableExtractionError("Il file supera il limite di 15 MB.")

    extension = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if extension == "pdf":
        try:
            with pymupdf.open(stream=file_bytes, filetype="pdf") as document:
                if document.needs_pass:
                    raise TableExtractionError("Il PDF è protetto da password.")
                if document.page_count == 0:
                    raise TableExtractionError("Il PDF non contiene pagine.")
                if document.page_count > MAX_PDF_PAGES:
                    raise TableExtractionError(
                        f"Il PDF supera il limite di {MAX_PDF_PAGES} pagine."
                    )
        except TableExtractionError:
            raise
        except (pymupdf.FileDataError, RuntimeError, OSError) as exc:
            raise TableExtractionError("Il PDF non è valido o non è leggibile.") from exc
        return "application/pdf"

    extension = f".{extension}"
    media_type = SUPPORTED_IMAGES.get(extension)
    if media_type is None:
        raise TableExtractionError("Formato non supportato. Usa PNG, JPG, JPEG o PDF.")
    try:
        with Image.open(io.BytesIO(file_bytes)) as image:
            image.verify()
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        raise TableExtractionError("L'immagine non è valida o non è leggibile.") from exc
    try:
        with Image.open(io.BytesIO(file_bytes)) as image:
            expected_format = {
                ".png": "PNG",
                ".jpg": "JPEG",
                ".jpeg": "JPEG",
            }[extension]
            if image.format != expected_format:
                raise TableExtractionError(
                    "L'estensione del file non corrisponde al formato dell'immagine."
                )
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        raise TableExtractionError("L'immagine non è valida o non è leggibile.") from exc
    return media_type


def _normalize_table(data: Any) -> dict[str, list[Any]]:
    if not isinstance(data, dict):
        raise TableExtractionError("La risposta AI non contiene una tabella valida.")

    headers = data.get("headers")
    rows = data.get("rows")
    if not isinstance(headers, list) or not headers:
        raise TableExtractionError("Non è stata trovata alcuna intestazione.")
    if not isinstance(rows, list):
        raise TableExtractionError("La tabella estratta non ha righe valide.")

    cleaned_headers: list[str] = []
    seen: dict[str, int] = {}
    for index, header in enumerate(headers):
        if not isinstance(header, str):
            raise TableExtractionError("Le intestazioni devono essere testo.")
        name = header.strip() or f"Colonna {index + 1}"
        seen[name] = seen.get(name, 0) + 1
        if seen[name] > 1:
            name = f"{name} ({seen[name]})"
        cleaned_headers.append(name)

    cleaned_rows = []
    for row_number, row in enumerate(rows, start=1):
        if not isinstance(row, list) or len(row) != len(cleaned_headers):
            raise TableExtractionError(
                f"La riga {row_number} non corrisponde alle intestazioni."
            )
        for value in row:
            if value is not None and not isinstance(value, (str, int, float, bool)):
                raise TableExtractionError(
                    f"La riga {row_number} contiene un valore non supportato."
                )
            if isinstance(value, float) and not math.isfinite(value):
                raise TableExtractionError(
                    f"La riga {row_number} contiene un numero non valido."
                )
        cleaned_rows.append(row)

    return {"headers": cleaned_headers, "rows": cleaned_rows}


def extract_table(
    file_bytes: bytes, filename: str, api_key: str | None = None
) -> dict[str, list[Any]]:
    """Extract and validate a table from an image or PDF using Gemini."""
    if not api_key or not api_key.strip():
        raise TableExtractionError(
            "Chiave Gemini mancante. Configura GEMINI_API_KEY nelle variabili "
            "d'ambiente del server."
        )

    media_type = _validate_document(file_bytes, filename)
    prompt = (
        "Estrai con precisione la tabella principale dal documento. Correggi solo "
        "piccoli refusi o disallineamenti, non inventare valori mancanti. Se la tabella "
        "continua su più pagine, unisci le righe nell'ordine originale e includi "
        "l'intestazione una sola volta. Restituisci esclusivamente un oggetto JSON nel "
        'formato {"headers":["Colonna 1"],"rows":[["valore"]]}. Usa null per le celle '
        'vuote o illeggibili. Se non trovi una tabella, restituisci '
        '{"headers":[],"rows":[]}.'
    )
    client = genai.Client(api_key=api_key)
    response = None
    for model_index, model_name in enumerate(MODEL_NAMES):
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[
                    types.Part.from_bytes(data=file_bytes, mime_type=media_type),
                    prompt,
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.1,
                    max_output_tokens=8192,
                ),
            )
            break
        except genai_errors.APIError as exc:
            is_missing_model = (
                exc.code == 404
                or str(exc.status or "").upper() in {"NOT_FOUND", "MODEL_NOT_FOUND"}
            )
            if is_missing_model and model_index + 1 < len(MODEL_NAMES):
                logger.info(
                    "Gemini model %s is unavailable; trying the configured fallback.",
                    model_name,
                )
                continue

            logger.warning(
                "Gemini request failed for model %s (HTTP status %s, provider status %s).",
                model_name,
                exc.code,
                exc.status,
            )
            raise TableExtractionError(
                _api_error_message(exc),
                http_status=503,
            ) from exc

    if response is None:
        raise TableExtractionError(
            "Nessun modello Gemini configurato ha restituito una risposta.",
            http_status=503,
        )

    if not response.text:
        raise TableExtractionError("Gemini ha restituito una risposta vuota.")
    try:
        data = json.loads(response.text)
    except json.JSONDecodeError as exc:
        raise TableExtractionError("Gemini ha restituito una risposta non valida.") from exc

    if data == {"headers": [], "rows": []}:
        raise TableExtractionError("Non è stata trovata alcuna tabella nel file.")
    return _normalize_table(data)
