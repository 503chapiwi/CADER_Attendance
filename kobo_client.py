"""Conexión con KoboToolbox y normalización de datos del formulario Planillas MAGA Toto."""

import io
import re
import unicodedata

import pandas as pd
import requests

KOBO_URL = "https://kf.kobotoolbox.org"
ASSET_UID = "aMAY4yQjPMm3Nvw6w3AdEv"

PARTICIPANTES_FILENAME = "participantes_toto_2026.csv"
PROMOTORES_FILENAME = "toto_promoters.csv"


def _headers(token):
    return {"Authorization": f"Token {token}"}


def _slug(text):
    """Normaliza un texto para comparaciones: sin acentos, minúsculas, _ por espacios."""
    if text is None:
        return ""
    text = str(text)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower().strip()
    text = re.sub(r"[\s]+", "_", text)
    text = re.sub(r"[^a-z0-9_]", "", text)
    return text


def fetch_asset(token):
    """Descarga la definición del formulario (para mapear códigos -> etiquetas)."""
    r = requests.get(f"{KOBO_URL}/api/v2/assets/{ASSET_UID}.json", headers=_headers(token), timeout=60)
    r.raise_for_status()
    return r.json()


def build_label_maps(asset):
    """Devuelve {list_name: {codigo: etiqueta}} y {campo: list_name}."""
    content = asset["content"]
    choice_maps = {}
    for ch in content.get("choices", []):
        labels = ch.get("label") or []
        label = next((l for l in labels if l), None) or ch["name"]
        choice_maps.setdefault(ch["list_name"], {})[ch["name"]] = label
    field_lists = {}
    for q in content.get("survey", []):
        name = q.get("name") or q.get("$autoname")
        if name and q.get("select_from_list_name"):
            field_lists[name] = q["select_from_list_name"]
    return choice_maps, field_lists


def label_for(choice_maps, field_lists, field, value):
    """Traduce un valor codificado a su etiqueta. Tolera datos de versiones viejas
    del formulario que ya venían con etiquetas en vez de códigos."""
    if value is None or value == "":
        return ""
    value = str(value)
    lst = field_lists.get(field)
    if lst and lst in choice_maps:
        cmap = choice_maps[lst]
        if value in cmap:
            return cmap[value]
        vslug = _slug(value)
        for code, label in cmap.items():
            if _slug(code) == vslug or _slug(label) == vslug:
                return label
    return value


def fetch_submissions(token):
    """Descarga todas las respuestas del formulario, con claves normalizadas
    (solo el último segmento de la ruta del grupo)."""
    results = []
    url = f"{KOBO_URL}/api/v2/assets/{ASSET_UID}/data.json?limit=500"
    while url:
        r = requests.get(url, headers=_headers(token), timeout=120)
        r.raise_for_status()
        d = r.json()
        results.extend(d.get("results", []))
        url = d.get("next")
    return [_normalize_keys(s) for s in results]


def _normalize_keys(record):
    out = {}
    for k, v in record.items():
        short = k.split("/")[-1]
        if isinstance(v, list) and v and isinstance(v[0], dict):
            v = [_normalize_keys(item) for item in v]
        out[short] = v
    return out


def _fetch_media_csv(token, filename):
    r = requests.get(f"{KOBO_URL}/api/v2/assets/{ASSET_UID}/files.json", headers=_headers(token), timeout=60)
    r.raise_for_status()
    for f in r.json().get("results", []):
        if f.get("metadata", {}).get("filename") == filename:
            rc = requests.get(f["content"], headers=_headers(token), timeout=120)
            rc.raise_for_status()
            return pd.read_csv(io.BytesIO(rc.content), dtype=str)
    raise FileNotFoundError(f"No se encontró {filename} entre los archivos del formulario en Kobo.")


def fetch_participantes(token):
    """Base de datos de participantes CADER adjunta al formulario en Kobo."""
    return _fetch_media_csv(token, PARTICIPANTES_FILENAME)


def fetch_promotores(token):
    """Listado de promotores (cader_id -> nombre) adjunto al formulario en Kobo."""
    return _fetch_media_csv(token, PROMOTORES_FILENAME)
