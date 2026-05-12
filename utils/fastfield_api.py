"""
FastField REST API v3 — descarga automática de fotos desde submissions.

Credenciales almacenadas en st.secrets:
    fastfield_email    = "user@empresa.com"
    fastfield_password = "password"
    fastfield_org_id   = "ORG_ID"   (opcional, si la cuenta tiene varias orgs)
"""
import base64
import io
import requests


BASE_URL = "https://api.fastfieldforms.com/services/v3"
TIMEOUT  = 20  # segundos


def _basic_auth_header(email: str, password: str) -> str:
    token = base64.b64encode(f"{email}:{password}".encode()).decode()
    return f"Basic {token}"


def authenticate(email: str, password: str, org_id: str = "") -> str:
    """
    Autentica con FastField y retorna el sessionToken.
    Lanza RuntimeError si falla.
    """
    headers = {"Authorization": _basic_auth_header(email, password)}
    if org_id:
        headers["X-Gatekeeper-OrgId"] = org_id

    resp = requests.post(f"{BASE_URL}/authenticate", headers=headers, timeout=TIMEOUT)
    if resp.status_code != 200:
        raise RuntimeError(f"FastField auth falló ({resp.status_code}): {resp.text[:200]}")

    return resp.json()["sessionToken"]


def get_photo_bytes(filename: str, session_token: str) -> bytes | None:
    """
    Dado el nombre de archivo de una foto (de multiphoto_picker_1),
    devuelve los bytes de la imagen o None si no se puede descargar.
    """
    # Paso 1: obtener URL autenticada de descarga
    headers = {"X-Gatekeeper-SessionToken": session_token}
    resp = requests.get(
        f"{BASE_URL}/media/download",
        params={"key": filename},
        headers=headers,
        timeout=TIMEOUT,
    )
    if resp.status_code != 200:
        return None

    download_url = resp.json().get("downloadUrl", "")
    if not download_url:
        return None

    # Paso 2: descargar la imagen desde la URL firmada
    img_resp = requests.get(download_url, timeout=30)
    if img_resp.status_code != 200:
        return None

    return img_resp.content


def download_submission_photos(photo_filenames: list[str],
                               email: str,
                               password: str,
                               org_id: str = "") -> list[bytes | None]:
    """
    Descarga todas las fotos de un submission.
    Retorna lista de bytes (o None si falló) en el mismo orden.
    """
    if not photo_filenames:
        return []

    try:
        token = authenticate(email, password, org_id)
    except RuntimeError:
        return [None] * len(photo_filenames)

    return [get_photo_bytes(fn, token) for fn in photo_filenames]
