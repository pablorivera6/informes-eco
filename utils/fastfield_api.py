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


def _base_headers(subscription_key: str) -> dict:
    """Headers comunes para todas las llamadas a la API."""
    h = {}
    if subscription_key:
        h["Ocp-Apim-Subscription-Key"] = subscription_key
    return h


def authenticate(email: str, password: str,
                 org_id: str = "", subscription_key: str = "") -> str:
    """
    Autentica con FastField y retorna el sessionToken.
    Lanza RuntimeError si falla.
    """
    headers = {
        **_base_headers(subscription_key),
        "Authorization": _basic_auth_header(email, password),
    }
    if org_id:
        headers["X-Gatekeeper-OrgId"] = org_id

    resp = requests.post(f"{BASE_URL}/authenticate", headers=headers, timeout=TIMEOUT)
    if resp.status_code != 200:
        raise RuntimeError(f"FastField auth falló ({resp.status_code}): {resp.text[:300]}")

    return resp.json()["sessionToken"]


def get_photo_bytes(filename: str, session_token: str,
                    subscription_key: str = "") -> bytes | None:
    """
    Dado el nombre de archivo de una foto (de multiphoto_picker_1),
    devuelve los bytes de la imagen o None si no se puede descargar.
    """
    # Paso 1: obtener URL autenticada de descarga
    headers = {
        **_base_headers(subscription_key),
        "X-Gatekeeper-SessionToken": session_token,
    }
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
                               org_id: str = "",
                               subscription_key: str = "") -> tuple[list[bytes | None], str]:
    """
    Descarga todas las fotos de un submission.
    Retorna (lista_de_bytes, mensaje_de_error).
    """
    if not photo_filenames:
        return [], ""

    try:
        token = authenticate(email, password, org_id, subscription_key)
    except RuntimeError as e:
        return [None] * len(photo_filenames), str(e)

    results = []
    errors  = []
    for fn in photo_filenames:
        b = get_photo_bytes(fn, token, subscription_key)
        results.append(b)
        if b is None:
            errors.append(fn)

    err_msg = f"{len(errors)} foto(s) no descargadas: {errors[:2]}" if errors else ""
    return results, err_msg
