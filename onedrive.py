"""
Integração OneDrive / SharePoint via Microsoft Graph.

Cria a estrutura de pastas para um evento:  <BASE>/<tipo>/<edificio>/<codigo>
Ex.:  Documentos/Fotos/Locação/Edifício X/AP1234

Tudo é controlado por variáveis de ambiente. Enquanto as credenciais e o
caminho-base não forem configurados, as funções fazem "no-op" de forma segura
(retornam status "skipped") — assim o calendário continua funcionando normalmente
e a integração é ativada só preenchendo as variáveis abaixo.

Variáveis de ambiente esperadas (preencher quando formos incrementar):
    MS_TENANT_ID        -> ID do tenant (Azure AD / Entra)
    MS_CLIENT_ID        -> App registration (client id)
    MS_CLIENT_SECRET    -> Segredo do app
    MS_DRIVE_ID         -> ID do drive (OneDrive do usuário ou biblioteca do SharePoint)
                           Alternativa: MS_USER_ID (usa /users/{id}/drive)
    ONEDRIVE_BASE_PATH  -> Caminho-base dentro do drive (ex.: "Fotos" ou "Documents/Fotos")
                           Vazio = raiz do drive.
"""
import os
import time
import urllib.parse
import requests

GRAPH = "https://graph.microsoft.com/v1.0"

# Cache simples de token (client-credentials, válido ~1h)
_token_cache = {"value": None, "exp": 0}


def _cfg():
    return {
        "tenant": os.environ.get("MS_TENANT_ID", "").strip(),
        "client_id": os.environ.get("MS_CLIENT_ID", "").strip(),
        "client_secret": os.environ.get("MS_CLIENT_SECRET", "").strip(),
        "drive_id": os.environ.get("MS_DRIVE_ID", "").strip(),
        "user_id": os.environ.get("MS_USER_ID", "").strip(),
        "base_path": os.environ.get("ONEDRIVE_BASE_PATH", "").strip().strip("/"),
    }


def is_configured():
    c = _cfg()
    has_creds = c["tenant"] and c["client_id"] and c["client_secret"]
    has_target = c["drive_id"] or c["user_id"]
    return bool(has_creds and has_target)


def _get_token():
    c = _cfg()
    now = time.time()
    if _token_cache["value"] and _token_cache["exp"] > now + 60:
        return _token_cache["value"]

    url = f"https://login.microsoftonline.com/{c['tenant']}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": c["client_id"],
        "client_secret": c["client_secret"],
        "scope": "https://graph.microsoft.com/.default",
    }
    r = requests.post(url, data=data, timeout=20)
    r.raise_for_status()
    tok = r.json()
    _token_cache["value"] = tok["access_token"]
    _token_cache["exp"] = now + int(tok.get("expires_in", 3600))
    return _token_cache["value"]


def _drive_base():
    """Retorna o prefixo de URL do drive para o Graph."""
    c = _cfg()
    if c["drive_id"]:
        return f"{GRAPH}/drives/{c['drive_id']}"
    return f"{GRAPH}/users/{c['user_id']}/drive"


def _safe(name):
    """Remove caracteres inválidos para nome de pasta no OneDrive/SharePoint."""
    name = (name or "").strip()
    for ch in '\\/:*?"<>|':
        name = name.replace(ch, "-")
    return name.strip(". ") or "sem-nome"


def _ensure_child(headers, parent_path, name):
    """
    Garante que a pasta `name` existe dentro de `parent_path` (caminho relativo
    ao root do drive). Retorna o novo caminho relativo. Idempotente.
    """
    base = _drive_base()
    name = _safe(name)

    # Endereçamento por caminho. parent_path vazio = root.
    if parent_path:
        parent_enc = urllib.parse.quote(parent_path)
        list_url = f"{base}/root:/{parent_enc}:/children"
    else:
        list_url = f"{base}/root/children"

    body = {
        "name": name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "fail",  # não duplica se já existe
    }
    r = requests.post(list_url, headers=headers, json=body, timeout=20)

    # 201 = criada; 409 = já existe (ok); outros = erro
    if r.status_code not in (201, 409):
        r.raise_for_status()

    return f"{parent_path}/{name}" if parent_path else name


def create_event_folder(tipo, edificio, codigo):
    """
    Cria <BASE>/<tipo>/<edificio>/<codigo> de forma idempotente.

    Retorna dict com:
        status: "ok" | "skipped" | "error"
        path:   caminho criado (quando ok)
        message: detalhe
    """
    if not is_configured():
        return {"status": "skipped", "message": "OneDrive não configurado"}

    # Sem código não há pasta-alvo (o nome da pasta é o código do imoview)
    if not (codigo and str(codigo).strip()):
        return {"status": "skipped", "message": "Sem código Imoview"}

    try:
        c = _cfg()
        headers = {
            "Authorization": f"Bearer {_get_token()}",
            "Content-Type": "application/json",
        }

        # Monta a cadeia de segmentos, ignorando os vazios
        segments = [s for s in [c["base_path"], tipo, edificio, str(codigo)] if s and str(s).strip()]

        path = ""
        # base_path pode ter várias partes (ex.: "Documents/Fotos") — trata cada nível
        flat = []
        for seg in segments:
            flat.extend([p for p in str(seg).split("/") if p.strip()])

        for seg in flat:
            path = _ensure_child(headers, path, seg)

        return {"status": "ok", "path": path, "message": "Pasta criada/garantida"}

    except Exception as e:  # nunca derruba a criação do evento
        return {"status": "error", "message": str(e)}
