"""
Integração OneDrive / SharePoint via Microsoft Graph.

Estrutura de pastas por evento (imóveis):
    <BASE>/<Tipo>/<Condomínio>/<Código>
    ex.:  .../Ribeirão Preto/Venda/Alphaville 2/0001

A BASE é a pasta "Ribeirão Preto" (Comum/4. Fotos e Vídeos - Imóveis/Ribeirão Preto),
apontada de preferência por um LINK de compartilhamento (ONEDRIVE_BASE_LINK), que é
resolvido para drive+item via Graph /shares — assim não dependemos de digitar o caminho
longo com acentos. Há fallback por MS_DRIVE_ID + ONEDRIVE_BASE_PATH.

Anti-duplicata de condomínios: o nome digitado é comparado às pastas existentes de forma
"number-aware" — os números precisam bater exatamente e só a parte de texto é comparada por
similaridade. Assim "Alphaville 1" == "alfavile1", mas "Alphaville 1" != "Alphaville 2".

Tudo é best-effort: enquanto não configurado (ou em erro), as funções não derrubam o app —
retornam status "skipped"/None e o calendário segue normal.

Variáveis de ambiente:
    MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET   -> credenciais do app (Entra)
    ONEDRIVE_BASE_LINK                              -> link da pasta "Ribeirão Preto" (recomendado)
    MS_DRIVE_ID  + ONEDRIVE_BASE_PATH               -> alternativa (drive + caminho)
    MS_USER_ID                                      -> alternativa p/ OneDrive de um usuário
"""
import os
import re
import time
import json
import base64
import difflib
import threading
import unicodedata
import urllib.parse
import requests

GRAPH = "https://graph.microsoft.com/v1.0"

TIPOS_IMOVEL = ("Venda", "Locação")
AUTO_THRESHOLD = 0.86   # auto-merge no servidor (quando não veio confirmado da UI)
SUGGEST_MIN    = 0.60   # abaixo disso nem sugere

ROMANOS = {"i": 1, "ii": 2, "iii": 3, "iv": 4, "v": 5,
           "vi": 6, "vii": 7, "viii": 8, "ix": 9, "x": 10}

# Caches em memória
_token_cache       = {"value": None, "exp": 0}
_base_cache        = {"drive_id": None, "item_id": None}
_condo_cache       = {"value": None, "exp": 0}
_backup_read_cache = {"value": None, "exp": 0}
_backup_lock       = threading.Lock()


# ---------------------------------------------------------------------------
# Config / auth
# ---------------------------------------------------------------------------

def _cfg():
    return {
        "tenant":        os.environ.get("MS_TENANT_ID", "").strip(),
        "client_id":     os.environ.get("MS_CLIENT_ID", "").strip(),
        "client_secret": os.environ.get("MS_CLIENT_SECRET", "").strip(),
        "base_link":     os.environ.get("ONEDRIVE_BASE_LINK", "").strip(),
        "drive_id":      os.environ.get("MS_DRIVE_ID", "").strip(),
        "user_id":       os.environ.get("MS_USER_ID", "").strip(),
        "base_path":     os.environ.get("ONEDRIVE_BASE_PATH", "").strip().strip("/"),
    }


def is_configured():
    c = _cfg()
    has_creds  = c["tenant"] and c["client_id"] and c["client_secret"]
    has_target = c["base_link"] or c["drive_id"] or c["user_id"]
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


def _auth_headers():
    return {"Authorization": f"Bearer {_get_token()}", "Content-Type": "application/json"}


# ---------------------------------------------------------------------------
# Resolução da pasta-base (drive_id + item_id)
# ---------------------------------------------------------------------------

def _resolve_base():
    if _base_cache["item_id"]:
        return _base_cache["drive_id"], _base_cache["item_id"]

    c = _cfg()
    headers = _auth_headers()

    if c["base_link"]:
        share_id = "u!" + base64.urlsafe_b64encode(c["base_link"].encode()).decode().rstrip("=")
        r = requests.get(f"{GRAPH}/shares/{share_id}/driveItem", headers=headers, timeout=20)
        r.raise_for_status()
        item = r.json()
        drive_id = item["parentReference"]["driveId"]
        item_id = item["id"]
    else:
        drive_id = c["drive_id"]
        if not drive_id and c["user_id"]:
            r = requests.get(f"{GRAPH}/users/{c['user_id']}/drive", headers=headers, timeout=20)
            r.raise_for_status()
            drive_id = r.json()["id"]
        if c["base_path"]:
            enc = urllib.parse.quote(c["base_path"])
            r = requests.get(f"{GRAPH}/drives/{drive_id}/root:/{enc}", headers=headers, timeout=20)
        else:
            r = requests.get(f"{GRAPH}/drives/{drive_id}/root", headers=headers, timeout=20)
        r.raise_for_status()
        item_id = r.json()["id"]

    _base_cache["drive_id"] = drive_id
    _base_cache["item_id"] = item_id
    return drive_id, item_id


# ---------------------------------------------------------------------------
# Helpers de pasta (Graph, por item-id)
# ---------------------------------------------------------------------------

def _safe(name):
    name = (name or "").strip()
    for ch in '\\/:*?"<>|':
        name = name.replace(ch, "-")
    return name.strip(". ") or "sem-nome"


def _list_child_folders(headers, drive_id, parent_id):
    out = []
    url = f"{GRAPH}/drives/{drive_id}/items/{parent_id}/children?$select=id,name,folder&$top=200"
    while url:
        r = requests.get(url, headers=headers, timeout=20)
        r.raise_for_status()
        j = r.json()
        for it in j.get("value", []):
            if "folder" in it:
                out.append({"id": it["id"], "name": it["name"]})
        url = j.get("@odata.nextLink")
    return out


def _find_child_folder(headers, drive_id, parent_id, name):
    target = _norm(name)
    for f in _list_child_folders(headers, drive_id, parent_id):
        if _norm(f["name"]) == target:
            return f
    return None


def _find_tipo_folder(headers, drive_id, base_id, tipo):
    """
    Localiza a pasta do tipo (Venda/Locação) mesmo quando ela tem prefixo/variação,
    ex.: '2. Venda', '1. Locação'. Prioriza match exato; depois 'tipo' como token.
    """
    alvo = _norm(tipo)  # ex.: "venda"
    filhos = _list_child_folders(headers, drive_id, base_id)
    for f in filhos:                       # 1) nome idêntico
        if _norm(f["name"]) == alvo:
            return f
    for f in filhos:                       # 2) contém o tipo como palavra ("2 venda")
        if alvo in _norm(f["name"]).split():
            return f
    return None


def _ensure_folder(headers, drive_id, parent_id, name):
    """Garante a pasta `name` sob `parent_id`. Retorna (item_id, criada_bool)."""
    name = _safe(name)
    url = f"{GRAPH}/drives/{drive_id}/items/{parent_id}/children"
    body = {"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
    r = requests.post(url, headers=headers, json=body, timeout=20)
    if r.status_code == 201:
        return r.json()["id"], True
    if r.status_code == 409:
        existing = _find_child_folder(headers, drive_id, parent_id, name)
        if existing:
            return existing["id"], False
    r.raise_for_status()
    raise RuntimeError(f"Falha ao criar pasta '{name}': HTTP {r.status_code}")


# ---------------------------------------------------------------------------
# Normalização e matching de condomínios (number-aware)
# ---------------------------------------------------------------------------

def _norm(s):
    s = unicodedata.normalize("NFKD", (s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^a-zA-Z0-9]+", " ", s).lower()
    return re.sub(r"\s+", " ", s).strip()


def _split(s):
    """Retorna (parte_alfabetica, tupla_de_numeros). Números glud. e romanos viram int."""
    n = _norm(s)
    nums = [int(x) for x in re.findall(r"\d+", n)]
    alpha_tokens = []
    for t in re.sub(r"\d+", " ", n).split():
        if t in ROMANOS:
            nums.append(ROMANOS[t])
        else:
            alpha_tokens.append(t)
    return " ".join(alpha_tokens), tuple(sorted(nums))


def best_match(typed, known):
    """
    Retorna (nome_canonico, score, exato) do melhor candidato em `known`.
    Só considera candidatos cujos NÚMEROS batem exatamente; compara só o texto.
    """
    if not typed:
        return (None, 0.0, False)
    tnorm = _norm(typed)
    t_alpha, t_num = _split(typed)
    best, best_score = None, 0.0
    for k in known:
        if _norm(k) == tnorm:
            return (k, 1.0, True)
        k_alpha, k_num = _split(k)
        if k_num != t_num:
            continue
        score = difflib.SequenceMatcher(None, t_alpha, k_alpha).ratio()
        if score > best_score:
            best, best_score = k, score
    return (best, best_score, False)


# ---------------------------------------------------------------------------
# Listagem de condomínios conhecidos (união Venda + Locação)
# ---------------------------------------------------------------------------

def _condominios_uncached(headers, drive_id, base_id):
    nomes = {}
    for tipo in TIPOS_IMOVEL:
        tf = _find_tipo_folder(headers, drive_id, base_id, tipo)
        if not tf:
            continue
        for f in _list_child_folders(headers, drive_id, tf["id"]):
            if not f["name"].startswith("_"):
                nomes[f["name"]] = True
    return sorted(nomes.keys(), key=str.lower)


def listar_condominios(force=False):
    """Lista cacheada (TTL 5 min) p/ o autocomplete. [] se não configurado/erro."""
    if not is_configured():
        return []
    now = time.time()
    if not force and _condo_cache["value"] is not None and _condo_cache["exp"] > now:
        return _condo_cache["value"]
    try:
        headers = _auth_headers()
        drive_id, base_id = _resolve_base()
        nomes = _condominios_uncached(headers, drive_id, base_id)
        _condo_cache.update(value=nomes, exp=now + 300)
        return nomes
    except Exception:
        return _condo_cache["value"] or []


def sugerir_condominio(tipo, nome):
    """Para o endpoint de confirmação pré-envio. Retorna {match, score, exato}."""
    known = listar_condominios()
    name, score, exato = best_match(nome, known)
    return {
        "match": name,
        "score": round(score, 3),
        "exato": exato,
        "sugerir": bool(name and (exato or score >= SUGGEST_MIN)),
    }


# ---------------------------------------------------------------------------
# Criação da pasta do evento
# ---------------------------------------------------------------------------

def create_event_folder(tipo, condominio, codigo, confirmado=False):
    """
    Cria <BASE>/<Tipo>/<Condomínio>/<Código> de forma idempotente.
    Retorna {status, path, condominio_final, merged, message}.
    """
    if not is_configured():
        return {"status": "skipped", "message": "OneDrive não configurado"}
    if tipo not in TIPOS_IMOVEL:
        return {"status": "skipped", "message": "Tipo não gera pasta"}
    if not (condominio and str(condominio).strip()) or not (codigo and str(codigo).strip()):
        return {"status": "skipped", "message": "Faltam condomínio/código"}

    try:
        headers = _auth_headers()
        drive_id, base_id = _resolve_base()

        nome = str(condominio).strip()
        # Auto-merge no servidor só quando a escolha NÃO veio confirmada da UI
        if not confirmado:
            known = _condominios_uncached(headers, drive_id, base_id)
            cand, score, exato = best_match(nome, known)
            if cand and (exato or score >= AUTO_THRESHOLD):
                nome = cand

        # Usa a pasta de tipo existente (ex.: "2. Venda"); só cria se não houver
        tf = _find_tipo_folder(headers, drive_id, base_id, tipo)
        if tf:
            tipo_id, tipo_nome = tf["id"], tf["name"]
        else:
            tipo_id, _ = _ensure_folder(headers, drive_id, base_id, tipo)
            tipo_nome = tipo

        cond_id, cond_nova = _ensure_folder(headers, drive_id, tipo_id, nome)
        _ensure_folder(headers, drive_id, cond_id, str(codigo).strip())

        _condo_cache["exp"] = 0  # invalida cache do autocomplete
        return {
            "status": "ok",
            "path": f"{tipo_nome}/{nome}/{str(codigo).strip()}",
            "condominio_final": nome,
            "merged": not cond_nova,   # True = reaproveitou condomínio existente
            "message": "Pasta criada/garantida",
        }
    except Exception as e:
        return {"status": "error", "message": str(e)}


# ---------------------------------------------------------------------------
# Backup do eventos.json no OneDrive (redundância contra perda em deploy)
# ---------------------------------------------------------------------------

def salvar_backup_eventos(events):
    """Best-effort, em thread daemon — não bloqueia a resposta nem derruba nada."""
    if not is_configured():
        return
    payload = json.dumps(events, ensure_ascii=False, indent=2)
    threading.Thread(target=_backup_worker, args=(payload,), daemon=True).start()


def _backup_worker(content):
    try:
        with _backup_lock:
            headers = _auth_headers()
            drive_id, base_id = _resolve_base()
            folder_id, _ = _ensure_folder(headers, drive_id, base_id, "_backup")
            url = f"{GRAPH}/drives/{drive_id}/items/{folder_id}:/eventos.json:/content"
            up = {"Authorization": headers["Authorization"], "Content-Type": "application/json"}
            requests.put(url, headers=up, data=content.encode("utf-8"), timeout=30)
            _backup_read_cache["exp"] = 0
    except Exception:
        pass


def ler_backup_eventos():
    """Lê o backup do OneDrive. None se não há / não configurado. Cacheado p/ não martelar o Graph."""
    if not is_configured():
        return None
    now = time.time()
    if _backup_read_cache["exp"] > now:
        return _backup_read_cache["value"]
    try:
        headers = _auth_headers()
        drive_id, base_id = _resolve_base()
        folder = _find_child_folder(headers, drive_id, base_id, "_backup")
        if not folder:
            _backup_read_cache.update(value=None, exp=now + 300)
            return None
        url = f"{GRAPH}/drives/{drive_id}/items/{folder['id']}:/eventos.json:/content"
        r = requests.get(url, headers={"Authorization": headers["Authorization"]}, timeout=30)
        if r.status_code == 200:
            data = r.json()
            _backup_read_cache.update(value=data, exp=now + 60)
            return data
    except Exception:
        pass
    _backup_read_cache.update(value=None, exp=now + 300)
    return None
