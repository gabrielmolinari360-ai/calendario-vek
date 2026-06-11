"""
Calendario VEK - Veguel Imoveis
Calendario compartilhado com sync em tempo real via WebSocket
"""
import os, json, uuid
from datetime import datetime
from flask import Flask, render_template, request, jsonify, session, redirect
from flask_socketio import SocketIO, emit

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "calendario-vek-2026")
socketio = SocketIO(app, cors_allowed_origins="*", async_mode="threading")

APP_PASS   = os.environ.get("APP_PASSWORD", "veguel2026")
DATA_FILE  = os.path.join(os.path.dirname(__file__), "data", "eventos.json")

# ---------------------------------------------------------------------------
# Storage helpers
# ---------------------------------------------------------------------------

def load_events():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, encoding="utf-8") as f:
            return json.load(f)
    return []

def save_events(events):
    os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(events, f, ensure_ascii=False, indent=2)

# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        if request.form.get("senha") == APP_PASS:
            session["ok"] = True
            session["nome"] = request.form.get("nome", "Usuário").strip() or "Usuário"
            return redirect("/")
        error = "Senha incorreta."
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

def require_login():
    if not session.get("ok"):
        return redirect("/login")
    return None

# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    redir = require_login()
    if redir: return redir
    return render_template("index.html", nome=session.get("nome", ""))

# ---------------------------------------------------------------------------
# REST API
# ---------------------------------------------------------------------------

@app.route("/api/eventos", methods=["GET"])
def api_get_eventos():
    redir = require_login()
    if redir: return jsonify({"error": "unauthorized"}), 401
    return jsonify(load_events())

@app.route("/api/eventos", methods=["POST"])
def api_criar_evento():
    redir = require_login()
    if redir: return jsonify({"error": "unauthorized"}), 401

    data = request.json or {}
    evento = {
        "id":         data.get("id") or str(uuid.uuid4()),
        "title":      data.get("title", "Sem título"),
        "start":      data.get("start", ""),
        "end":        data.get("end", ""),
        "allDay":     data.get("allDay", False),
        "color":      data.get("color", "#3B6E3A"),
        "descricao":  data.get("descricao", ""),
        "tipo":       data.get("tipo", ""),
        "criado_por": session.get("nome", ""),
        "criado_em":  datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    events = load_events()
    events.append(evento)
    save_events(events)

    socketio.emit("evento_criado", evento)
    return jsonify(evento), 201

@app.route("/api/eventos/<evento_id>", methods=["PUT"])
def api_editar_evento(evento_id):
    redir = require_login()
    if redir: return jsonify({"error": "unauthorized"}), 401

    data = request.json or {}
    events = load_events()
    for i, ev in enumerate(events):
        if ev["id"] == evento_id:
            ev.update({
                "title":     data.get("title", ev["title"]),
                "start":     data.get("start", ev["start"]),
                "end":       data.get("end", ev["end"]),
                "allDay":    data.get("allDay", ev.get("allDay", False)),
                "color":     data.get("color", ev.get("color", "#3B6E3A")),
                "descricao": data.get("descricao", ev.get("descricao", "")),
                "tipo":      data.get("tipo", ev.get("tipo", "")),
                "editado_por": session.get("nome", ""),
                "editado_em":  datetime.now().strftime("%d/%m/%Y %H:%M"),
            })
            events[i] = ev
            save_events(events)
            socketio.emit("evento_editado", ev)
            return jsonify(ev)
    return jsonify({"error": "não encontrado"}), 404

@app.route("/api/eventos/<evento_id>", methods=["DELETE"])
def api_deletar_evento(evento_id):
    redir = require_login()
    if redir: return jsonify({"error": "unauthorized"}), 401

    events = load_events()
    events = [e for e in events if e["id"] != evento_id]
    save_events(events)
    socketio.emit("evento_deletado", {"id": evento_id})
    return jsonify({"ok": True})

# ---------------------------------------------------------------------------
# SocketIO
# ---------------------------------------------------------------------------

@socketio.on("connect")
def on_connect():
    pass  # cliente conectado

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    socketio.run(app, host="0.0.0.0", port=port, debug=False, allow_unsafe_werkzeug=True)
