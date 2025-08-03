import os, msal, requests
from flask import Flask, render_template, redirect, request, session, url_for
from flask_session import Session
from werkzeug.middleware.proxy_fix import ProxyFix
import app_config

app = Flask(__name__, template_folder="templates")
app.config.from_object(app_config)
Session(app)

app.secret_key = os.urandom(24)
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)

@app.route("/")
def index():
    if not session.get("user"):
        return redirect(url_for("login"))
    return render_template("index.html", user=session["user"])

@app.route("/login")
def login():
    session["flow"] = _build_auth_code_flow(scopes=app_config.SCOPE)
    return redirect(session["flow"]["auth_uri"])

@app.route(app_config.REDIRECT_PATH)
def authorized():
    try:
        cache = _load_cache()
        result = _build_msal_app(cache=cache).acquire_token_by_auth_code_flow(
            session.get("flow", {}), request.args)
        if "error" in result:
            return render_template("auth_error.html", result=result)
        session["user"] = result.get("id_token_claims")
        _save_cache(cache)
    except ValueError:
        pass
    return redirect(url_for("index"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(
        app_config.AUTHORITY + "/oauth2/v2.0/logout" +
        "?post_logout_redirect_uri=" + url_for("index", _external=True)
    )

@app.route("/graphcall")
def graphcall():
    token = _get_token_from_cache(app_config.SCOPE)
    if not token:
        return redirect(url_for("login"))
    headers = {'Authorization': 'Bearer ' + token['access_token']}
    response = requests.get(app_config.ENDPOINT, headers=headers)

    if response.status_code == 200:
        chats = response.json().get("value", [])
        return render_template("graph_result.html", chats=chats)
    else:
        return f"Graph API call failed: {response.status_code}<br>{response.text}"

@app.route("/chat/<chat_id>")
def chat_messages(chat_id):
    token = _get_token_from_cache(app_config.SCOPE)
    if not token:
        return redirect(url_for("login"))

    headers = {'Authorization': 'Bearer ' + token['access_token']}
    messages_url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"
    response = requests.get(messages_url, headers=headers)

    if response.status_code == 200:
        messages = response.json().get("value", [])
        return render_template("chat_messages.html", messages=messages, chat_id=chat_id)
    else:
        return f"Failed to get messages: {response.status_code}<br>{response.text}"

def _load_cache():
    cache = msal.SerializableTokenCache()
    if session.get("token_cache"):
        cache.deserialize(session["token_cache"])
    return cache

def _save_cache(cache):
    if cache.has_state_changed:
        session["token_cache"] = cache.serialize()

def _build_msal_app(cache=None, authority=None):
    return msal.ConfidentialClientApplication(
        app_config.CLIENT_ID,
        authority=authority or app_config.AUTHORITY,
        client_credential=app_config.CLIENT_SECRET,
        token_cache=cache
    )

def _build_auth_code_flow(authority=None, scopes=None):
    return _build_msal_app(authority=authority).initiate_auth_code_flow(
        scopes or [],
        redirect_uri=url_for("authorized", _external=True)
    )

def _get_token_from_cache(scope=None):
    cache = _load_cache()
    cca = _build_msal_app(cache=cache)
    accounts = cca.get_accounts()
    if accounts:
        result = cca.acquire_token_silent(scope, account=accounts[0])
        _save_cache(cache)
        return result

@app.route("/download_chat/<chat_id>")
def download_chat(chat_id):
    token = _get_token_from_cache(app_config.SCOPE)
    if not token:
        return redirect(url_for("login"))

    headers = {'Authorization': 'Bearer ' + token['access_token']}
    url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages"

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        messages_json = response.json()
        filename = f"chat_{chat_id.replace(':', '_')}.json"

        # Save JSON to local file
        with open(filename, "w", encoding="utf-8") as f:
            import json
            json.dump(messages_json, f, indent=2)

        return f"✅ Messages saved to: <code>{filename}</code>"
    else:
        return f"❌ Failed to get messages: {response.status_code}<br>{response.text}"


if __name__ == "__main__":
    app.run(debug=True, host="localhost", port=5000)
