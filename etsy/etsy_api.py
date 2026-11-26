#!/usr/bin/env python3
"""GYMZUP Etsy helper — Etsy Open API v3, Python stdlib only.

Usage:
  export ETSY_KEYSTRING=...  ETSY_SHARED_SECRET=...
  python3 etsy_api.py auth                 # one-time OAuth login (opens browser)
  python3 etsy_api.py me                   # show user_id / shop_id
  python3 etsy_api.py taxonomy <word>      # find taxonomy_id for a category
  python3 etsy_api.py create <listing.json> # create DRAFT listing + images + files
  python3 etsy_api.py listings [state]     # list your listings (draft|active|inactive)
  python3 etsy_api.py orders [limit]       # recent receipts
"""
import base64
import hashlib
import http.server
import json
import mimetypes
import os
import secrets
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
import webbrowser
from pathlib import Path

API = "https://openapi.etsy.com/v3/application"
AUTH_URL = "https://www.etsy.com/oauth/connect"
TOKEN_URL = "https://api.etsy.com/v3/public/oauth/token"
SCOPES = "listings_r listings_w listings_d shops_r shops_w transactions_r profile_r email_r"
REDIRECT_URI = os.environ.get("ETSY_REDIRECT_URI", "http://localhost:3003/callback")
TOKEN_FILE = Path(__file__).resolve().parent / ".etsy_tokens.json"


def die(msg):
    print(f"Error: {msg}", file=sys.stderr)
    sys.exit(1)


def keystring():
    k = os.environ.get("ETSY_KEYSTRING")
    if not k:
        die("set ETSY_KEYSTRING (and ETSY_SHARED_SECRET) first")
    return k


def api_key_header():
    k, s = keystring(), os.environ.get("ETSY_SHARED_SECRET", "")
    if os.environ.get("ETSY_KEY_FORMAT") == "keystring_only" or not s:
        return k
    return f"{k}:{s}"


# ---------- tokens ----------

def load_tokens():
    if not TOKEN_FILE.exists():
        die("not logged in — run: python3 etsy_api.py auth")
    return json.loads(TOKEN_FILE.read_text())


def save_tokens(tok):
    tok["expires_at"] = int(time.time()) + int(tok.get("expires_in", 3600)) - 60
    TOKEN_FILE.write_text(json.dumps(tok, indent=2))
    try:
        os.chmod(TOKEN_FILE, 0o600)
    except OSError:
        pass
    return tok


def token_request(data):
    body = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(TOKEN_URL, data=body, method="POST",
                                 headers={"Content-Type": "application/x-www-form-urlencoded"})
    try:
        with urllib.request.urlopen(req) as r:
            return json.loads(r.read())
    except urllib.error.HTTPError as e:
        die(f"token request failed {e.code}: {e.read().decode(errors='replace')}")


def access_token():
    tok = load_tokens()
    if time.time() >= tok.get("expires_at", 0):
        tok = save_tokens(token_request({
            "grant_type": "refresh_token",
            "client_id": keystring(),
            "refresh_token": tok["refresh_token"],
        }))
    return tok["access_token"]


# ---------- HTTP ----------

def call(method, path, *, params=None, json_body=None, form=None, files=None):
    url = API + path
    if params:
        url += "?" + urllib.parse.urlencode(params, doseq=True)
    headers = {"x-api-key": api_key_header(), "Authorization": f"Bearer {access_token()}"}
    data = None
    if json_body is not None:
        data = json.dumps(json_body).encode()
        headers["Content-Type"] = "application/json"
    elif files:
        data, ctype = multipart(form or {}, files)
        headers["Content-Type"] = ctype
    elif form:
        data = urllib.parse.urlencode(form, doseq=True).encode()
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req) as r:
            raw = r.read()
            return json.loads(raw) if raw else {}
    except urllib.error.HTTPError as e:
        die(f"{method} {path} -> {e.code}: {e.read().decode(errors='replace')}")


def multipart(fields, files):
    boundary = uuid.uuid4().hex
    out = bytearray()
    for k, v in fields.items():
        out += (f"--{boundary}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n{v}\r\n").encode()
    for field, path in files.items():
        p = Path(path)
        ctype = mimetypes.guess_type(p.name)[0] or "application/octet-stream"
        out += (f"--{boundary}\r\nContent-Disposition: form-data; name=\"{field}\"; "
                f"filename=\"{p.name}\"\r\nContent-Type: {ctype}\r\n\r\n").encode()
        out += p.read_bytes() + b"\r\n"
    out += f"--{boundary}--\r\n".encode()
    return bytes(out), f"multipart/form-data; boundary={boundary}"


def get_shop_id():
    me = call("GET", "/users/me")
    if not me.get("shop_id"):
        die("this Etsy account has no shop")
    return me["shop_id"]


# ---------- commands ----------

def cmd_auth():
    verifier = base64.urlsafe_b64encode(secrets.token_bytes(32)).rstrip(b"=").decode()
    challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
    state = secrets.token_urlsafe(16)
    url = AUTH_URL + "?" + urllib.parse.urlencode({
        "response_type": "code", "client_id": keystring(), "redirect_uri": REDIRECT_URI,
        "scope": SCOPES, "state": state, "code_challenge": challenge,
        "code_challenge_method": "S256",
    })
    result = {}

    class Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            result.update({k: v[0] for k, v in q.items()})
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write("<h2>GYMZUP: login OK. You can close this tab.</h2>".encode())

        def log_message(self, *a):
            pass

    parsed = urllib.parse.urlparse(REDIRECT_URI)
    server = http.server.HTTPServer((parsed.hostname, parsed.port or 80), Handler)
    print("Opening browser for Etsy login...\nIf it does not open, visit:\n" + url)
    webbrowser.open(url)
    while "code" not in result and "error" not in result:
        server.handle_request()
    if "error" in result:
        die(f"Etsy refused: {result}")
    if result.get("state") != state:
        die("state mismatch — try again")
    save_tokens(token_request({
        "grant_type": "authorization_code", "client_id": keystring(),
        "redirect_uri": REDIRECT_URI, "code": result["code"], "code_verifier": verifier,
    }))
    print(f"Logged in. Tokens saved to {TOKEN_FILE}")
    cmd_me()


def cmd_me():
    me = call("GET", "/users/me")
    print(json.dumps(me, indent=2))
    if me.get("shop_id"):
        shop = call("GET", f"/shops/{me['shop_id']}")
        print(f"Shop: {shop.get('shop_name')}  ({shop.get('url')})")


def cmd_taxonomy(word):
    nodes = call("GET", "/seller-taxonomy/nodes")["results"]
    word = word.lower()

    def walk(ns, trail):
        for n in ns:
            t = trail + [n["name"]]
            if word in n["name"].lower():
                print(f"{n['id']:>6}  {' > '.join(t)}")
            walk(n.get("children", []), t)

    walk(nodes, [])


def cmd_create(listing_path):
    lp = Path(listing_path).resolve()
    spec = json.loads(lp.read_text())
    base = lp.parent
    if not spec.get("taxonomy_id"):
        die("set taxonomy_id in listing.json (find it with: etsy_api.py taxonomy planner)")
    images = [base / p for p in spec.get("images", [])]
    files = [base / p for p in spec.get("files", [])]
    missing = [str(p) for p in images + files if not p.exists()]
    if missing:
        die("missing files:\n  " + "\n  ".join(missing))
    tags = spec.get("tags", [])
    if len(tags) > 13 or any(len(t) > 20 for t in tags):
        die("Etsy allows max 13 tags, each max 20 characters")
    if len(spec["title"]) > 140:
        die("title is longer than 140 characters")

    shop_id = get_shop_id()
    body = {k: spec[k] for k in ("title", "description", "price", "quantity", "who_made",
                                  "when_made", "taxonomy_id", "type", "is_supply", "tags")
            if k in spec}
    listing = call("POST", f"/shops/{shop_id}/listings", json_body=body)
    lid = listing["listing_id"]
    print(f"Draft listing created: {lid}")

    for i, img in enumerate(images[:10], start=1):
        call("POST", f"/shops/{shop_id}/listings/{lid}/images",
             form={"rank": i}, files={"image": img})
        print(f"  image {i}/{len(images)} uploaded: {img.name}")
    for i, f in enumerate(files[:5], start=1):
        call("POST", f"/shops/{shop_id}/listings/{lid}/files",
             form={"name": f.name, "rank": i}, files={"file": f})
        print(f"  file uploaded: {f.name}")
    print(f"Done. Review and publish: https://www.etsy.com/your/shops/me/listing-editor/edit/{lid}")


def cmd_listings(state="active"):
    shop_id = get_shop_id()
    res = call("GET", f"/shops/{shop_id}/listings", params={"state": state, "limit": 100})
    for l in res.get("results", []):
        price = l.get("price", {})
        amount = price.get("amount", 0) / max(price.get("divisor", 1), 1)
        print(f"{l['listing_id']}  {l['state']:<8} {amount:>7.2f} {price.get('currency_code', '')}  "
              f"views={l.get('views', '-')}  {l['title'][:70]}")
    print(f"{res.get('count', 0)} listing(s) [{state}]")


def cmd_orders(limit="25"):
    shop_id = get_shop_id()
    res = call("GET", f"/shops/{shop_id}/receipts", params={"limit": int(limit)})
    for r in res.get("results", []):
        total = r.get("grandtotal", {})
        amount = total.get("amount", 0) / max(total.get("divisor", 1), 1)
        when = time.strftime("%Y-%m-%d", time.localtime(r.get("create_timestamp", 0)))
        items = ", ".join(t.get("title", "")[:40] for t in r.get("transactions", []))
        print(f"{when}  #{r['receipt_id']}  {amount:>7.2f} {total.get('currency_code', '')}  {items}")
    print(f"{res.get('count', 0)} order(s) total")


def main():
    cmds = {"auth": cmd_auth, "me": cmd_me, "taxonomy": cmd_taxonomy, "create": cmd_create,
            "listings": cmd_listings, "orders": cmd_orders}
    if len(sys.argv) < 2 or sys.argv[1] not in cmds:
        print(__doc__)
        sys.exit(1)
    fn, args = cmds[sys.argv[1]], sys.argv[2:]
    co = fn.__code__
    required = co.co_argcount - len(fn.__defaults__ or ())
    if not required <= len(args) <= co.co_argcount:
        print(__doc__)
        sys.exit(1)
    fn(*args)


if __name__ == "__main__":
    main()
