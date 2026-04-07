# Vellum — Web app deployment guide

This document covers deploying the Vellum **web app** to a Linux server
behind nginx + Let's Encrypt TLS, using the Docker image built by
`docker/Dockerfile`. The CLI and Tkinter desktop GUI don't need any
deployment — they run wherever you can `pip install -r requirements.txt`.

## Quick path: single container, no proxy

For local testing or trusted-LAN deployment:

```bash
git clone https://github.com/om0wt/vellum.git
cd vellum
docker compose -f docker/docker-compose.yml up --build -d
```

Open <http://YOUR_HOST:5001>. Done.

This binds the container to host port `5001` directly. No TLS, no auth,
no IP-trust headers — fine for `localhost` or a trusted network.

For anything reachable from the internet, use the production setup
below.

---

## Production: container behind nginx with TLS

The recommended layout:

```
Internet
    │  https://YOUR_HOST:8443/vellum/
    ▼
┌───────────────────────────────┐
│ nginx (TLS termination,      │
│ rate limiting, large body)   │
│  /vellum/ → 127.0.0.1:5001   │
└──────────────┬────────────────┘
               │ HTTP
               ▼
┌───────────────────────────────┐
│ Vellum container (waitress)  │
│  127.0.0.1:5001              │
│  TRUST_PROXY=1                │
│  ./logs/access.log mounted    │
└───────────────────────────────┘
```

### 1. Prerequisites on the server

* Docker engine + `docker compose` v2 plugin
* nginx ≥ 1.18
* A TLS certificate for your hostname. The example assumes Let's
  Encrypt under `/etc/letsencrypt/live/YOUR_HOST/` — if you use a
  different CA, update the `ssl_certificate` paths in the nginx config.
* A user that can manage the container (typically the `docker` group)

### 2. Clone the repo

```bash
sudo mkdir -p /srv && cd /srv
sudo git clone https://github.com/om0wt/vellum.git
sudo chown -R $(id -u):$(id -g) vellum
cd vellum
```

### 3. Enable `TRUST_PROXY` in the compose file

By default `docker/docker-compose.yml` has the `TRUST_PROXY` environment
variable commented out (it's only meaningful behind a reverse proxy).
Uncomment it for the production deployment:

```yaml
    environment:
      PORT: "5001"
      TRUST_PROXY: "1"          # ← uncomment this line
```

This makes Flask honor the `X-Forwarded-For`, `-Proto`, `-Host`, and
`-Prefix` headers nginx sends. Without it, the access log records
nginx's IP (always `127.0.0.1`) instead of the real client, and the
form's `url_for('convert')` won't include the `/vellum/` prefix.

### 4. Start the container

```bash
docker compose -f docker/docker-compose.yml up --build -d
```

Verify it's running and listening on `127.0.0.1:5001`:

```bash
docker compose -f docker/docker-compose.yml ps
docker compose -f docker/docker-compose.yml logs --tail 20
ss -tlnp | grep 5001
```

You should see something like `[INFO] starting waitress on 0.0.0.0:5001`
in the logs and a listener on `0.0.0.0:5001`. The compose port mapping
publishes it as `127.0.0.1:5001` on the host — local-only, not
internet-reachable directly.

### 5. Install the nginx config

A ready-to-use nginx site config is in this repo at
[`deploy/nginx/vellum.conf`](deploy/nginx/vellum.conf). Copy it into
nginx's site-enabled directory and substitute your hostname:

```bash
sudo cp deploy/nginx/vellum.conf /etc/nginx/sites-enabled/vellum.conf
sudo sed -i 's/YOUR_HOST/far-far-away.mooo.com/g' /etc/nginx/sites-enabled/vellum.conf
sudo nginx -t
sudo systemctl reload nginx
```

(Replace `far-far-away.mooo.com` with whatever your real hostname is.)

The config:

* Terminates TLS on port `8443` (override if you want `443`).
* Mounts Vellum at `/vellum/` and redirects `/` → `/vellum/`.
* Sets `X-Forwarded-Proto`, `X-Forwarded-For`, `X-Forwarded-Host`,
  `X-Forwarded-Prefix /vellum` so the Flask app generates correct URLs.
* Sets `client_max_body_size 60M` (Vellum's app limit is 50 MB; 60M
  leaves headroom for multipart boundaries).
* Sets `proxy_read_timeout 300s` so long OCR conversions don't 504.
* Adds an HSTS header at the TLS layer in addition to the security
  headers Vellum sets on every response.

### 6. Verify end-to-end

```bash
curl -kI https://YOUR_HOST:8443/vellum/
```

Expected: `HTTP/2 200`, with `Server: pdf2docx-web`, `X-App-Version: 1.0.0`,
`Content-Security-Policy: ...`, etc.

Then open <https://YOUR_HOST:8443/vellum/> in a browser and convert a
test PDF.

Tail the access log on the host to confirm the IP and request flow:

```bash
tail -f logs/access.log
```

Expected line for a successful conversion:

```
2026-04-08 09:31:15 ip=203.0.113.42 START file='input.pdf' ocr=False ocr_lang=- no_stream=True ua='Mozilla/5.0 ...'
2026-04-08 09:31:17 ip=203.0.113.42 OK    file='input.pdf' in=343138B out=94770B
```

Note the **real client IP**, not `127.0.0.1` — that confirms
`TRUST_PROXY=1` is wired correctly.

---

## Updating to a new release

```bash
cd /srv/vellum
git pull
docker compose -f docker/docker-compose.yml up --build -d
```

The container restarts with the new image. nginx doesn't need touching
(no config change). The access log on the host (`./logs/access.log`)
persists across rebuilds because it's a bind mount.

---

## Tunable settings

All set via the `environment:` block in `docker/docker-compose.yml`:

| Var | Default | What it does |
|---|---|---|
| `PORT` | `5001` (compose) | Server listen port inside the container |
| `TRUST_PROXY` | unset | Set to `1` behind a reverse proxy so `X-Forwarded-*` is honored |
| `ACCESS_LOG_FILE` | `logs/access.log` | Path to the access log file |
| `RATE_LIMIT_REQUESTS` | `30` | Max requests per window per IP |
| `RATE_LIMIT_WINDOW` | `60` | Window length in seconds |

To change one, edit `docker/docker-compose.yml` then re-run
`docker compose ... up -d` (no `--build` needed for pure env changes).

---

## Operations cheatsheet

```bash
# Status / logs
docker compose -f docker/docker-compose.yml ps
docker compose -f docker/docker-compose.yml logs -f          # general app logs
tail -f logs/access.log                                       # request log

# Restart / stop
docker compose -f docker/docker-compose.yml restart
docker compose -f docker/docker-compose.yml down

# Rebuild after pulling new code
git pull && docker compose -f docker/docker-compose.yml up --build -d

# nginx
sudo nginx -t && sudo systemctl reload nginx
sudo tail -f /var/log/nginx/vellum.access.log /var/log/nginx/vellum.error.log
```

---

## Security notes

The Flask app's built-in posture is documented in [README.md](README.md#security-posture-web-app).
Things to check on the production deployment specifically:

* **Container is not directly exposed.** The host port mapping is
  `5001:5001` but nginx is the only thing meant to talk to it. If your
  server has a public IP, make sure your firewall blocks port 5001 from
  the outside (the default Docker bridge does **not** firewall the
  published port — it's open by default).
  ```bash
  sudo ufw allow 8443/tcp
  sudo ufw deny 5001/tcp     # belt + suspenders
  ```
* **No authentication is built in.** If the converter shouldn't be
  open to the public internet, add HTTP basic auth at the nginx layer:
  ```nginx
  location /vellum/ {
      auth_basic           "Vellum";
      auth_basic_user_file /etc/nginx/.htpasswd;
      proxy_pass http://vellum_app/;
      ...
  }
  ```
  Generate the password file with `htpasswd -c /etc/nginx/.htpasswd youruser`.
* **PDF parser CVE history.** PyMuPDF and pdf2docx have had memory-
  corruption CVEs. If exposing the converter to untrusted internet,
  consider running the container with reduced capabilities and a
  read-only root filesystem:
  ```yaml
      read_only: true
      cap_drop: [ALL]
      tmpfs:
        - /tmp
        - /app/logs
        - /app/build
  ```
  (Note: with `read_only: true` you'll need to keep the access log on
  a tmpfs or volume since the app needs to write to it.)
* **Rate limiting** is on by default — 30 requests per 60 seconds per
  IP — but if you're behind Cloudflare or similar, the per-IP key may
  collapse onto a few CDN egress IPs unless you also pass the real
  client IP via the appropriate header (`CF-Connecting-IP` for
  Cloudflare). The current Vellum code uses `request.remote_addr` which
  is what `ProxyFix` resolves to — that's the right IP for nginx-style
  reverse proxies but won't pick up Cloudflare-only headers
  automatically.
