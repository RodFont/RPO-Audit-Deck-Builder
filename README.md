# RPO Report Audit — Web Service

Turns a checklist-style audit spreadsheet (`.xlsx`/`.xlsm`/`.csv`) into an
executive summary PowerPoint deck + JSON, via a web upload form, instead of
running `audit_to_deck.py` from the command line.

```
User → Akamai (your domain, CDN/WAF) → Linode origin (nginx → FastAPI app) → audit_to_deck.py
                                                                     ↑
                                                        you (developer) push to `main`
                                                        → GitHub Actions builds image,
                                                          pushes to GHCR, deploys over SSH
```

## Repo layout

| Path | Purpose |
|---|---|
| `audit_to_deck.py` | Original CLI script — **unmodified**. All parsing/deck logic lives here. |
| `app/main.py` | FastAPI wrapper: `POST /api/audit` (upload) + `GET /` (upload form) + `GET /healthz`. Imports `audit_to_deck` directly, so edits to the script are picked up automatically. |
| `app/static/index.html` | Minimal upload form. |
| `Dockerfile` | Builds the app image. |
| `docker-compose.yml` | **Local dev**: builds from source, runs app + nginx. |
| `docker-compose.prod.yml` | **Linode/production**: pulls the pre-built image from GHCR (no build step on the server). |
| `nginx/nginx.conf` | Reverse proxy + TLS termination for the Akamai → origin hop. |
| `.github/workflows/deploy.yml` | CI/CD: build & push image on every push to `main`, then SSH-deploy to Linode. |

## Local development

```bash
docker compose up --build
curl http://localhost:8000/healthz
open http://localhost:8000          # upload form
```

Or without Docker:
```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r app/requirements.txt
uvicorn app.main:app --reload --port 8000
```

## Making changes to the script

Because `app/main.py` imports `audit_to_deck` as a module (not a subprocess),
any edit to `audit_to_deck.py` — new column synonyms, new status buckets, deck
layout tweaks, etc. — takes effect the next time the container is rebuilt.
Normal workflow:

1. Edit `audit_to_deck.py` (or the app) on a branch.
2. Test locally: `docker compose up --build` and upload a sample report.
3. Open a PR; merge to `main`.
4. GitHub Actions automatically builds a new image and redeploys it to the
   Linode instance — no manual server access needed for routine changes.

## API

- `GET /healthz` → `{"status": "ok"}`
- `POST /api/audit` (multipart form: `file`, optional `title`, `client` query
  params) → `application/zip` containing `<name>_Summary.pptx` and
  `<name>_Summary.json`.

## CI/CD secrets required (GitHub repo → Settings → Secrets and variables → Actions)

| Secret | Value |
|---|---|
| `LINODE_HOST` | Linode public IP or hostname |
| `LINODE_USER` | SSH user (e.g. `deploy`) |
| `LINODE_SSH_KEY` | Private key matching a public key installed on the Linode for that user |

`GITHUB_TOKEN` is provided automatically by Actions and used to push/pull
images from GHCR (`ghcr.io/<owner>/<repo>`).

---

## Setting up the Linode instance

### 1. Create the instance
- Linode dashboard → Create → Linode. Ubuntu 24.04 LTS, Nanode/Shared 2GB is
  plenty for this workload. Choose a region close to Akamai's edge/your users.
- Add your SSH public key during creation (or after, via `ssh-copy-id`).

### 2. Base hardening
```bash
ssh root@<linode-ip>
adduser deploy && usermod -aG sudo deploy
rsync --archive --chown=deploy:deploy ~/.ssh /home/deploy
ufw allow OpenSSH && ufw allow 80 && ufw allow 443 && ufw enable
```
Disable root SSH login and password auth in `/etc/ssh/sshd_config`
(`PermitRootLogin no`, `PasswordAuthentication no`), then `systemctl restart ssh`.

### 3. Install Docker
```bash
curl -fsSL https://get.docker.com | sh
usermod -aG docker deploy
```

### 4. Prepare the deploy directory
```bash
su - deploy
sudo mkdir -p /opt/rpo-audit && sudo chown deploy:deploy /opt/rpo-audit
cd /opt/rpo-audit
mkdir -p nginx/certs certbot-www
```
Copy these files from the repo to `/opt/rpo-audit` on the server (once,
manually, or via `scp`): `docker-compose.prod.yml`, `nginx/nginx.conf`,
`.env.example` (rename to `.env` and fill in real values). These rarely
change — routine code changes deploy via CI, not by re-copying files.

### 5. TLS certificate for the origin
Since Akamai is your public edge, terminate TLS again at the Linode origin
("Origin over HTTPS" in Akamai — recommended over plaintext origin traffic).
Easiest path: a certificate issued for your origin hostname.
```bash
sudo apt install certbot
sudo certbot certonly --standalone -d origin.your-domain.example.com
sudo cp /etc/letsencrypt/live/origin.your-domain.example.com/fullchain.pem /opt/rpo-audit/nginx/certs/
sudo cp /etc/letsencrypt/live/origin.your-domain.example.com/privkey.pem /opt/rpo-audit/nginx/certs/
```
Set up a cron/systemd timer to renew and re-copy certs, or automate via a
`certbot renew --deploy-hook` script that restarts the `nginx` container.

### 6. First deploy (manual, one-time)
```bash
cd /opt/rpo-audit
echo "IMAGE=ghcr.io/<owner>/<repo>:latest" >> .env
docker login ghcr.io -u <owner> -p <a GHCR read token>
docker compose -f docker-compose.prod.yml up -d
curl -sk https://localhost/healthz
```
From then on, CI/CD handles redeploys automatically on every push to `main`.

### 7. Point Akamai at the origin
- In Akamai (Property Manager / App & API Protector), set the origin
  hostname to your Linode IP or `origin.your-domain.example.com`.
- Enable "Origin over HTTPS" using the cert from step 5.
- Configure the public hostname (your domain, e.g. `reports.your-domain.com`)
  in Akamai Edge DNS or your DNS provider, pointing to the Akamai edge
  hostname it assigns.
- In the WAF/App & API Protector policy, add rate limiting and a max request
  body size around 30MB on the upload path to match `client_max_body_size`
  in `nginx.conf` / `MAX_UPLOAD_BYTES` in `.env`.

### 8. Verify end-to-end
Visit `https://reports.your-domain.com`, upload a sample report, confirm the
zip download contains the `.pptx` and `.json`.
