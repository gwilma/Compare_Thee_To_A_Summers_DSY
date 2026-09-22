#!/usr/bin/env bash
# Install (or update) "Compare thee to a summer's DSY" on an Ubuntu VPS.
#
# Runs the Streamlit app as a systemd service behind nginx, with a password
# (HTTP basic auth) and, if you give a domain, HTTPS from Let's Encrypt.
# Safe to re-run: it pulls the latest code and restarts the app. Your library
# and download cache live in DATA_DIR and are never touched.
#
# Usage (as root):
#   sudo DOMAIN=dsy.example.com EMAIL=you@example.com bash deploy_vps.sh
#   sudo bash deploy_vps.sh                      # no domain: http://<server-ip>
#
# Optional settings (environment variables):
#   DOMAIN         domain pointing at this server (enables HTTPS)
#   EMAIL          email for Let's Encrypt expiry notices
#   GITHUB_TOKEN   token with read access, if the repository is private
#   AUTH_USER      login name for the site              (default: dsy)
#   AUTH_PASSWORD  login password (default: ask, or keep the existing one)
#   CEDA_TOKEN     CEDA access token for MIDAS Open downloads
#   BRANCH         git branch  (default: claude/weather-overheating-analysis-y14bhk)
#   REPO_URL       git URL     (default: the GitHub repository)
set -euo pipefail

REPO_URL="${REPO_URL:-https://github.com/gwilma/Compare_Thee_To_A_Summers_DSY.git}"
BRANCH="${BRANCH:-claude/weather-overheating-analysis-y14bhk}"
APP_USER="dsy"
APP_HOME="/home/${APP_USER}"
APP_DIR="${APP_HOME}/app"
DATA_DIR="${APP_HOME}/data"
PORT=8501
DOMAIN="${DOMAIN:-}"
EMAIL="${EMAIL:-}"
AUTH_USER="${AUTH_USER:-dsy}"
AUTH_PASSWORD="${AUTH_PASSWORD:-}"
HTPASSWD="/etc/nginx/dsy.htpasswd"

log() { printf '\n\033[1;34m==> %s\033[0m\n' "$*"; }
die() { printf '\033[1;31mError: %s\033[0m\n' "$*" >&2; exit 1; }

[[ $EUID -eq 0 ]] || die "run as root (e.g. sudo bash $0)"
command -v apt-get >/dev/null || die "this script is for Ubuntu/Debian"

log "Installing system packages"
export DEBIAN_FRONTEND=noninteractive
apt-get update -q
apt-get install -y -q python3 python3-venv python3-pip git nginx apache2-utils ufw curl

log "Creating app user '${APP_USER}'"
id "$APP_USER" &>/dev/null || adduser --disabled-password --gecos "" "$APP_USER"
install -d -o "$APP_USER" -g "$APP_USER" "$DATA_DIR"

log "Fetching code (${BRANCH})"
git_auth=()
if [[ -n "${GITHUB_TOKEN:-}" ]]; then
    basic=$(printf 'x-access-token:%s' "$GITHUB_TOKEN" | base64 -w0)
    git_auth=(-c "http.extraHeader=Authorization: Basic ${basic}")
fi
if [[ -d "${APP_DIR}/.git" ]]; then
    sudo -u "$APP_USER" git "${git_auth[@]}" -C "$APP_DIR" fetch -q origin "$BRANCH"
    sudo -u "$APP_USER" git -C "$APP_DIR" checkout -q -B "$BRANCH" "origin/${BRANCH}"
    sudo -u "$APP_USER" git -C "$APP_DIR" reset -q --hard "origin/${BRANCH}"
else
    sudo -u "$APP_USER" git "${git_auth[@]}" clone -q -b "$BRANCH" "$REPO_URL" "$APP_DIR" \
        || die "clone failed; if the repository is private, set GITHUB_TOKEN"
fi

log "Installing Python packages"
[[ -x "${APP_DIR}/.venv/bin/python" ]] || sudo -u "$APP_USER" python3 -m venv "${APP_DIR}/.venv"
sudo -u "$APP_USER" "${APP_DIR}/.venv/bin/pip" install -q --upgrade pip
sudo -u "$APP_USER" "${APP_DIR}/.venv/bin/pip" install -q -e "$APP_DIR"

log "Writing systemd service"
ENV_FILE="/etc/dsy.env"
if [[ ! -f "$ENV_FILE" ]]; then
    echo "SUMMERS_DSY_DATA=${DATA_DIR}" > "$ENV_FILE"
fi
if [[ -n "${CEDA_TOKEN:-}" ]]; then
    sed -i '/^CEDA_TOKEN=/d' "$ENV_FILE"
    echo "CEDA_TOKEN=${CEDA_TOKEN}" >> "$ENV_FILE"
fi
chmod 600 "$ENV_FILE"

cat > /etc/systemd/system/dsy.service <<EOF
[Unit]
Description=Compare thee to a summer's DSY
After=network.target

[Service]
User=${APP_USER}
WorkingDirectory=${APP_DIR}
EnvironmentFile=${ENV_FILE}
ExecStart=${APP_DIR}/.venv/bin/streamlit run app/streamlit_app.py --server.address 127.0.0.1 --server.port ${PORT} --server.headless true --browser.gatherUsageStats false
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF
systemctl daemon-reload
systemctl enable -q dsy
systemctl restart dsy

log "Setting the site password"
if [[ -n "$AUTH_PASSWORD" ]]; then
    htpasswd -bcB "$HTPASSWD" "$AUTH_USER" "$AUTH_PASSWORD" 2>/dev/null
elif [[ ! -s "$HTPASSWD" ]]; then
    if [[ -t 0 ]]; then
        echo "Choose a password for user '${AUTH_USER}':"
        htpasswd -cB "$HTPASSWD" "$AUTH_USER"
    else
        AUTH_PASSWORD=$(tr -dc 'A-Za-z0-9' </dev/urandom | head -c 16)
        htpasswd -bcB "$HTPASSWD" "$AUTH_USER" "$AUTH_PASSWORD" 2>/dev/null
        GENERATED_PASSWORD="$AUTH_PASSWORD"
    fi
else
    echo "Keeping the existing password in ${HTPASSWD}"
fi
chown root:www-data "$HTPASSWD"
chmod 640 "$HTPASSWD"

log "Configuring nginx"
cat > /etc/nginx/sites-available/dsy <<EOF
server {
    listen 80;
    listen [::]:80;
    server_name ${DOMAIN:-_};

    client_max_body_size 200M;

    location / {
        auth_basic "Compare thee to a summer's DSY";
        auth_basic_user_file ${HTPASSWD};

        proxy_pass http://127.0.0.1:${PORT};
        proxy_http_version 1.1;
        proxy_set_header Upgrade \$http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host \$host;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
        proxy_read_timeout 86400;
    }
}
EOF
ln -sf /etc/nginx/sites-available/dsy /etc/nginx/sites-enabled/dsy
rm -f /etc/nginx/sites-enabled/default
nginx -t -q
systemctl reload nginx

log "Opening the firewall (SSH, HTTP, HTTPS)"
ssh_port=$(ss -tlnpH 2>/dev/null | awk '/sshd/ {sub(/.*:/, "", $4); print $4; exit}')
ufw allow "${ssh_port:-22}/tcp" >/dev/null
ufw allow OpenSSH >/dev/null 2>&1 || true
ufw allow 'Nginx Full' >/dev/null
ufw --force enable >/dev/null

if [[ -n "$DOMAIN" ]]; then
    log "Requesting an HTTPS certificate for ${DOMAIN}"
    apt-get install -y -q certbot python3-certbot-nginx
    email_args=(--register-unsafely-without-email)
    [[ -n "$EMAIL" ]] && email_args=(-m "$EMAIL")
    certbot --nginx -d "$DOMAIN" --non-interactive --agree-tos --redirect "${email_args[@]}" \
        || echo "Certificate request failed: check that ${DOMAIN}'s DNS A record points at this server, then re-run."
fi

log "Checking the app"
for _ in $(seq 1 30); do
    curl -fs "http://127.0.0.1:${PORT}/_stcore/health" >/dev/null && break
    sleep 1
done
if curl -fs "http://127.0.0.1:${PORT}/_stcore/health" >/dev/null; then
    echo "App is running."
else
    echo "App did not respond yet. See: journalctl -u dsy -n 50"
fi

ip=$(curl -fs4 --max-time 5 https://api.ipify.org || hostname -I | awk '{print $1}')
url="http://${ip}"
[[ -n "$DOMAIN" ]] && url="https://${DOMAIN}"
cat <<EOF

Done. Open: ${url}
Login user: ${AUTH_USER}${GENERATED_PASSWORD:+
Password:   ${GENERATED_PASSWORD}   (generated; change with: htpasswd -B ${HTPASSWD} ${AUTH_USER})}

Update later: re-run this script.
App logs:     journalctl -u dsy -f
Your data:    ${DATA_DIR} (back this up)
EOF
