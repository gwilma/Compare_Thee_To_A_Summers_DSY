#!/usr/bin/env bash
# Run the app as a Docker container behind a Traefik that already owns ports 80/443
# (for example the one Hostinger's Docker templates install). deploy_vps.sh calls this
# when it finds Traefik; it can also be run on its own, as root:
#
#   sudo DOMAIN=dsy.example.com APP_DIR=/home/dsy/app DATA_DIR=/home/dsy/data \
#        HTPASSWD=/etc/nginx/dsy.htpasswd bash scripts/deploy_traefik.sh
#
# The entry point, certificate resolver and Docker network are read from the running
# Traefik container. Override them if needed: TRAEFIK_ENTRYPOINT (HTTPS entry point),
# TRAEFIK_HTTP_ENTRYPOINT, TRAEFIK_CERTRESOLVER, TRAEFIK_NETWORK.
set -euo pipefail

: "${APP_DIR:?set APP_DIR}" "${DATA_DIR:?set DATA_DIR}" "${HTPASSWD:?set HTPASSWD}"
DOMAIN="${DOMAIN:-}"
APP_USER="${APP_USER:-dsy}"
ENV_FILE="${ENV_FILE:-/etc/dsy.env}"
NAME="${CONTAINER_NAME:-summers-dsy}"
IMAGE="summers-dsy"

log() { printf '\n\033[1;34m==> %s\033[0m\n' "$*"; }
die() { printf '\033[1;31mError: %s\033[0m\n' "$*" >&2; exit 1; }

# The ID of the running Traefik container: the one publishing ports 80/443, else (host networking) the first.
traefik_container() {
    local rows
    rows=$(docker ps --format '{{.ID}}|{{.Image}}|{{.Ports}}' 2>/dev/null | awk -F'|' 'tolower($2) ~ /(^|\/)traefik([:@]|$)/')
    awk -F'|' '$3 ~ /:(80|443)->/ {print $1; exit}' <<<"$rows" | grep . || awk -F'|' 'NF {print $1; exit}' <<<"$rows"
}

[[ -n "$DOMAIN" ]] || die "Traefik on this server routes websites by domain name, so the app needs one.
Point a DNS A record (e.g. dsy.yourdomain.com) at this server, then re-run with DOMAIN=dsy.yourdomain.com."
[[ -s "$HTPASSWD" ]] || die "no password file at ${HTPASSWD}"

tid=$(traefik_container)
[[ -n "$tid" ]] || die "no running Traefik container found"

log "Reading the Traefik configuration"
# Static configuration given as command-line flags or TRAEFIK_* environment variables.
cfg=$(
    docker inspect -f '{{range .Args}}{{println .}}{{end}}' "$tid" | sed -E 's/^--//'
    docker inspect -f '{{range .Config.Env}}{{println .}}{{end}}' "$tid" \
        | sed -nE 's/^TRAEFIK_([^=]+)=(.*)$/\1=\2/p' | awk -F= -v OFS== '{k = tolower($1); gsub("_", ".", k); $1 = k; print}'
)
entrypoint_on() {  # name of the entry point listening on port $1 (Traefik option names ignore case)
    sed -nE "s/^entrypoints\.([A-Za-z0-9_-]+)\.address=.*:$1\$/\1/Ip" <<<"$cfg" | head -1
}
ep_https="${TRAEFIK_ENTRYPOINT:-$(entrypoint_on 443)}"
ep_http="${TRAEFIK_HTTP_ENTRYPOINT:-$(entrypoint_on 80)}"
resolver="${TRAEFIK_CERTRESOLVER:-$(sed -nE 's/^certificatesresolvers\.([A-Za-z0-9_-]+)\..*/\1/Ip' <<<"$cfg" | head -1)}"
network="${TRAEFIK_NETWORK:-$(sed -nE 's/^providers\.docker\.network=(.+)$/\1/Ip' <<<"$cfg" | head -1)}"
if [[ -z "$network" ]]; then
    network=$(docker inspect -f '{{range $k, $v := .NetworkSettings.Networks}}{{println $k}}{{end}}' "$tid" | grep -v '^$' | head -1)
fi
[[ "$network" == "host" ]] && network=""  # Traefik on the host network reaches containers on the default bridge

if [[ -z "$ep_https" || -z "$resolver" ]]; then
    die "could not tell which HTTPS entry point (found: '${ep_https}') and certificate resolver (found: '${resolver}')
Traefik uses; it is probably configured from a file. Look for them with:
  docker inspect ${tid} | grep -iE 'entrypoints|certresolver|certificatesresolvers'
then re-run with TRAEFIK_ENTRYPOINT=<name> TRAEFIK_CERTRESOLVER=<name>."
fi
echo "Traefik container ${tid}: HTTPS entry point '${ep_https}', HTTP entry point '${ep_http:-none}',"
echo "certificate resolver '${resolver}', network '${network:-bridge}'"

log "Building the app image"
docker build -q -t "$IMAGE" "$APP_DIR" >/dev/null

log "Starting the app container"
docker rm -f "$NAME" >/dev/null 2>&1 || true
users=$(grep -v '^$' "$HTPASSWD" | paste -sd, -)
r="summers-dsy"
labels=(
    -l "traefik.enable=true"
    -l "traefik.http.routers.${r}.rule=Host(\`${DOMAIN}\`)"
    -l "traefik.http.routers.${r}.entrypoints=${ep_https}"
    -l "traefik.http.routers.${r}.tls=true"
    -l "traefik.http.routers.${r}.tls.certresolver=${resolver}"
    -l "traefik.http.routers.${r}.middlewares=${r}-auth"
    -l "traefik.http.routers.${r}.service=${r}"
    -l "traefik.http.middlewares.${r}-auth.basicauth.users=${users}"
    -l "traefik.http.services.${r}.loadbalancer.server.port=8501"
)
if [[ -n "$ep_http" ]]; then  # plain HTTP redirects to HTTPS
    labels+=(
        -l "traefik.http.routers.${r}-http.rule=Host(\`${DOMAIN}\`)"
        -l "traefik.http.routers.${r}-http.entrypoints=${ep_http}"
        -l "traefik.http.routers.${r}-http.middlewares=${r}-https"
        -l "traefik.http.routers.${r}-http.service=${r}"
        -l "traefik.http.middlewares.${r}-https.redirectscheme.scheme=https"
        -l "traefik.http.middlewares.${r}-https.redirectscheme.permanent=true"
    )
fi
run_args=(--name "$NAME" --restart unless-stopped -v "${DATA_DIR}:/data")
id "$APP_USER" &>/dev/null && run_args+=(--user "$(id -u "$APP_USER"):$(id -g "$APP_USER")")
[[ -f "$ENV_FILE" ]] && run_args+=(--env-file "$ENV_FILE")
run_args+=(-e SUMMERS_DSY_DATA=/data)  # after --env-file, so it wins over the host path there
if [[ -n "$network" ]]; then
    run_args+=(--network "$network")
    labels+=(-l "traefik.docker.network=${network}")
fi
docker run -d "${run_args[@]}" "${labels[@]}" "$IMAGE" >/dev/null

log "Checking the app"
status=""
for _ in $(seq 1 60); do
    status=$(docker inspect -f '{{.State.Health.Status}}' "$NAME" 2>/dev/null || true)
    [[ "$status" == "healthy" ]] && break
    sleep 2
done
if [[ "$status" == "healthy" ]]; then
    echo "App container is running."
else
    echo "App container is not healthy yet (${status:-unknown}). See: docker logs ${NAME}"
fi
code=$(curl --noproxy '*' -sk -o /dev/null -w '%{http_code}' --max-time 10 --resolve "${DOMAIN}:443:127.0.0.1" "https://${DOMAIN}/" || true)
case "$code" in
    401) echo "Traefik is routing https://${DOMAIN} to the app (it asks for the password)." ;;
    *)   echo "Traefik answered HTTP ${code:-nothing} for https://${DOMAIN}; it can take a minute to pick up the app."
         echo "If it persists: docker logs ${tid} 2>&1 | tail -20" ;;
esac
