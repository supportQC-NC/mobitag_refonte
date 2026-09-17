FROM caddy:2-alpine

# Ressources statiques : aucune étape de build, aucune dépendance à installer.
# Le HTML, le CSS et le JS sont des fichiers distincts, servis tels quels.
COPY src/ /srv/
COPY Caddyfile /etc/caddy/Caddyfile

EXPOSE 8080

HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD wget -q -O /dev/null http://127.0.0.1:8080/ || exit 1
