FROM caddy:2-alpine

# Page statique, aucune dépendance de build.
COPY mobitag-envoi.html /srv/index.html
COPY Caddyfile /etc/caddy/Caddyfile

# Caddy tourne déjà sans privilèges dans cette image, mais on l'explicite :
# le conteneur n'a besoin d'écrire nulle part.
EXPOSE 8080

HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD wget -q -O /dev/null http://127.0.0.1:8080/ || exit 1
