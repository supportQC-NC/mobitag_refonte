FROM caddy:2-alpine

# Page statique, aucune dépendance de build.
COPY mobitag-envoi.html /srv/index.html

RUN printf ':8080 {\n\troot * /srv\n\tfile_server\n\tencode gzip zstd\n}\n' > /etc/caddy/Caddyfile

EXPOSE 8080
