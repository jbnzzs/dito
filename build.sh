#!/usr/bin/env bash
set -euo pipefail

# Build da aplicacao Django no Render.
python -m pip install -r requirements.txt
python manage.py collectstatic --noinput

# Migracoes nao sao executadas automaticamente em producao.
# Habilite somente apos revisar as migracoes e confirmar um backup.
if [[ "${RUN_MIGRATIONS_ON_BUILD:-0}" == "1" ]]; then
    python manage.py migrate --noinput
fi
