#!/bin/bash
# Sobe o Comparador de Inventário em produção com waitress (porta 8080).
set -e
cd "$(dirname "$0")"

if [ ! -f .env ]; then
    echo "ERRO: arquivo .env não encontrado. Crie um com SECRET_KEY antes de rodar em produção."
    exit 1
fi

if [ ! -x venv/bin/python3 ]; then
    echo "ERRO: venv não encontrado em ./venv. Rode: python3 -m venv venv && venv/bin/pip install -r requirements.txt"
    exit 1
fi

venv/bin/python3 -c "
from waitress import serve
from app import app
serve(app, host='0.0.0.0', port=8080)
"
