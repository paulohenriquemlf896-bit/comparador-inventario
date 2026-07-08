#!/bin/bash
# Sobe o Comparador de Inventário em produção com waitress (porta 6000).
set -e
cd "$(dirname "$0")"

if [ ! -f .env ]; then
    echo "ERRO: arquivo .env não encontrado. Crie um com SECRET_KEY antes de rodar em produção."
    exit 1
fi

python3 -c "
from waitress import serve
from app import app
serve(app, host='0.0.0.0', port=6000)
"
