# Guia de Deploy para Produção

A aplicação roda hoje com o servidor de desenvolvimento do Flask, adequado apenas para uso local/interno.
Para produção, siga as recomendações abaixo.

## Checklist antes do deploy

- [ ] Definir `SECRET_KEY` no ambiente (ou em um arquivo `.env` na raiz do projeto) com um valor
      aleatório e seguro — sem isso, `app.py` cai no valor padrão de desenvolvimento e registra um
      aviso no log
  ```bash
  SECRET_KEY=<valor aleatório seguro>
  ```
- [ ] Não definir `FLASK_DEBUG=1` em produção (o padrão já é desligado; `run.bat`/`run.ps1` ligam
      debug apenas para uso local)
- [ ] Definir `HOST=0.0.0.0` se o servidor de produção precisar aceitar conexões de fora da própria
      máquina (o padrão é `localhost`)
- [ ] Configurar HTTPS/TLS (Let's Encrypt é gratuito)
- [ ] Configurar backup periódico de `comparacoes.db`

## Opção 1: Windows Server

`gunicorn` não roda no Windows (depende do módulo `fcntl`, exclusivo de sistemas Unix). Use
`waitress` — puro Python, já está em `requirements.txt`.

`run_production.bat`:
```batch
@echo off
cd /d C:\Apps\comparador-inventario
set SECRET_KEY=seu_secret_muito_seguro_aqui
waitress-serve --host=0.0.0.0 --port=5000 app:app
```

Registrar como serviço com NSSM:
```batch
nssm install ComparadorInventario "C:\Apps\comparador-inventario\run_production.bat"
nssm start ComparadorInventario
```

## Opção 2: Docker

`Dockerfile`:
```dockerfile
FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt gunicorn
COPY . .
ENV FLASK_APP=app
CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:5000", "app:app"]
```

`docker-compose.yml`:
```yaml
version: '3.8'
services:
  web:
    build: .
    ports:
      - "5000:5000"
    environment:
      - SECRET_KEY=${SECRET_KEY}
    volumes:
      - ./comparacoes.db:/app/comparacoes.db
    restart: always
```

## Opção 3: Linux (VPS/Cloud) com Nginx + Gunicorn + Supervisor

```bash
sudo apt update
sudo apt install python3.12 python3-pip nginx supervisor
pip3 install -r requirements.txt gunicorn
```

`/etc/nginx/sites-available/comparador`:
```nginx
server {
    listen 80;
    server_name seu-dominio.com;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }

    location /static/ {
        alias /home/usuario/comparador/static/;
    }
}
```

`/etc/supervisor/conf.d/comparador.conf`:
```ini
[program:comparador]
directory=/home/usuario/comparador
command=gunicorn -w 4 -b 127.0.0.1:5000 app:app
autostart=true
autorestart=true
user=usuario
```

## Segurança em produção

`app.py` já carrega `SECRET_KEY`, `FLASK_DEBUG`, `HOST` e `PORT` do ambiente (via `python-dotenv`,
lendo um `.env` na raiz do projeto se existir). Sem `SECRET_KEY` definida, a aplicação sobe com uma
chave padrão insegura e grava um aviso no log — não deixe isso acontecer em produção.

`.env` de exemplo (não commitar):
```bash
SECRET_KEY=seu_secret_muito_seguro_aqui_com_caracteres_aleatorios
HOST=0.0.0.0
PORT=5000
```

Headers de segurança:
```python
@app.after_request
def set_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'SAMEORIGIN'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Strict-Transport-Security'] = 'max-age=31536000; includeSubDomains'
    return response
```

Rate limiting no login:
```python
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

limiter = Limiter(app, key_func=get_remote_address)

@app.route('/login', methods=['POST'])
@limiter.limit("5 per minute")
def login():
    ...
```

HTTPS com Let's Encrypt:
```bash
sudo certbot certonly --webroot -w /var/www/html -d seu-dominio.com
sudo systemctl enable certbot.timer
```

## Escalando além do SQLite

- Migrar para PostgreSQL quando o volume/concorrência crescer:
  ```bash
  sqlite3 comparacoes.db .dump | psql database_name
  ```
- Adicionar cache com Redis (`flask-caching`)
- Load balancing com múltiplas instâncias Gunicorn atrás do Nginx

## Pós-deploy

- [ ] Confirmar que a aplicação responde em produção
- [ ] Confirmar que HTTPS funciona
- [ ] Testar upload/download de arquivos
- [ ] Confirmar acesso ao painel admin
- [ ] Configurar monitoramento de erros (ex.: Sentry) e alertas de uptime

## Troubleshooting

| Problema | Solução |
|----------|---------|
| `Address already in use` | `lsof -i :5000` e finalizar o processo na porta |
| `Permission denied` | Ajustar permissões do diretório/serviço |
| Arquivos estáticos não carregam | Conferir o bloco `location /static/` no Nginx |
| Banco corrompido | Restaurar do backup mais recente de `comparacoes.db` |
