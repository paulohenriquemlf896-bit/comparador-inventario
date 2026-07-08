# Arquitetura

## Visão geral

```
Navegador (HTTP)
      │
      ▼
Frontend (templates/ + static/) ── HTML5, CSS3, JS vanilla, Chart.js, Jinja2
      │  HTTP/REST
      ▼
Backend (app.py) ── Flask, sessions, processamento de Excel, geração de PDF
      │
      ▼
SQLite (comparacoes.db) ── usuarios, comparacoes, comparacao_detalhes, alertas
      │
      ▼
Sistemas auxiliares ── logging (app.log), storage local (uploads/)
```

## Stack tecnológico

| Camada | Tecnologia |
|--------|-----------|
| Backend | Flask, Python |
| Frontend | HTML5, CSS3, JavaScript ES6+ (sem framework) |
| Gráficos | Chart.js |
| Banco | SQLite |
| Excel | pandas, openpyxl |
| PDF | ReportLab |
| Segurança | Werkzeug (hash de senha) |

## Modelo de dados

```sql
CREATE TABLE usuarios (
    id INTEGER PRIMARY KEY,
    username TEXT UNIQUE NOT NULL,
    senha TEXT NOT NULL,              -- hash Werkzeug
    nome TEXT NOT NULL,
    email TEXT,
    unidade TEXT NOT NULL,
    é_admin BOOLEAN DEFAULT 0,
    ativo BOOLEAN DEFAULT 1,
    criado_em TIMESTAMP
);

CREATE TABLE comparacoes (
    id INTEGER PRIMARY KEY,
    usuario_id INTEGER NOT NULL,      -- FK usuarios
    unidade TEXT NOT NULL,
    data_criacao TIMESTAMP,
    arquivo_inventario TEXT,
    arquivo_posicao TEXT,
    total_itens INTEGER,
    itens_iguais INTEGER,
    itens_diferentes INTEGER,
    itens_falta INTEGER,
    resultado_json TEXT
);

CREATE TABLE comparacao_detalhes (
    id INTEGER PRIMARY KEY,
    comparacao_id INTEGER NOT NULL,   -- FK comparacoes
    codigo TEXT,
    descricao TEXT,
    loja TEXT,
    qtd_inv REAL,
    valor_inv REAL,
    qtd_pos REAL,
    valor_pos REAL,
    diff_qtd REAL,
    diff_valor REAL,
    status TEXT                       -- OK, DIFERENCA, FALTA
);

CREATE TABLE alertas (
    id INTEGER PRIMARY KEY,
    comparacao_id INTEGER NOT NULL,   -- FK
    usuario_id INTEGER NOT NULL,      -- FK
    tipo_alerta TEXT,                 -- ALERTA_CRITICO, ALERTA_AVISO
    mensagem TEXT,
    taxa_acuracidade REAL,
    limite_alerta REAL,
    criado_em TIMESTAMP,
    lido BOOLEAN DEFAULT 0
);
```

## Fluxo de comparação

1. **Upload**: usuário seleciona unidade e envia os dois Excel; arquivos são salvos em `uploads/`
2. **Extração**: leitura de cada planilha, conversão de tipos e formatação do código como `0000001`
3. **Comparação**: merge outer pelas chaves `codigo + loja`, cálculo de `diff_qtd`/`diff_valor` e status
4. **Armazenamento**: grava em `comparacoes` e `comparacao_detalhes`, cria `alertas` se a taxa de acurácia
   ficar abaixo dos limites configurados
5. **Saída**: exibição na interface, geração de Excel/PDF, filtros por status/loja/descrição

## Rotas principais

```
GET  /                              → redireciona para /login
GET  /login                         → página de login
POST /login                         → valida credenciais
GET  /logout                        → encerra sessão
GET  /signup                        → página de cadastro
POST /signup                        → cria novo usuário

GET  /dashboard                     → dashboard principal
GET  /admin                         → painel administrativo (requer é_admin)

POST /api/comparar                  → executa a comparação
GET  /api/historico                 → lista histórico
GET  /api/comparacao/<id>           → detalhes de uma comparação
GET  /api/comparacoes-filtrado      → busca com filtros (status, loja, descrição)
GET  /api/estatisticas              → estatísticas dos últimos 30 dias
GET  /api/alertas                   → alertas do usuário
GET  /api/download-excel            → download em Excel
GET  /api/download-pdf/<id>         → download em PDF
GET  /api/usuarios                  → lista usuários (admin)
POST /api/usuario/<id>/toggle       → ativa/desativa usuário (admin)
```

## Segurança

- Senhas com hash via `werkzeug.security.generate_password_hash`
- Queries parametrizadas (proteção contra SQL injection)
- Autenticação obrigatória via decorator `@login_required` nas rotas sensíveis
- Acesso admin restrito por flag `é_admin` no banco
- Painel admin consome dados via API (evita injeção de JSON no HTML)

## Logging

Configurado em `app.py` via módulo `logging` padrão do Python:

```python
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log'),
        logging.StreamHandler()
    ]
)
```

Registra tentativas de login, início/fim de comparações, downloads e exceções (com stack trace).
Senhas nunca são logadas — apenas usernames.

## Limitações atuais

- SQLite: recomendado até ~10 usuários simultâneos
- Excel: até ~10.000 linhas recomendado por arquivo
- Upload: limite de 50 MB (configurável em `app.config['MAX_CONTENT_LENGTH']`)
- Servidor de desenvolvimento Flask (Werkzeug) — não usar em produção sem WSGI server (ver [DEPLOY.md](DEPLOY.md))
