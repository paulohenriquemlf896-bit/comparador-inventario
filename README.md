# Comparador de Inventário

Aplicação web em Flask para comparar automaticamente o **Inventário Fiscal** com a **Posição Diária**
de Estoque, identificando discrepâncias de quantidade e valor por unidade.

## Features

- Autenticação com login/logout e cadastro de novos usuários (`/signup`)
- Upload de dois arquivos Excel (Inventário e Posição) e comparação automática por código + loja
- Relatório com status por item: `OK`, `DIFERENCA`, `FALTA NO INVENTARIO`, `FALTA NA POSICAO`
- Download do resultado em Excel (colorido) ou PDF
- Histórico de comparações e detalhes por comparação (SQLite)
- Filtros avançados por status, loja e descrição
- Alertas automáticos quando a taxa de acurácia cai abaixo de limites configuráveis
- Painel de administração (`/admin`) para gerenciar usuários
- Dark mode com persistência em `localStorage` e detecção de preferência do sistema
- Sistema de logging (login, comparações, erros, downloads) em `app.log`

## Stack

- **Backend**: Flask 3.x, Python
- **Frontend**: HTML5, CSS3, JavaScript (vanilla), Chart.js
- **Banco de dados**: SQLite (`comparacoes.db`)
- **Processamento**: pandas, openpyxl
- **PDF**: ReportLab
- **Segurança**: Werkzeug (hash de senha)

## Instalação e execução

```bash
pip install -r requirements.txt
python app.py
```

No Windows, também é possível usar `run.bat` (duplo-clique) ou `run.ps1`.

Acesse: `http://localhost:5000`

**Credenciais demo**: `admin` / `admin123`

## Como usar

1. Faça login (ou crie uma conta em `/signup`)
2. Selecione a unidade
3. Faça upload dos arquivos **Inventário Fiscal** e **Posição Diária** (`.xlsx` ou `.xls`)
4. Clique em **Comparar Arquivos**
5. Visualize o resultado, aplique filtros se necessário
6. Baixe o relatório em **Excel** ou **PDF**

## Estrutura dos arquivos de entrada

### Inventário Fiscal
| Coluna | Conteúdo |
|--------|----------|
| C | Código do produto |
| E | Loja |
| F | Descrição |
| K | Quantidade |
| N | Valor total |

### Posição Diária
| Coluna | Conteúdo |
|--------|----------|
| A | Código |
| C | Descrição |
| H | Loja |
| K | Quantidade |
| L | Valor total |

## Estrutura do projeto

```
.
├── app.py                # Aplicação Flask principal
├── requirements.txt
├── run.bat / run.ps1     # Scripts de inicialização (Windows)
├── comparacoes.db         # Banco SQLite (usuários, comparações, alertas)
├── static/
│   ├── css/               # style, animations, icons, theme-toggle
│   └── js/                # themes.js (dark mode)
├── templates/             # login, signup, dashboard, admin
├── uploads/                # Pasta temporária de upload (runtime)
└── docs/
    ├── ARQUITETURA.md      # Arquitetura, schema do banco, rotas da API
    └── DEPLOY.md           # Guia de deploy para produção
```

## Troubleshooting

- **Python não encontrado**: instale em https://www.python.org
- **Módulos faltando**: `pip install -r requirements.txt`
- **Porta 5000 em uso**: altere `port=5000` no final de `app.py`
- **Arquivo não reconhecido**: confirme que é `.xlsx`/`.xls` e segue a estrutura de colunas acima

---

Desenvolvido para Rancho Alegre Produtos Agropecuários.
