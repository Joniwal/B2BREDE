# REDEB2B — Painel Web (Flask + Excel local)

Aplicação web para gerenciar os registros de agendamento de clientes e
atividades de rede (REDEB2B), com CRUD completo, busca, filtros, paginação,
ordenação e um dashboard analítico com gráficos. Os dados são lidos e
gravados diretamente em um arquivo Excel local (`.xlsx`), sem nenhuma
dependência de SharePoint, Azure AD ou Microsoft Graph.

---

## 1. Arquitetura

```
redeb2b_app/
├── app.py                  # Inicialização do Flask, registro de blueprint, error handlers
├── api.py                  # Blueprint com os endpoints REST (/api/...)
├── excel_client.py         # Camada de dados: leitura/escrita no Excel via pandas/openpyxl
├── gerar_dados_exemplo.py  # Script para popular um Excel de teste
├── requirements.txt
├── .env.example
├── templates/
│   ├── index.html          # Página única: topbar, filtros, KPIs, gráficos, tabela
│   └── modals.html         # Modais de criar/editar, excluir e "itens por data"
└── static/
    ├── css/styles.css
    └── js/main.js           # fetch/async-await, paginação, ordenação, Chart.js
```

**`excel_client.py`** expõe a classe `DataClient`, com os métodos
`list_items`, `get_item`, `create_item`, `update_item`, `delete_item`,
`dashboard_aggregates` e `items_by_date`. O arquivo Excel usado é decidido
assim:
1. Se `EXCEL_PATH` estiver preenchido no `.env`, esse caminho é usado
   diretamente (modo manual).
2. Se `EXCEL_PATH` estiver em branco, o sistema procura automaticamente um
   arquivo chamado `EXCEL_FILENAME` (padrão: `REDE_B2B.xlsx`) dentro das
   pastas do OneDrive sincronizadas nesta máquina (pessoal e/ou
   corporativo), inclusive dentro de subpastas.
3. `EXCEL_SEARCH_ROOTS` é opcional. Deixe em branco para que o mesmo pacote
   funcione em computadores com nomes de usuário e caminhos diferentes. Use-o
   apenas quando precisar restringir a busca a uma biblioteca específica.

Nas escritas (criar/atualizar/excluir), o sistema usa `openpyxl` para alterar
somente os valores necessários dentro da Tabela estruturada indicada por
`EXCEL_TABLE`. A tabela não é recriada: nome, estilo, cores e formatação são
preservados. O salvamento usa um arquivo temporário no mesmo volume e só
substitui o original depois da validação. Mantenha o arquivo fechado no Excel
Desktop durante uma gravação para evitar bloqueio de concorrência.

**Endpoints da API** (`api.py`, todas as respostas em JSON):

| Método | Rota | Descrição |
|---|---|---|
| GET | `/api/items` | Lista paginada, com filtros e ordenação |
| GET | `/api/items/<id>` | Detalhe de um registro (por IDCLIENTE) |
| POST | `/api/items` | Cria um registro |
| PATCH | `/api/items/<id>` | Atualiza um registro |
| DELETE | `/api/items/<id>` | Exclui um registro |
| GET | `/api/dashboard` | Agregações para KPIs e gráficos (Chart.js) |
| GET | `/api/items-by-date?data=YYYY-MM-DD` | Registros agendados em uma data (usado ao clicar no gráfico de linha) |
| GET | `/api/export` | Baixa em `.xlsx` os registros que atendem aos mesmos filtros/ordenação de `/api/items`, sem paginação (todos os que passam pelo filtro) |
| GET | `/api/excel-status` | Diagnóstico: mostra se o Excel foi localizado e em qual caminho (útil para depurar problemas de OneDrive) |

---

## 2. Instalação e configuração

### 2.1 Pré-requisitos
- Python 3.10+
- pip

### 2.2 Passos

```bash
# 1. Entre na pasta do projeto
cd redeb2b_app

# 2. Crie e ative um ambiente virtual (recomendado)
python -m venv venv
source venv/bin/activate        # Linux/Mac
venv\Scripts\activate           # Windows

# 3. Instale as dependências
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

# 4. Copie o arquivo de exemplo de variáveis de ambiente
cp .env.example .env             # Linux/Mac
copy .env.example .env           # Windows

# 5. Edite o .env se necessário (veja abaixo)
```

No `.env`, você tem duas opções:

**Opção 1 — automática (recomendada):** deixe `EXCEL_PATH` em branco. O
sistema vai procurar sozinho o arquivo `REDE_B2B.xlsx` (ou o nome que você
colocar em `EXCEL_FILENAME`) dentro das pastas do OneDrive sincronizadas
nesta máquina — não precisa saber o caminho completo.
```
EXCEL_PATH=
EXCEL_FILENAME=REDE_B2B.xlsx
EXCEL_SEARCH_ROOTS=
EXCEL_SHEET_NAME=REDEB2B
EXCEL_TABLE=REDEB2B
```

Com `EXCEL_SEARCH_ROOTS=` vazio, cada usuário procura no próprio OneDrive.
Isso permite distribuir o mesmo `.exe` e o mesmo `.env` para computadores com
caminhos diferentes, desde que a biblioteca compartilhada esteja sincronizada
e o arquivo mantenha o nome `REDE_B2B.xlsx`.

**Opção 2 — manual:** preencha `EXCEL_PATH` com o caminho completo. Quando
preenchido, ele tem prioridade sobre a busca automática.
```
EXCEL_PATH=C:\Users\PEGGY\OneDrive - Telefonica\PROJETO\REDE_B2B.xlsx
EXCEL_SHEET_NAME=REDEB2B
EXCEL_TABLE=REDEB2B
```

Para conferir se o arquivo foi encontrado (e onde), acesse no navegador,
com a aplicação rodando:
```
http://localhost:5000/api/excel-status
```

Se você ainda não tem o arquivo, gere um de exemplo com dados fictícios:
```bash
python gerar_dados_exemplo.py
# ou, para um caminho específico:
python gerar_dados_exemplo.py "C:\Users\PEGGY\Desktop\PROJETO\REDE_B2B.xlsx"
```

### 2.3 Executando a aplicação

```bash
python app.py
```
Acesse **http://localhost:5000** no navegador.

### 2.4 Atalho de duplo-clique (Windows)

Depois de configurar o `venv` uma vez (passos 2.2), dê um duplo-clique em
**`iniciar_redeb2b.bat`** — ele ativa o ambiente virtual, sobe o servidor e
abre o navegador automaticamente em `http://localhost:5000`. Não precisa
abrir terminal nenhum a partir daí.

### 2.5 Gerando um executável (.exe) standalone

Se quiser um `.exe` que roda **mesmo em outro computador sem Python
instalado**, instale o PyInstaller (só é necessário para gerar o `.exe`,
não faz parte das dependências normais da aplicação):

```bash
python -m pip install pyinstaller
```

**Opção A — script `build_exe.bat`** (mais rápido, mas alguns ambientes
corporativos bloqueiam a execução de arquivos `.bat` por política de
segurança):
```
build_exe.bat
```
Após clonar o repositório, esse script cria `venv`, instala as dependências e
o PyInstaller e gera o pacote automaticamente. É necessário ter Python 3.10+
no `PATH` e acesso à internet durante a primeira execução.

**Opção B — comando direto** (funciona em qualquer ambiente, inclusive
onde `.bat` é bloqueado — quem "executa" aqui é o `python.exe`, já
autorizado, não um `.exe`/`.bat` novo):
```bash
python -m PyInstaller --name REDEB2B --noconfirm --onefile --icon static/favicon.ico --add-data "templates;templates" --add-data "static;static" --collect-all pandas --collect-all openpyxl --collect-all flask app.py
```
(a flag `--icon static/favicon.ico` usa o ícone do projeto como ícone do
`.exe` gerado; pode omitir essa flag se não quiser um ícone customizado)

Qualquer uma das duas opções gera **um único arquivo** executável em:
```
dist\REDEB2B.exe
```

O `build_exe.bat` também cria `dist\.env` a partir do `.env.example` quando
esse arquivo ainda não existir. Distribua os dois arquivos juntos.

**Antes de usar/distribuir**, coloque na mesma pasta onde ficar o
`REDEB2B.exe`:
- o seu `.env` já configurado (copiado do `.env.example`, com o
  `EXCEL_PATH`/`EXCEL_FILENAME` corretos) — mantenha `FLASK_DEBUG=false`
  nessa cópia, já que o modo debug do Flask não é recomendado para um
  executável empacotado.

Como é modo `--onefile`, é só copiar o `REDEB2B.exe` (e o `.env` do lado)
para outro computador — sem pasta extra ao redor. O arquivo fica grande
(dezenas a centenas de MB, já que inclui o Python e as bibliotecas), e a

primeira abertura pode demorar alguns segundos a mais (ele extrai tudo
para uma pasta temporária antes de iniciar) — isso é normal.

---

## 3. Segurança e observações

- Como não há mais integração externa (Graph/Azure AD), não existe segredo
  algum a proteger — apenas o caminho do arquivo Excel, que fica só no
  backend (nunca é exposto ao navegador).
- Todos os inputs recebidos pela API são sanitizados (strip de espaços) e
  validados (ex.: datas devem estar em `yyyy-mm-dd`) antes de serem
  gravados no Excel.
- Evite ter o arquivo Excel aberto no próprio Excel/LibreOffice enquanto a
  aplicação estiver rodando — alguns sistemas bloqueiam o arquivo para
  escrita externa enquanto ele está aberto, o que pode gerar erro ao
  salvar. Feche o arquivo no Excel antes de criar/editar/excluir pelo
  painel.
- **Backup:** como a aplicação regrava o arquivo inteiro a cada alteração,
  é recomendável manter uma rotina simples de backup do `.xlsx` (cópia
  periódica), especialmente antes de operações em lote.

---

## 4. Exemplos de uso da API (cURL)

### Listar itens com filtros, paginação e ordenação
```bash
curl "http://localhost:5000/api/items?cliente=Empresa&status=Agendado&page=1&page_size=20&sort=DATAAGENDAMENTO:desc"
```

### Buscar um item pelo IDCLIENTE
```bash
curl "http://localhost:5000/api/items/CLI1001"
```

### Criar um item
```bash
curl -X POST "http://localhost:5000/api/items" \
  -H "Content-Type: application/json" \
  -d '{
    "IDCLIENTE": "CLI2001",
    "CLIENTE": "Nova Empresa LTDA",
    "ENDERECO": "Av. Central, 500",
    "CIDADE": "Campinas",
    "PRODUTO": "Internet Dedicada",
    "ATIVIDADE": "Instalação",
    "TECNOLOGIA": "FTTH",
    "VT": "VT9001",
    "DATAAGENDAMENTO": "2026-08-05",
    "STATUS": "Agendado",
    "EXECUTADOPOR": "João Silva",
    "TIPOCABO": "Drop 1FO",
    "METRAGEM": "120",
    "NUMDRAFT": "DR00099",
    "ROTA": "Rota-2",
    "USUARIO": "admin"
  }'
```

### Atualizar um item (PATCH)
```bash
curl -X PATCH "http://localhost:5000/api/items/CLI2001" \
  -H "Content-Type: application/json" \
  -d '{"STATUS": "Concluído", "DATACONCLUSAO": "2026-08-06"}'
```

### Excluir um item
```bash
curl -X DELETE "http://localhost:5000/api/items/CLI2001"
```

### Dashboard (agregações para os gráficos)
```bash
curl "http://localhost:5000/api/dashboard?cidade=Campinas"
```

### Itens agendados em uma data específica
```bash
curl "http://localhost:5000/api/items-by-date?data=2026-08-05"
```

---

## 5. Testes rápidos

```bash
# 1. Gere dados de exemplo
python gerar_dados_exemplo.py

# 2. Rode a aplicação
python app.py

# 3. Abra http://localhost:5000 e confirme que:
#    - os KPIs e os 4 gráficos aparecem preenchidos
#    - a tabela lista os registros com paginação e permite ordenar por coluna
#      (por padrão, ordenado por Data de Agendamento, mais recente primeiro)
#    - a busca rápida e os filtros do painel esquerdo filtram a tabela e os gráficos
#    - clicar no ícone de "olho" abre o modal de edição com os dados carregados
#    - "Novo registro" abre o modal em branco e salva um registro novo (POST)
#    - o ícone de lixeira pede confirmação e exclui o registro (DELETE)
#    - clicar em um ponto do gráfico "Por Data de Agendamento" abre o modal
#      com a lista de registros daquela data
```
