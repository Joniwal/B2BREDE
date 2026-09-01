# Perguntas iniciais

Antes de habilitar escrita no arquivo real, envie:

1. Uma cópia anonimizada de REDE_B2B.xlsx, preservando tabelas, estilos, fórmulas,
   gráficos, pivôs e conexões. Se possível, uma versão anterior e outra posterior
   ao problema. Não converta para CSV, pois isso remove justamente os objetos.
2. O que significa “formatação rede”: cores, estilo de tabela, nome/ID da tabela,
   validações, fórmulas, segmentadores, gráficos ou pivôs? Envie capturas do antes/depois.
3. O caminho de armazenamento e a forma de acesso: disco local, SMB, pasta
   sincronizada pelo OneDrive, biblioteca SharePoint ou cópia baixada via API.
4. Aba, linha de cabeçalho, nome da tabela estruturada, chave única/imutável,
   quantidade de linhas, fórmulas nos campos de entrada e formato XLSX/XLSM/XLSB.
5. Script em execução, versão do Python/pandas/openpyxl, sistema operacional,
   Excel instalado e configuração do servidor (processos/workers/instâncias).
6. Power Automate: conexão OneDrive/SharePoint, gatilhos, ações, concorrência,
   retries, duração e erros do histórico. Informe também edições manuais.
   Remova tokens, senhas, URLs de gatilhos assinadas e dados pessoais.

# Solução recomendada

## Diagnóstico confirmado no código anterior

O adaptador distribuído anteriormente tinha dois problemas concretos:

- `fallback_excel_read()` criava IDs/campos faltantes e chamava
  `fallback_excel_write()` durante uma consulta.
- `fallback_excel_write()` executava `workbook.remove(sheet)` e recriava a aba.
  Isso destruía tabela estruturada, referências e objetos daquela aba; reaplicar
  cores e AutoFilter não recriava a tabela original usada pelo Power Automate.

Portanto, a leitura de pandas isoladamente não é a causa. O gatilho destrutivo
era a gravação/recriação. Se outro serviço volta a produzir o arquivo sem os
campos técnicos, essa migração pode reaparecer em acessos subsequentes.

Um DataFrame representa dados, não o pacote completo do Excel. `to_excel()` em
modo de escrita substitui o arquivo; `if_sheet_exists='replace'` recria a aba.
`overlay` também não garante preservar objetos ou concorrência e pode deixar
dados antigos se o novo conjunto for menor. Referência:
https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_excel.html

## Escolha recomendada por cenário

- **Contenção imediata:** consultas sem qualquer save/migração; habilitar
  `EXCEL_READ_ONLY=true` no arquivo real até concluir a homologação.
- **Arquivo XLSX simples local:** usar o adaptador conservador deste pacote.
- **App + Power Automate + vários usuários:** lista SharePoint/banco como fonte
  de verdade; um único serviço materializa a tabela/relatório Excel. O modo
  SharePoint do app já existe; migrar dados/configurar o fluxo depende do tenant.
- **Macros, pivôs, slicers, Power Query, Data Model, formas ou controles:** não
  reserializar com openpyxl sem homologação; usar Excel nativo controlado ou
  atualizar a tabela pela plataforma Microsoft, conforme os objetos envolvidos.

Não é possível garantir fidelidade de um arquivo ainda não recebido. O perfil
local bloqueia recursos complexos em vez de salvá-los silenciosamente.

# Implementação técnica

## Arquivos alterados/adicionados

- `excel_safe.py`: CRUD pontual, esquema explícito, proteção de fórmulas, leitura
  sem gravação, triagem de objetos e verificação de partes OOXML.
- `excel_io.py`: lock cooperativo por caminho, snapshot, backup exato, validação
  ZIP, retry limitado de arquivo ocupado e publicação por `os.replace`.
- `sharepoint_client.py`: adaptador Excel agora delega ao motor seguro; filtros,
  paginação, ordenação e agregações do app permanecem.
- `app.py` e `static/js/main.js`: PATCH envia somente campos alterados. `_etag`
  e `If-Match` detectam uma versão antiga do registro; erro 412 pede recarga.
- `excel_native.py`: alternativa de PATCH via pywin32 em cópia temporária.
- `automation/update_table.ts`: Office Script de upsert/delete por chave,
  atualizado exclusivamente por um fluxo escritor serializado.
- `automation/staging_commands.py`: exemplo de staging de comandos JSON.
- `scripts/inspect_excel.py`: inventário somente leitura.
- `tests/test_excel_preservation.py`: testes de regressão e concorrência local.

## Instalação

Base validada nesta revisão: Python 3.12.13, openpyxl 3.1.5, pandas 2.3.3,
pytest 8.4.2, Flask 3.1.3. O projeto usa Python 3.10+; revalide ao mudar versões.
Pandas continua disponível para análises, mas não grava o workbook operacional.
O lock usa a biblioteca padrão (`msvcrt` no Windows, `fcntl` no POSIX).
Pillow é uma dependência de suporte a imagens; a preservação de imagens reais
não foi homologada nesta revisão.

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item .env.example .env
.\.venv\Scripts\python.exe -m pytest -q
```

Em ambiente já configurado, não sobrescreva `.env`: incorpore apenas as novas
variáveis. Nunca copie senhas para repositório ou arquivo ZIP de entrega.

Configuração inicial para uma CÓPIA de homologação:

```dotenv
USE_EXCEL_FALLBACK=true
EXCEL_PATH=
EXCEL_FILENAME=REDE_B2B.xlsx
EXCEL_SEARCH_ROOTS=
EXCEL_SHEET=REDEB2B
EXCEL_TABLE=
EXCEL_KEY_COLUMN=IDCLIENTE
EXCEL_HEADER_ROW=1
EXCEL_BACKUP_DIR=./backups
EXCEL_READ_ONLY=true
FLASK_DEBUG=false
```

Use o nome real da tabela, não seu estilo visual. `EXCEL_TABLE=` vazio detecta
a única tabela da aba. Mais de uma tabela exige nome explícito. Sem tabela,
o cabeçalho é localizado na linha indicada. Não há escolha silenciosa da primeira aba.

Com `EXCEL_PATH=` vazio, o app procura `EXCEL_FILENAME` nos mounts locais do
OneDrive pessoal/corporativo e nas bibliotecas SharePoint registradas pelo
cliente de sincronização. A descoberta só aceita uma correspondência. Se houver
duplicatas, use `EXCEL_SEARCH_ROOTS` para limitar a biblioteca ou renomeie as
cópias; nunca é correto escolher a primeira silenciosamente.

`_ITEM_ID` é a chave do projeto anterior. Você pode usar `IDCLIENTE` somente se
ele for único e imutável; se vários registros têm o mesmo cliente, crie uma
coluna técnica estável durante manutenção (com todos os escritores parados).
Não use números de linha como IDs. Atualize o Power Automate para preencher a
mesma chave ao criar registros. O conector impõe restrições a colunas ocultas
e nomes em filtros OData: prefira uma chave visível alfanumérica, como
`IDREGISTRO`, no desenho definitivo. Não renomeie uma chave em produção sem
ajustar todos os consumidores. `_ITEM_ID` pode ser usado diretamente pelo Office
Script, que não depende de Filter Query/OData.

## Uso direto do motor openpyxl

Execute na raiz deste projeto, depois de homologar uma cópia simples:

```python
from excel_safe import SafeExcel
from sharepoint_client import FIELDS, DATE_FIELDS

excel = SafeExcel(
    r"C:\Homologacao\REDE_B2B.xlsx", "Dados", FIELDS,
    table="TabelaRede", key="_ITEM_ID", date_fields=DATE_FIELDS,
    backup_dir=r"C:\Homologacao\backups",
)

registro = excel.get("1")  # Apenas leitura.
excel.update("1", {"STATUS": "Concluído"},
             expected_version=registro["_etag"])
novo = excel.create({"IDCLIENTE": "B2B-999", "CLIENTE": "Cliente teste"})
excel.delete(novo["id"], expected_version=novo["_etag"])
```

O caminho de gravação local é:

1. Adquirir o lock do caminho canônico (`arquivo.xlsx.lock`).
2. Ler os bytes originais e abrir o workbook em memória, `data_only=False`.
3. Localizar cabeçalhos/chave, validar duplicidades e proteger fórmulas.
4. Alterar somente `cell.value`. Não remover aba, tabela, linha ou coluna.
5. Salvar num arquivo temporário no MESMO diretório/volume.
6. Validar ZIP, reabertura e ausência de partes OOXML perdidas.
7. Verificar se o original continua igual ao snapshot, guardar backup exato,
   verificar novamente e substituir com `os.replace`.
8. Liberar o lock. Em falha antes do replace, manter o original.

Internamente, a exclusão limpa apenas valores dos campos gerenciados e da
chave, com `.value = None`. Não desloca linhas nem apaga estilos. Colunas extras
e fórmulas de relatório permanecem; a linha física vazia é um slot reutilizável.
Power Automate e relatórios precisam ignorar chave vazia. Se houver dados
privados fora dos campos gerenciados, essa exclusão não os apaga: defina sua
política de retenção e o conjunto completo de campos antes de operar.

**Tabela estruturada e formatação:** se houver slot vazio, ele é reutilizado. Se
a tabela estiver cheia, o motor aumenta uma linha no `ref` da MESMA tabela; não
remove/recria aba, tabela ou workbook. `Table.name`, `id`, `displayName`, colunas
e `TableStyle` permanecem. A nova linha copia estilo/altura da última linha,
traduz fórmulas de colunas auxiliares e estende as validações e regras
condicionais que alcançavam a linha-modelo. O novo registro entra normalmente
nos filtros, paginação, CRUD e agregações/Chart.js do app.

Para não sobrescrever nada, a inclusão é recusada antes do save quando existe
conteúdo, outra tabela ou célula mesclada logo abaixo. Tabela com linha de totais
também exige Excel nativo/Office Script homologado. Objetos de gráfico continuam
no arquivo e não são recriados; ranges diretos de gráficos e nomes definidos
permanecem como estavam. Para que um gráfico Excel passe a incluir novas linhas,
use a tabela/referência estruturada como origem. Pivôs, VBA e outros objetos
complexos seguem bloqueados no perfil openpyxl. A exclusão limpa os valores e
mantém a linha formatada dentro da tabela, pronta para reutilização.

Em aba simples sem tabela, o motor só anexa quando não há ranges dependentes.
Não apagamos um intervalo inteiro para substituí-lo por um DataFrame; o PATCH é
mais conservador.

**Limites da integridade:** lock protege apenas escritores que utilizam este
protocolo no mesmo arquivo local. O SHA detecta diversas alterações externas,
mas ainda existe uma janela de corrida entre comparação e replace para quem
ignora o lock. OneDrive em computadores diferentes não compartilha o lock.
`os.replace` no disco local não equivale a commit atômico no SharePoint.
Disco/rede/antivírus/ACLs precisam ser homologados; não há promessa de
durabilidade universal em queda de energia ou armazenamento SMB/NFS.

**Conflitos de usuário:** a UI envia If-Match. APIs externas devem fazer o
mesmo. Sem esse cabeçalho, o último PATCH vence nos campos que ele envia.
Retries internos de PermissionError reabrem o snapshot, mas o servidor não
possui um log durável de idempotência para POST; não repita criação após timeout
sem reconciliação. Na nuvem, use OperationId/EntityId estáveis em uma fila durável.

## Power Automate: padrão de escritor único

Não use a pasta sincronizada como banco transacional. Não deixe o Flask e o
conector editarem o mesmo workbook em paralelo.

Padrão recomendado:

1. Grave o CRUD na lista SharePoint/banco. Gere uma chave de registro estável e
   uma versão. O Excel é projeção/relatório, não a fonte de verdade concorrente.
2. Enfileire alterações em lista/serviço durável com OperationId único, EntityId,
   versão, ação, payload, status, número de tentativas e mensagem de falha.
3. Um ÚNICO fluxo processa essa fila. Habilite controle de concorrência do
   gatilho com grau 1 e desative paralelismo do Apply to each. Outros fluxos,
   serviços e usuários também não podem escrever nesse mesmo arquivo.
4. Use Update a row/Add a row/Delete a row na tabela existente, ou Run script
   com `automation/update_table.ts`. Passe nome da tabela, coluna-chave,
   EntityId estável, upsert/delete e changesJson. Para datas, informe o sistema
   1900/1904 real do workbook; o exemplo aceita datas a partir de 1904.
5. Processar apenas a versão atual de cada registro (ou consultar sua versão
   mais recente antes do upsert) evita aplicar eventos antigos fora de ordem.
6. Em 429 respeite Retry-After quando exposto; para lock/transiente use retry
   exponencial, por exemplo 5, 10, 20, 40, 80, 160 segundos, com jitter e limite
   total de 10 minutos. Em validação/403 pare e alerte. Em 412 releia/reconcilie;
   nunca sobrescreva automaticamente a versão mais nova.
7. Após resposta incerta/time-out, consulte a chave/versão antes de repetir.
   Não use Add a row cegamente: retry pode duplicar. O upsert usa a mesma chave,
   mas não é uma transação exatamente-uma-vez. Só libere o próximo comando após
   confirmar/reconciliar o atual. Falhas persistentes vão para revisão.
8. Confirme o resultado por polling limitado (não um Delay fixo tratado como
   mutex), então marque o comando como concluído. Mantenha registro processado
   pelo prazo de deduplicação. Não remova o histórico logo após sucesso.

O exemplo de staging Python escreve somente comandos JSON, num destino
separado, por temporário + rename. Não altera o workbook. Se a entrada for uma
pasta OneDrive, o fluxo deve reagir apenas a `.json` final, ignorar `.tmp`,
`.lock`, backups e staging; ainda são necessários deduplicação e monitoramento
de sincronização. A lista/fila durável é preferível à pasta.

```python
from automation.staging_commands import enqueue

# UUID estável da operação: persista/reutilize o mesmo no retry.
enqueue("./entrada-comandos", "f19819c0-4f85-47ea-b22f-b84974e0ad3d", "1",
        {"STATUS": "Concluído"}, action="upsert")
```

**Update file não é Update a row.** Ele substitui o conteúdo binário do arquivo
inteiro. Só considere publicação integral com todos os escritores suspensos,
arquivo staging validado, backup/versionamento e precondição ETag no endpoint
que explicitamente a suporte. Se a ação escolhida não oferece precondição,
comparar metadados e depois chamar Update file não é compare-and-swap atômico.
Não apague e recrie o arquivo final: IDs/links/triggers podem mudar. Não envie
um workbook reconstruído por pandas por cima do arquivo formatado.

O conector documenta lock de até seis minutos, atraso de até 30 segundos nas
alterações e falta de suporte a modificações simultâneas. Uma leitura do
conector também pode gerar nova versão por mecanismos internos. Logo, não
use "arquivo modificado" como fila recursiva sem condições de origem/versão.
Fonte: https://learn.microsoft.com/en-us/connectors/excelonlinebusiness/

# Plano de testes e implantação

## O que foi testado localmente

- Os testes originais de API, CRUD, filtros, ordenação, dashboard e validação.
- GET da página, itens, detalhes, metadados e dashboard repetido sem alterar
  bytes/hash/mtime do workbook nem criar backups/lock.
- Estilos, bordas, fontes, formatos, larguras, altura, panes, nomes definidos,
  regras condicionais, validações e tabela estruturada após PATCH/CREATE/DELETE.
- Preservação do XML de um gráfico simples e fórmulas em aba separada.
- Backup byte a byte, falha antes da substituição e retry de arquivo ocupado.
- Dois processos Python alterando campos diferentes sem perder a outra edição.
- If-Match desatualizado, chave duplicada, ampliação segura de tabela cheia,
  bloqueio por conteúdo abaixo, fórmula protegida, texto com '=' e recusa de
  arquivo com objeto complexo.

```powershell
python -m pytest -q
python scripts\inspect_excel.py "C:\Homologacao\REDE_B2B.xlsx"
python examples\update_record.py --file "C:\Homologacao\REDE_B2B.xlsx" --sheet Dados --table TabelaRede --id 1 --patch-file examples\patch.json
# Só após revisar a simulação:
python examples\update_record.py --file "C:\Homologacao\REDE_B2B.xlsx" --sheet Dados --table TabelaRede --id 1 --patch-file examples\patch.json --apply
```

## Homologação obrigatória no seu ambiente

1. Suspenda gravações locais, fluxos e edições manuais. Preserve uma versão
   íntegra imutável; se a formatação já se perdeu, restaure o histórico/backup.
   Uma biblioteca não consegue adivinhar os objetos que já foram apagados.
2. Faça inventário e compare os dois arquivos no Excel Desktop, não apenas
   lendo DataFrames. Inspecione tabela/nome/ID/ref, gráficos/series, pivôs/cache,
   slicers, conexões, macros/assinaturas, nomes definidos, validação e estilo.
3. Mantenha `EXCEL_READ_ONLY=true` na cópia até confirmar o perfil suportado.
4. Compare hash e data de modificação antes/depois de 20 acessos simultâneos.
5. Teste criação, edição, exclusão, datas e ampliação da tabela cheia; confira
   nome/ID/estilo/ref da tabela, os 21 campos, filtro por data, modais e gráficos
   Chart.js. Gráficos do app são distintos dos objetos de gráfico/pivô do Excel.
6. Teste arquivo aberto no Excel, permissão negada, interrupção de gravação,
   duas edições concorrentes, conflito com escritor externo e recuperação.
7. Duplique o fluxo Power Automate apontando para um arquivo/lista de teste.
   Verifique IDs, reentrada de gatilho, filas, deduplicação, lock, leitura após
   gravação e resultado final, não apenas o status verde de uma ação.
8. Recalcule fórmulas/atualize pivôs pelo método suportado no ambiente e compare
   os resultados. Não aceite um gráfico vazio ou cache antigo como preservado.
9. Faça piloto; só então aponte o caminho/lista de produção e reative UM escritor.
10. Monitore logs, erros 409/412/423/429, tempo de fila, número de backups e
    disponibilidade de espaço. Defina retenção e teste restauração periodicamente.

Os testes usam fixtures locais; não houve teste com o arquivo real, Excel
Desktop, macros/pivôs reais ou tenant Power Automate. A biblioteca não recalcula
as fórmulas do fixture. Não foi realizada validação visual no Excel instalado.

# Alternativas e notas importantes

| Opção | Vantagem | Limite |
|---|---|---|
| openpyxl pontual | Leve, sem Excel instalado, bom para XLSX comum | Regrava o pacote OOXML; objetos não suportados e caches exigem testes |
| pywin32/xlwings | Usa Excel nativo; maior fidelidade para recursos complexos | Windows/Excel/licença, processo controlado; não usar dentro de cada request Flask |
| Staging de dados/comandos | Separa ingestão de layout e reduz conflitos | Exige fila, deduplicação, escritor e tratamento de falhas |
| Lista SharePoint/banco + Excel relatório | Melhor caminho multiusuário | Migração/fluxo/projeção precisam ser configurados no tenant |

**VBA:** não é correto afirmar que openpyxl nunca preserva macros. Com
`keep_vba=True`, pode manter partes VBA em XLSM, mas não executa/edita macros e
não garante ActiveX, formas, assinaturas ou todos os objetos associados.
Por segurança, o adaptador deste pacote recusa escrita em XLSM. Fontes:
https://openpyxl.readthedocs.io/en/stable/tutorial.html
https://openpyxl.readthedocs.io/en/stable/pivot.html

Openpyxl não calcula fórmulas; resultados em cache podem desaparecer ao salvar.
Gráficos simples/pivôs suportados podem ser mantidos, mas caches/conexões e
recursos não suportados precisam de homologação. Para inserir/remover linhas,
a biblioteca não atualiza todas as dependências; por isso esta solução mantém
slots/ranges. Fonte:
https://openpyxl.readthedocs.io/en/stable/editing_worksheets.html

O Office Script é um exemplo não instalado nem executado no tenant. Métodos
de refresh de pivôs e muitas conexões não funcionam em flows mesmo quando
retornam sucesso; valide o mecanismo de refresh separadamente:
https://learn.microsoft.com/en-us/office/dev/scripts/testing/power-automate-troubleshooting

Alternativa nativa com pywin32:

```powershell
python -m pip install -r requirements-native.txt
```

```python
from excel_native import update_native

update_native(
    r"C:\Homologacao\REDE_B2B.xlsm", "Dados", "TabelaRede", "_ITEM_ID", "1",
    {"STATUS": "Concluído"}, backup_dir=r"C:\Homologacao\backups",
)
```

Essa alternativa preserva o formato ao salvar uma cópia com o próprio Excel;
nunca faz conversão XLSM->XLSX. Macros/eventos são desabilitados, links não são
atualizados e a instância criada é fechada em finally. Não use com workbook
não confiável, protegido por senha, macros XLM ou conexões auto-refresh sem
revisão prévia. Falhas COM não são repetidas cegamente. xlwings pode usar o
mesmo padrão (`App`, `books.open`, `range.value`, `save`, `close`), mas não
elimina essas limitações. Microsoft não recomenda/suporta automação Office
não interativa no servidor:
https://support.microsoft.com/en-US/Visio/considerations-for-server-side-automation-of-office

Quick-fix de contenção, caso não possa trocar o adaptador imediatamente:

```python
import pandas as pd

def ler_apenas(caminho, aba="Dados"):
    # Não incluir save, to_excel, ExcelWriter, migração de IDs ou criação de aba.
    return pd.read_excel(caminho, sheet_name=aba,
                         dtype=str, keep_default_na=False, engine="openpyxl")
```

Bloqueie temporariamente POST/PATCH/DELETE no adaptador antigo. Retirar apenas
a gravação do GET não corrige a recriação de aba durante o CRUD. O pacote novo
substitui ambos os caminhos, mantendo as consultas e análises do app.

Envie o arquivo REDE_B2B.xlsx ou uma cópia mínima reproduzível, o script atual,
as versões do Python/pandas e os detalhes do Power Automate (OneDrive/SharePoint,
gatilhos e ações usadas) para homologar a solução no seu arquivo real.
