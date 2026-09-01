# Validação local — 29/08/2026

- Python 3.12.13 / Windows.
- openpyxl 3.1.5; pandas 2.3.3; pytest 8.4.2; Flask 3.1.3.
- `python -m pytest -q`: **38 testes aprovados**.
- JavaScript do frontend: verificação de sintaxe aprovada.
- Módulos Python principais, exemplos e alternativa nativa: compilação aprovada.

Cobertura: CRUD e dashboard, consultas sem mudança de hash/mtime, estilos,
tabela estruturada/ranges, validação, formatação condicional, fórmulas e gráfico
simples de fixture, backup, falha/retry de replace, conflito externo detectável,
If-Match, dois processos locais, chave/esquema inválidos, recusa de objetos
complexos, ampliação da mesma tabela mantendo nome/ID/TableStyle, cópia de
estilo/fórmula/validação/formatação condicional, bloqueio de conteúdo abaixo e
staging de comandos com repetição idempotente no arquivo local. A descoberta
OneDrive/SharePoint cobre arquivo único, ausência, nome inválido, roots
configurados e recusa de múltiplas cópias com o mesmo nome.

Não homologados: arquivo REDE_B2B.xlsx real, aparência no Excel Desktop, macros,
ActiveX, pivôs/slicers reais, Power Query, cenário distribuído OneDrive/SharePoint
e execução de Office Script/pywin32. O script TypeScript deve ser validado no
editor Office Scripts e no fluxo de homologação antes de produção.

Os arquivos de teste foram gerados em diretórios temporários. Nenhum workbook
operacional do usuário foi aberto para escrita ou modificado. Os testes com
falha de gravação confirmam manutenção do arquivo original.
