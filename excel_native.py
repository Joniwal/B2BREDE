"""Alternativa para Excel complexo: processo Excel nativo, fora do servidor web.

Exige Windows + Excel instalado + pywin32, numa sessão de usuário licenciada.
Não executar como serviço IIS/Windows não interativo. Homologar no arquivo real.
"""
from pathlib import Path

from excel_io import ExcelSafetyError, transact


def update_native(path, sheet, table, key_column, item_id, changes, *, backup_dir=None):
    """PATCH em tabela estruturada sobre cópia temporária; nunca roda macros.

    Preserva o formato original XLSX/XLSM. Não faz SaveAs nem conversão.
    Arquivos criptografados, assinados e macros XLM requerem revisão separada.
    Apenas valores escalares de entrada; fórmulas existentes são protegidas.
    """
    import pythoncom
    import win32com.client

    if Path(path).suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Use XLSX ou XLSM homologado.")
    if not changes or key_column in changes:
        raise ValueError("Informe alterações sem modificar a chave.")
    if any(not isinstance(v, (str, int, float, bool, type(None))) for v in changes.values()):
        raise ValueError("Aceita somente valores escalares.")

    def transform(original, staged):
        staged.write_bytes(original)
        pythoncom.CoInitialize()
        excel, workbook = None, None
        try:
            # DispatchEx isola esta instância: não altera o Excel aberto do usuário.
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            workbook = excel.Workbooks.Open(
                str(staged), UpdateLinks=0, ReadOnly=False,
                IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False,
                Password="", WriteResPassword="",
            )
            if workbook.ReadOnly:
                raise ExcelSafetyError("Excel abriu a cópia em modo somente leitura.")
            target = workbook.Worksheets(sheet).ListObjects(table)
            headers = {str(target.HeaderRowRange.Cells(1, c).Value2): c
                       for c in range(1, target.ListColumns.Count + 1)}
            if key_column not in headers or set(changes) - set(headers):
                raise ValueError("Chave/colunas não encontradas na tabela.")
            if target.DataBodyRange is None:
                raise ValueError("Tabela vazia.")
            matches = [r for r in range(1, target.ListRows.Count + 1)
                       if str(target.DataBodyRange.Cells(r, headers[key_column]).Value2) == str(item_id)]
            if len(matches) != 1:
                raise ValueError("Chave ausente ou duplicada.")
            row = matches[0]
            # Valida TODAS as alterações antes de atribuir a primeira célula.
            for field in changes:
                if target.DataBodyRange.Cells(row, headers[field]).HasFormula:
                    raise ValueError(f"Fórmula protegida: {field}.")
            for field, value in changes.items():
                cell = target.DataBodyRange.Cells(row, headers[field])
                if value is None:
                    cell.ClearContents()  # não Clear(), que também remove formatos.
                else:
                    # Evita execução de uma fórmula recebida como texto externo.
                    literal = "'" + value if isinstance(value, str) and value.startswith(("=", "+", "-", "@")) else value
                    cell.Value2 = literal
            # Sem RefreshAll automático: conexões podem exigir credenciais e
            # iniciar efeitos externos. Calcular/atualizar deve ser homologado.
            workbook.Save()
            return {"id": str(item_id), "changed": list(changes)}, True
        finally:
            try:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)
            finally:
                try:
                    if excel is not None:
                        excel.Quit()
                finally:
                    pythoncom.CoUninitialize()

    return transact(path, transform, backup_dir=backup_dir)
