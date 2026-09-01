/** Office Script: chamar exclusivamente pelo único fluxo escritor.
 * Não contém credenciais. Não calcula pivôs/Power Query. O OperationId e a
 * ordem/retries devem ser controlados numa lista/fila durável fora do workbook.
 * action = upsert ou delete; entityId é estável (não gere outro UUID no retry).
 */
function main(
  workbook: ExcelScript.Workbook,
  tableName: string,
  keyColumn: string,
  entityId: string,
  action: string,
  changesJson: string,
  dateSystem: string = "1900"
): string {
  const table = workbook.getTable(tableName);
  if (!table) throw new Error("Tabela não encontrada.");
  if (!entityId || !["upsert", "delete"].includes(action)) throw new Error("Comando inválido.");
  if (/^[=+@-]/.test(entityId) || !["1900", "1904"].includes(dateSystem)) throw new Error("Chave/sistema de datas inválido.");
  const fields = ["IDCLIENTE", "CLIENTE", "ENDERECO", "CIDADE", "PRODUTO", "ATIVIDADE", "TECNOLOGIA", "VT", "DATADISPARO", "RETORNOPCC", "DATAAGENDAMENTO", "DATACONCLUSAO", "OBSERVACAO", "STATUS", "EXECUTADOPOR", "TIPOCABO", "METRAGEM", "OBSERVACAOCONCLUSAO", "NUMDRAFT", "ROTA", "USUARIO"];
  const patch = JSON.parse(changesJson) as { [key: string]: string | number | boolean | null };
  if (!patch || Array.isArray(patch) || typeof patch !== "object") throw new Error("JSON inválido.");
  const headers = table.getHeaderRowRange().getTexts()[0];
  const keyIndex = headers.indexOf(keyColumn);
  if (keyIndex < 0) throw new Error("Chave não encontrada.");
  const columns = Object.keys(patch);
  const dates = ["DATADISPARO", "RETORNOPCC", "DATAAGENDAMENTO", "DATACONCLUSAO"];
  for (const name of columns) {
    if (name === keyColumn || !fields.includes(name) || !headers.includes(name)) throw new Error("Coluna não autorizada: " + name);
    const value = patch[name];
    if (value !== null && !["string", "number", "boolean"].includes(typeof value)) throw new Error("Somente escalares.");
    if (typeof value === "string" && /^[=+@-]/.test(value)) throw new Error("Texto potencialmente interpretável como fórmula.");
    if (dates.includes(name) && value !== null && value !== "") {
      if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value)) throw new Error("Data deve usar AAAA-MM-DD.");
      const instant = new Date(value + "T00:00:00Z");
      if (isNaN(instant.getTime()) || instant.toISOString().slice(0, 10) !== value || value < "1904-01-01") throw new Error("Data inválida ou anterior a 1904.");
      const base = dateSystem === "1904" ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
      patch[name] = (instant.getTime() - base) / 86400000;
    }
  }
  const rowCount = table.getRowCount();
  const values = rowCount ? table.getRangeBetweenHeaderAndTotal().getTexts() : [];
  const matches: number[] = [];
  values.forEach((row, i) => { if (row[keyIndex] === entityId) matches.push(i); });
  if (matches.length > 1) throw new Error("Chave duplicada.");
  if (action === "delete") {
    if (matches.length) table.deleteRowsAt(matches[0], 1);
    return JSON.stringify({ id: entityId, deleted: true });
  }
  let rowIndex: number;
  if (matches.length) {
    rowIndex = matches[0];
    const body = table.getRangeBetweenHeaderAndTotal();
    for (const name of columns) {
      if (body.getCell(rowIndex, headers.indexOf(name)).getFormula().startsWith("=")) throw new Error("Fórmula protegida: " + name);
    }
    for (const name of columns) {
      const cell = body.getCell(rowIndex, headers.indexOf(name));
      if (patch[name] === null) cell.clear(ExcelScript.ClearApplyTo.contents);
      else cell.setValue(patch[name] as string | number | boolean);
    }
  } else {
    if (!patch["IDCLIENTE"] || !patch["CLIENTE"]) throw new Error("IDCLIENTE e CLIENTE obrigatórios na criação.");
    if (rowCount && table.getRangeBetweenHeaderAndTotal().getFormulas().some(row => row.some(value => value.startsWith("=")))) throw new Error("Criação em tabela com fórmulas requer template homologado.");
    // Adicionar a linha já com a chave reduz risco de duplicação após timeout.
    const newRow: (string | number | boolean)[] = headers.map(name => name === keyColumn ? entityId : patch[name] ?? "");
    table.addRow(-1, newRow);
    rowIndex = table.getRowCount() - 1;
  }
  for (const name of columns.filter(name => dates.includes(name))) {
    const cell = table.getRangeBetweenHeaderAndTotal().getCell(rowIndex, headers.indexOf(name));
    if (patch[name] !== null && patch[name] !== "" && cell.getNumberFormat() === "General") cell.setNumberFormat("yyyy-mm-dd");
  }
  return JSON.stringify({ id: entityId, row: rowIndex, updated: true });
}
