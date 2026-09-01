"""Publica comandos JSON numa pasta de entrada SEPARADA do Excel.

Exemplo local de staging. Não é uma fila distribuída nem publica no SharePoint.
O caminho oficial recomendado é uma lista/fila durável com escritor único.
"""
import json
import os
import tempfile
from pathlib import Path
from uuid import UUID

from excel_io import local_lock


def enqueue(folder, operation_id, entity_id, changes, action="upsert"):
    operation_id = str(UUID(str(operation_id)))  # nunca usar nome vindo de URL diretamente
    if action not in {"upsert", "delete"}:
        raise ValueError("Ação inválida.")
    folder = Path(folder).resolve()
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / f"{operation_id}.json"
    payload = json.dumps({"operationId": operation_id, "entityId": str(entity_id),
                          "action": action, "changes": changes}, sort_keys=True,
                         ensure_ascii=False).encode("utf-8")
    with local_lock(target):
        if target.exists():
            if target.read_bytes() != payload:
                raise ValueError("OperationId reutilizado com outro conteúdo.")
            return target
        fd, name = tempfile.mkstemp(prefix=".command-", suffix=".tmp", dir=folder)
        temp = Path(name)
        try:
            with os.fdopen(fd, "wb") as stream:
                stream.write(payload)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temp, target)
        finally:
            temp.unlink(missing_ok=True)
    return target
