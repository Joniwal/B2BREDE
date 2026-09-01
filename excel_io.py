"""Transações locais: lock cooperativo entre processos, backup e troca atômica.

NÃO é um lock distribuído. Power Automate, Excel Web e OneDrive não participam
deste protocolo. Todos os escritores locais devem usar o mesmo caminho/lock.
"""
from __future__ import annotations

import errno
import hashlib
import logging
import os
import random
import tempfile
import time
import zipfile
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4

log = logging.getLogger(__name__)


class ExcelSafetyError(RuntimeError):
    def __init__(self, message, code="EXCEL_SAFETY_ERROR", status=409):
        super().__init__(message)
        self.code, self.status = code, status


def digest(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()


@contextmanager
def local_lock(path: Path, timeout: float = 15):
    """Não apagar o arquivo .lock: isso poderia criar locks em inodes distintos."""
    lock_path = path.with_name(path.name + ".lock")
    fd = os.open(lock_path, os.O_RDWR | os.O_CREAT, 0o600)
    acquired = False
    deadline = time.monotonic() + timeout
    try:
        if os.fstat(fd).st_size == 0:
            os.write(fd, b"0")
        while True:
            try:
                os.lseek(fd, 0, os.SEEK_SET)
                if os.name == "nt":
                    import msvcrt
                    msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
                acquired = True
                break
            except OSError as exc:
                if exc.errno not in {errno.EACCES, errno.EAGAIN, errno.EDEADLK}:
                    raise
                if time.monotonic() >= deadline:
                    raise ExcelSafetyError("Arquivo em uso; tente novamente.", "EXCEL_BUSY", 423) from exc
                time.sleep(0.1 + random.uniform(0, 0.1))
        yield
    finally:
        if acquired:
            os.lseek(fd, 0, os.SEEK_SET)
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(fd, fcntl.LOCK_UN)
        os.close(fd)


def validate_package(path: Path):
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
        if bad or "xl/workbook.xml" not in archive.namelist():
            raise ExcelSafetyError("O arquivo temporário não passou na validação ZIP.", "EXCEL_INVALID_ZIP")


def commit_snapshot(path: Path, staged: Path, original: bytes, backup_dir: Path):
    """Chamar SOMENTE enquanto local_lock estiver adquirido.

    A checagem SHA detecta muitas interferências externas, mas não elimina a
    janela entre comparação e replace para escritores que ignoram o lock.
    """
    validate_package(staged)
    # _commit/FlushFileBuffers no Windows exige um descritor gravável.
    with staged.open("rb+") as stream:
        os.fsync(stream.fileno())
    if digest(path.read_bytes()) != digest(original):
        raise ExcelSafetyError("Arquivo alterado por outro serviço. Recarregue antes de salvar.", "EXCEL_CONFLICT")
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
    backup = backup_dir / f"{path.stem}.{stamp}.{uuid4().hex}{path.suffix}"
    # Backup exato dos bytes de entrada, nunca um workbook reserializado.
    with backup.open("xb") as stream:
        stream.write(original)
        stream.flush()
        os.fsync(stream.fileno())
    if digest(path.read_bytes()) != digest(original):
        raise ExcelSafetyError("Conflito detectado antes da publicação; original mantido.", "EXCEL_CONFLICT")
    os.replace(staged, path)  # Mesmo diretório/volume; não é uma transação na nuvem.
    log.info("excel_commit file=%s backup=%s", path.name, backup.name)


def transact(path, transform, *, backup_dir=None, lock_timeout=15, attempts=4):
    """transform(original_bytes, staged_path) -> (resultado, precisa_salvar).

    Retenta apenas PermissionError ANTES de um commit bem-sucedido. Cada tentativa
    recarrega o estado mais recente. Conflitos de dados/esquema não são repetidos.
    """
    path = Path(path).resolve()
    backup_dir = Path(backup_dir).resolve() if backup_dir else path.parent / "_rede_backups"
    if not path.is_file():
        raise ExcelSafetyError("Excel não encontrado; criação automática desativada.", "EXCEL_NOT_FOUND", 404)
    for attempt in range(attempts):
        staged = None
        try:
            with local_lock(path, lock_timeout):
                # Sinal conservador: não equivale a todos os locks do Excel/servidor.
                if path.with_name("~$" + path.name).exists():
                    raise PermissionError("Excel Desktop mantém arquivo de proprietário.")
                original = path.read_bytes()
                fd, name = tempfile.mkstemp(prefix=".rede-stage-", suffix=path.suffix, dir=path.parent)
                os.close(fd)
                staged = Path(name)
                result, changed = transform(original, staged)
                if changed:
                    commit_snapshot(path, staged, original, backup_dir)
                return result
        except PermissionError as exc:
            if attempt == attempts - 1:
                raise ExcelSafetyError("Excel bloqueado ou sem permissão. Feche o arquivo e tente novamente.", "EXCEL_BUSY", 423) from exc
            time.sleep(min(2 ** attempt, 4) + random.uniform(0, 0.2))
        finally:
            if staged is not None:
                try:
                    staged.unlink(missing_ok=True)  # Apenas o temporário desta transação.
                except OSError:
                    log.warning("Temporário retido para limpeza manual: %s", staged.name)
    raise AssertionError("Número de tentativas inválido")
