"""Repositórios para SharePoint/Microsoft Graph e Excel local.

O restante da aplicação usa a mesma interface, independentemente da fonte.
Filtros, ordenação, paginação e agregações são aplicados neste módulo para que
o comportamento permaneça idêntico nos dois modos.
"""

from __future__ import annotations

import json
import logging
import math
import os
import re
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import msal
import pandas as pd
import requests
from excel_io import ExcelSafetyError
from excel_safe import SafeExcel
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


log = logging.getLogger(__name__)


FIELDS = [
    "IDCLIENTE",
    "CLIENTE",
    "ENDERECO",
    "CIDADE",
    "PRODUTO",
    "ATIVIDADE",
    "TECNOLOGIA",
    "VT",
    "DATADISPARO",
    "RETORNOPCC",
    "DATAAGENDAMENTO",
    "DATACONCLUSAO",
    "OBSERVACAO",
    "STATUS",
    "EXECUTADOPOR",
    "TIPOCABO",
    "METRAGEM",
    "OBSERVACAOCONCLUSAO",
    "NUMDRAFT",
    "ROTA",
    "USUARIO",
]

DATE_FIELDS = {"DATADISPARO", "RETORNOPCC", "DATAAGENDAMENTO", "DATACONCLUSAO"}
FILTER_FIELDS = {"IDCLIENTE", "CLIENTE", "CIDADE", "EXECUTADOPOR", "STATUS"}
SORT_ALIASES = {re.sub(r"[^a-z0-9]", "", field.lower()): field for field in FIELDS}
SORT_ALIASES.update(
    {
        "id": "IDCLIENTE",
        "cliente": "CLIENTE",
        "cidade": "CIDADE",
        "executadopor": "EXECUTADOPOR",
        "status": "STATUS",
        "dataagendamento": "DATAAGENDAMENTO",
    }
)


class DataStoreError(RuntimeError):
    def __init__(
        self,
        message: str,
        *,
        code: str = "DATA_SOURCE_ERROR",
        status: int = 500,
        details: Any = None,
    ) -> None:
        super().__init__(message)
        self.code = code
        self.status = status
        self.details = details


def _deduplicate_existing_directories(candidates: Iterable[Path]) -> list[Path]:
    result: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        try:
            resolved = candidate.expanduser().resolve()
            if not resolved.is_dir():
                continue
        except (OSError, RuntimeError):
            continue
        key = os.path.normcase(str(resolved))
        if key not in seen:
            seen.add(key)
            result.append(resolved)
    return result


def _onedrive_registry_roots() -> list[Path]:
    """Lê os pontos sincronizados registrados pelo cliente OneDrive no Windows."""
    if os.name != "nt":
        return []
    try:
        import winreg
    except ImportError:
        return []

    found: list[Path] = []

    def visit(key_path: str, depth: int = 0) -> None:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                for value_name in ("UserFolder", "MountPoint"):
                    try:
                        value, _kind = winreg.QueryValueEx(key, value_name)
                    except OSError:
                        continue
                    if isinstance(value, str) and value.strip():
                        found.append(Path(value.strip()))
                if depth >= 3:
                    return
                index = 0
                while True:
                    try:
                        child = winreg.EnumKey(key, index)
                    except OSError:
                        break
                    visit(key_path + "\\" + child, depth + 1)
                    index += 1
        except OSError:
            return

    visit(r"Software\Microsoft\OneDrive\Accounts")
    visit(r"Software\SyncEngines\Providers\OneDrive")
    return found


def _onedrive_search_roots(base_dir: Path) -> list[Path]:
    """Obtém roots explícitos ou mounts pessoais/corporativos do OneDrive."""
    configured = os.getenv("EXCEL_SEARCH_ROOTS", "").strip()
    if configured:
        candidates = []
        for raw in configured.split(os.pathsep):
            raw = raw.strip().strip('"')
            if not raw:
                continue
            path = Path(raw)
            candidates.append(path if path.is_absolute() else base_dir / path)
        roots = _deduplicate_existing_directories(candidates)
        if not roots:
            raise DataStoreError(
                "Nenhuma pasta configurada em EXCEL_SEARCH_ROOTS foi encontrada.",
                code="EXCEL_SEARCH_ROOTS_NOT_FOUND",
                status=404,
            )
        return roots

    candidates = [
        Path(value)
        for variable in ("OneDriveCommercial", "OneDriveConsumer", "OneDrive")
        if (value := os.getenv(variable, "").strip())
    ]
    candidates.extend(_onedrive_registry_roots())
    try:
        candidates.extend(path for path in Path.home().glob("OneDrive*") if path.is_dir())
    except OSError:
        pass
    return _deduplicate_existing_directories(candidates)


def discover_excel_file(filename: str, *, base_dir: Path) -> Path:
    """Localiza um único arquivo por nome em OneDrive/SharePoint sincronizado.

    Nunca escolhe silenciosamente entre cópias. Esse cuidado evita que o CRUD
    grave numa versão pessoal, backup ou biblioteca errada quando há duplicatas.
    """
    filename = filename.strip()
    if (
        not filename
        or Path(filename).name != filename
        or filename.startswith("~$")
        or Path(filename).suffix.lower() not in {".xlsx", ".xlsm", ".xlsb"}
    ):
        raise DataStoreError(
            "EXCEL_FILENAME deve conter apenas o nome de um arquivo Excel, sem pastas.",
            code="EXCEL_FILENAME_INVALID",
            status=422,
        )

    roots = _onedrive_search_roots(base_dir)
    if not roots:
        raise DataStoreError(
            "Nenhuma pasta OneDrive/SharePoint sincronizada foi encontrada neste usuário.",
            code="ONEDRIVE_NOT_FOUND",
            status=404,
        )

    target_name = filename.casefold()
    matches: dict[str, Path] = {}
    ignored_directories = {
        ".git", ".pytest_cache", "__pycache__", "node_modules",
        "_rede_backups", "backups",
    }

    for root in roots:
        try:
            for current, directories, files in os.walk(root, topdown=True, followlinks=False):
                directories[:] = [
                    directory for directory in directories
                    if directory.casefold() not in ignored_directories
                ]
                for actual_name in files:
                    if actual_name.casefold() != target_name or actual_name.startswith("~$"):
                        continue
                    candidate = Path(current) / actual_name
                    try:
                        resolved = candidate.resolve()
                        if not resolved.is_file():
                            continue
                    except OSError:
                        continue
                    matches[os.path.normcase(str(resolved))] = resolved
        except (OSError, PermissionError):
            log.warning("Não foi possível percorrer o root sincronizado: %s", root)

    ordered = sorted(matches.values(), key=lambda path: str(path).casefold())
    if len(ordered) == 1:
        log.info("Excel localizado automaticamente em armazenamento sincronizado: %s", ordered[0])
        return ordered[0]
    if len(ordered) > 1:
        raise DataStoreError(
            f"Foram encontradas {len(ordered)} cópias de '{filename}'. Configure EXCEL_SEARCH_ROOTS para indicar a biblioteca correta.",
            code="EXCEL_MULTIPLE_FILES",
            status=409,
            details={"matches": [str(path) for path in ordered]},
        )
    raise DataStoreError(
        f"O arquivo '{filename}' não foi encontrado nas pastas OneDrive/SharePoint sincronizadas.",
        code="EXCEL_NOT_FOUND",
        status=404,
        details={"searchedRoots": [str(path) for path in roots]},
    )


def _display_value(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, (datetime, date)):
        return value.isoformat()[:10]
    text = str(value).strip()
    if text.lower() in {"nan", "nat", "none"}:
        return ""
    # O Excel costuma converter identificadores inteiros em "123.0".
    if re.fullmatch(r"-?\d+\.0", text):
        return text[:-2]
    return text


def _normalized_item(item_id: Any, values: dict[str, Any]) -> dict[str, str]:
    item = {"id": _display_value(item_id)}
    for field in FIELDS:
        value = _display_value(values.get(field, ""))
        if field in DATE_FIELDS and re.match(r"^\d{4}-\d{2}-\d{2}", value):
            value = value[:10]
        item[field] = value
    return item


class QueryMixin:
    """Operações comuns sobre uma coleção normalizada de registros."""

    def _all_items(self) -> list[dict[str, str]]:
        raise NotImplementedError

    @staticmethod
    def _filtered(
        items: Iterable[dict[str, str]],
        filters: dict[str, str] | None = None,
        search: str = "",
    ) -> list[dict[str, str]]:
        filters = filters or {}
        search_lower = search.casefold().strip()
        result: list[dict[str, str]] = []

        for item in items:
            matches = True
            for field in FILTER_FIELDS:
                expected = filters.get(field, "").casefold().strip()
                if expected and expected not in item.get(field, "").casefold():
                    matches = False
                    break

            scheduled = item.get("DATAAGENDAMENTO", "")[:10]
            start = filters.get("dataInicio", "")
            end = filters.get("dataFim", "")
            if matches and start and (not scheduled or scheduled < start):
                matches = False
            if matches and end and (not scheduled or scheduled > end):
                matches = False

            if matches and search_lower:
                haystack = " ".join(item.get(field, "") for field in FIELDS).casefold()
                matches = search_lower in haystack

            if matches:
                result.append(item)

        return result

    @staticmethod
    def _sort(items: list[dict[str, str]], sort: str) -> tuple[list[dict[str, str]], str, str]:
        raw_field, separator, raw_direction = (sort or "DATAAGENDAMENTO:desc").partition(":")
        normalized = re.sub(r"[^a-z0-9]", "", raw_field.casefold())
        field = SORT_ALIASES.get(normalized)
        if not field:
            raise DataStoreError(
                f"Campo de ordenação inválido: {raw_field}",
                code="INVALID_SORT",
                status=422,
            )
        direction = raw_direction.casefold() if separator else "asc"
        if direction not in {"asc", "desc"}:
            raise DataStoreError(
                "A direção de ordenação deve ser asc ou desc.",
                code="INVALID_SORT",
                status=422,
            )

        def key(item: dict[str, str]):
            value = item.get(field, "").casefold()
            # Vazios ficam ao fim na ordenação ascendente.
            return (not bool(value), value)

        return sorted(items, key=key, reverse=direction == "desc"), field, direction

    def list_items(
        self,
        *,
        filters: dict[str, str] | None = None,
        search: str = "",
        page: int = 1,
        page_size: int = 10,
        sort: str = "DATAAGENDAMENTO:desc",
    ) -> dict[str, Any]:
        items = self._filtered(self._all_items(), filters, search)
        items, sort_field, sort_direction = self._sort(items, sort)
        total = len(items)
        pages = max(1, math.ceil(total / page_size))
        safe_page = min(page, pages)
        start = (safe_page - 1) * page_size
        return {
            "items": items[start : start + page_size],
            "pagination": {
                "page": safe_page,
                "pageSize": page_size,
                "total": total,
                "pages": pages,
                "hasPrevious": safe_page > 1,
                "hasNext": safe_page < pages,
            },
            "sort": {"field": sort_field, "direction": sort_direction},
        }

    def dashboard(
        self,
        *,
        filters: dict[str, str] | None = None,
        search: str = "",
    ) -> dict[str, Any]:
        items = self._filtered(self._all_items(), filters, search)

        def count_by(field: str, limit: int | None = None) -> list[tuple[str, int]]:
            counter = Counter(item.get(field, "").strip() or "Não informado" for item in items)
            values = sorted(counter.items(), key=lambda pair: (-pair[1], pair[0].casefold()))
            return values[:limit] if limit else values

        def chart(label: str, pairs: list[tuple[str, int]]) -> dict[str, Any]:
            return {
                "labels": [pair[0] for pair in pairs],
                "datasets": [{"label": label, "data": [pair[1] for pair in pairs]}],
            }

        status_pairs = count_by("STATUS")
        schedule_pairs = sorted(
            Counter(
                item["DATAAGENDAMENTO"][:10]
                for item in items
                if item.get("DATAAGENDAMENTO")
            ).items()
        )
        completed = sum("conclu" in item.get("STATUS", "").casefold() for item in items)
        cancelled = sum("cancel" in item.get("STATUS", "").casefold() for item in items)
        scheduled = sum(bool(item.get("DATAAGENDAMENTO")) for item in items)

        return {
            "kpis": {
                "total": len(items),
                "scheduled": scheduled,
                "completed": completed,
                "pending": max(0, len(items) - completed - cancelled),
            },
            "charts": {
                "status": chart("Registros", status_pairs),
                "clients": chart("Registros", count_by("CLIENTE", 10)),
                "cities": chart("Registros", count_by("CIDADE", 10)),
                "executors": chart("Registros", count_by("EXECUTADOPOR", 10)),
                "clientIds": chart("Registros", count_by("IDCLIENTE", 12)),
                "schedule": chart("Agendamentos", schedule_pairs),
            },
        }

    def filter_options(
        self,
        *,
        filters: dict[str, str] | None = None,
        search: str = "",
    ) -> dict[str, list[str]]:
        # Opções são obtidas do conjunto completo para não desaparecerem ao filtrar.
        items = self._all_items()
        return {
            field: sorted(
                {item.get(field, "").strip() for item in items if item.get(field, "").strip()},
                key=str.casefold,
            )
            for field in ("CLIENTE", "CIDADE", "EXECUTADOPOR", "STATUS")
        }


class ExcelDataStore(QueryMixin):
    """Adaptador compatível com o app; consultas jamais persistem o Excel."""
    mode_label = "Excel local"
    supports_expected_version = True

    def __init__(self, excel_path, sheet_name="Dados", *, key="_ITEM_ID",
                 table=None, header_row=1, backup_dir=None, read_only=False):
        self.path = Path(excel_path).resolve()
        self.sheet_name = sheet_name
        self._safe = SafeExcel(
            self.path, sheet_name, FIELDS, key=key, table=table,
            header_row=header_row, date_fields=DATE_FIELDS,
            backup_dir=backup_dir, read_only=read_only,
        )

    @staticmethod
    def _call(function, *args, **kwargs):
        try:
            return function(*args, **kwargs)
        except ExcelSafetyError as exc:
            raise DataStoreError(str(exc), code=exc.code, status=exc.status) from exc
        except PermissionError as exc:
            raise DataStoreError("Excel bloqueado ou sem permissão.", code="EXCEL_BUSY", status=423) from exc

    def _all_items(self):
        return self._call(self._safe.read_rows)

    def get_item(self, item_id):
        return self._call(self._safe.get, item_id)

    def create_item(self, data):
        return self._call(self._safe.create, data)

    def update_item(self, item_id, data, *, expected_version=None):
        return self._call(self._safe.update, item_id, data, expected_version=expected_version)

    def delete_item(self, item_id, *, expected_version=None):
        return self._call(self._safe.delete, item_id, expected_version=expected_version)

    def fallback_excel_read(self):
        # Pandas continua disponível para análises; não reconstrói o workbook.
        return pd.DataFrame([
            {"_ITEM_ID": item["id"], **{f: item[f] for f in FIELDS}}
            for item in self._all_items()
        ], columns=["_ITEM_ID", *FIELDS])

    def fallback_excel_write(self, frame):
        raise DataStoreError(
            "Gravação integral desativada. Use create_item/update_item/delete_item.",
            code="EXCEL_BULK_WRITE_DISABLED", status=409,
        )

class SharePointClient(QueryMixin):
    """Cliente Microsoft Graph usando o fluxo Client Credentials do MSAL."""

    mode_label = "SharePoint via Microsoft Graph"
    graph_root = "https://graph.microsoft.com/v1.0"

    def __init__(
        self,
        *,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        site_id: str,
        list_id: str,
        field_map: dict[str, str] | None = None,
        timeout: int = 30,
    ) -> None:
        missing = [
            name
            for name, value in {
                "TENANT_ID": tenant_id,
                "CLIENT_ID": client_id,
                "CLIENT_SECRET": client_secret,
                "SHAREPOINT_SITE_ID": site_id,
                "SHAREPOINT_LIST_ID": list_id,
            }.items()
            if not value
        ]
        if missing:
            raise DataStoreError(
                f"Configuração Graph incompleta: {', '.join(missing)}.",
                code="GRAPH_CONFIG_ERROR",
            )

        self.site_id = site_id
        self.list_id = list_id
        self.timeout = timeout
        self.field_map = {field: (field_map or {}).get(field, field) for field in FIELDS}
        self.reverse_field_map = {value: key for key, value in self.field_map.items()}
        self._msal_app = msal.ConfidentialClientApplication(
            client_id=client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret,
        )

        retry = Retry(
            total=3,
            backoff_factor=0.5,
            status_forcelist=(429, 500, 502, 503, 504),
            # POST não é repetido automaticamente para evitar criação duplicada.
            allowed_methods=frozenset({"GET", "PATCH", "DELETE"}),
            respect_retry_after_header=True,
        )
        adapter = HTTPAdapter(max_retries=retry)
        self.session = requests.Session()
        self.session.mount("https://", adapter)

    @property
    def items_url(self) -> str:
        return f"{self.graph_root}/sites/{self.site_id}/lists/{self.list_id}/items"

    def authenticate(self) -> str:
        result = self._msal_app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        token = result.get("access_token")
        if not token:
            message = result.get("error_description") or result.get("error") or "falha desconhecida"
            raise DataStoreError(
                f"Falha de autenticação no Microsoft Graph: {message}",
                code="GRAPH_AUTH_ERROR",
                status=502,
            )
        return token

    def _request(self, method: str, url: str, **kwargs) -> Any:
        headers = kwargs.pop("headers", {})
        headers.update(
            {
                "Authorization": f"Bearer {self.authenticate()}",
                "Accept": "application/json",
                "Content-Type": "application/json",
            }
        )
        try:
            response = self.session.request(
                method,
                url,
                headers=headers,
                timeout=self.timeout,
                **kwargs,
            )
        except requests.RequestException as exc:
            raise DataStoreError(
                "Não foi possível conectar ao Microsoft Graph.",
                code="GRAPH_CONNECTION_ERROR",
                status=502,
                details=str(exc),
            ) from exc

        if response.status_code >= 400:
            try:
                body = response.json()
                graph_error = body.get("error", {})
                message = graph_error.get("message") or response.text
                graph_code = graph_error.get("code", "GRAPH_ERROR")
            except ValueError:
                message = response.text or response.reason
                graph_code = "GRAPH_ERROR"
            status = 404 if response.status_code == 404 else 403 if response.status_code == 403 else 502
            raise DataStoreError(
                f"Microsoft Graph: {message}",
                code=graph_code,
                status=status,
            )

        if response.status_code == 204 or not response.content:
            return None
        return response.json()

    def _to_graph_fields(self, data: dict[str, str]) -> dict[str, str]:
        return {self.field_map[field]: value for field, value in data.items() if field in FIELDS}

    def _from_graph(self, raw: dict[str, Any]) -> dict[str, str]:
        raw_fields = raw.get("fields", {})
        app_fields = {
            self.reverse_field_map.get(name, name): value
            for name, value in raw_fields.items()
        }
        return _normalized_item(raw.get("id", ""), app_fields)

    def _all_items(self) -> list[dict[str, str]]:
        selected = ",".join(self.field_map[field] for field in FIELDS)
        url = self.items_url
        params: dict[str, str] | None = {
            "$expand": f"fields($select={selected})",
            "$top": "999",
        }
        items: list[dict[str, str]] = []
        while url:
            payload = self._request("GET", url, params=params)
            items.extend(self._from_graph(raw) for raw in payload.get("value", []))
            url = payload.get("@odata.nextLink")
            params = None  # nextLink já contém todos os query params e o skip token.
        return items

    def get_item(self, item_id: str) -> dict[str, str]:
        selected = ",".join(self.field_map[field] for field in FIELDS)
        payload = self._request(
            "GET",
            f"{self.items_url}/{item_id}",
            params={"$expand": f"fields($select={selected})"},
        )
        return self._from_graph(payload)

    def create_item(self, data: dict[str, str]) -> dict[str, str]:
        payload = self._request(
            "POST",
            self.items_url,
            json={"fields": self._to_graph_fields(data)},
        )
        return self.get_item(str(payload["id"]))

    def update_item(self, item_id: str, data: dict[str, str]) -> dict[str, str]:
        self._request(
            "PATCH",
            f"{self.items_url}/{item_id}/fields",
            json=self._to_graph_fields(data),
        )
        return self.get_item(item_id)

    def delete_item(self, item_id: str) -> None:
        self._request("DELETE", f"{self.items_url}/{item_id}")


def build_data_store(base_dir: Path | None = None):
    base_dir = base_dir or Path(__file__).resolve().parent
    use_excel = os.getenv("USE_EXCEL_FALLBACK", "true").strip().lower() in {
        "1",
        "true",
        "yes",
        "sim",
        "on",
    }
    if use_excel:
        # Sem .env, a demonstração continua pronta para uso. Quando a variável
        # existe mas está vazia, procura EXCEL_FILENAME nos mounts sincronizados.
        raw_path = os.getenv("EXCEL_PATH")
        if raw_path is None:
            excel_path = base_dir / "data" / "REDE_B2B_EXEMPLO.xlsx"
        elif raw_path.strip():
            excel_path = Path(raw_path.strip()).expanduser()
            if not excel_path.is_absolute():
                excel_path = base_dir / excel_path
        else:
            excel_path = discover_excel_file(
                os.getenv("EXCEL_FILENAME", "REDE_B2B.xlsx"),
                base_dir=base_dir,
            )
        raw_backup = os.getenv("EXCEL_BACKUP_DIR", "").strip()
        backup_dir = Path(raw_backup) if raw_backup else None
        if backup_dir and not backup_dir.is_absolute():
            backup_dir = base_dir / backup_dir
        sheet_name = (
            os.getenv("EXCEL_SHEET", "").strip()
            or os.getenv("EXCEL_SHEET_NAME", "").strip()
            or "Dados"
        )
        return ExcelDataStore(
            excel_path, sheet_name,
            key=os.getenv("EXCEL_KEY_COLUMN", "_ITEM_ID"),
            table=os.getenv("EXCEL_TABLE", "").strip() or None,
            header_row=int(os.getenv("EXCEL_HEADER_ROW", "1")),
            backup_dir=backup_dir,
            read_only=os.getenv("EXCEL_READ_ONLY", "false").lower() == "true",
        )

    field_map_raw = os.getenv("FIELD_MAP_JSON", "").strip()
    try:
        field_map = json.loads(field_map_raw) if field_map_raw else {}
    except json.JSONDecodeError as exc:
        raise DataStoreError(
            "FIELD_MAP_JSON não contém um objeto JSON válido.",
            code="GRAPH_CONFIG_ERROR",
        ) from exc
    if not isinstance(field_map, dict):
        raise DataStoreError(
            "FIELD_MAP_JSON deve ser um objeto JSON.",
            code="GRAPH_CONFIG_ERROR",
        )

    return SharePointClient(
        tenant_id=os.getenv("TENANT_ID", ""),
        client_id=os.getenv("CLIENT_ID", ""),
        client_secret=os.getenv("CLIENT_SECRET", ""),
        site_id=os.getenv("SHAREPOINT_SITE_ID", ""),
        list_id=os.getenv("SHAREPOINT_LIST_ID", ""),
        field_map=field_map,
        timeout=int(os.getenv("GRAPH_TIMEOUT", "30")),
    )


# Atalhos solicitados na especificação; úteis também em scripts externos.
def fallback_excel_read(path: str, sheet_name: str = "Dados") -> pd.DataFrame:
    return ExcelDataStore(path, sheet_name).fallback_excel_read()


def fallback_excel_write(frame: pd.DataFrame, path: str, sheet_name: str = "Dados") -> None:
    ExcelDataStore(path, sheet_name).fallback_excel_write(frame)
