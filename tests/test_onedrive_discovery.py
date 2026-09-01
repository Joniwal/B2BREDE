from __future__ import annotations

import os
from pathlib import Path

import pytest

import sharepoint_client
from sharepoint_client import DataStoreError, build_data_store, discover_excel_file


def _file(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"xlsx-placeholder-for-path-discovery")
    return path


def test_blank_excel_path_discovers_unique_synced_file(tmp_path, monkeypatch):
    root = tmp_path / "OneDrive - Empresa"
    expected = _file(root / "Biblioteca Compartilhada" / "REDE_B2B.xlsx")
    monkeypatch.setenv("EXCEL_PATH", "")
    monkeypatch.setenv("EXCEL_FILENAME", "REDE_B2B.xlsx")
    monkeypatch.setenv("EXCEL_SEARCH_ROOTS", str(root))
    monkeypatch.delenv("EXCEL_SHEET", raising=False)
    monkeypatch.setenv("EXCEL_SHEET_NAME", "REDEB2B")
    monkeypatch.setenv("EXCEL_KEY_COLUMN", "IDCLIENTE")

    repository = build_data_store(tmp_path / "app")

    assert repository.path == expected.resolve()
    assert repository.sheet_name == "REDEB2B"
    assert repository._safe.key == "IDCLIENTE"


def test_discovery_uses_onedrive_environment_mount(tmp_path, monkeypatch):
    home = tmp_path / "home"
    root = tmp_path / "OneDrive - Empresa"
    expected = _file(root / "Compartilhado" / "REDE_B2B.xlsx")
    home.mkdir()
    monkeypatch.delenv("EXCEL_SEARCH_ROOTS", raising=False)
    monkeypatch.setenv("OneDriveCommercial", str(root))
    monkeypatch.delenv("OneDriveConsumer", raising=False)
    monkeypatch.delenv("OneDrive", raising=False)
    monkeypatch.setattr(sharepoint_client, "_onedrive_registry_roots", lambda: [])
    monkeypatch.setattr(sharepoint_client.Path, "home", lambda: home)

    assert discover_excel_file("REDE_B2B.xlsx", base_dir=tmp_path) == expected.resolve()


def test_discovery_refuses_duplicate_file_names(tmp_path, monkeypatch):
    first = tmp_path / "OneDrive A"
    second = tmp_path / "OneDrive B"
    _file(first / "Equipe" / "REDE_B2B.xlsx")
    _file(second / "Operação" / "REDE_B2B.xlsx")
    monkeypatch.setenv("EXCEL_SEARCH_ROOTS", os.pathsep.join((str(first), str(second))))

    with pytest.raises(DataStoreError) as caught:
        discover_excel_file("REDE_B2B.xlsx", base_dir=tmp_path)

    assert caught.value.code == "EXCEL_MULTIPLE_FILES"
    assert len(caught.value.details["matches"]) == 2


def test_discovery_reports_missing_file_and_searched_root(tmp_path, monkeypatch):
    root = tmp_path / "OneDrive vazio"
    root.mkdir()
    monkeypatch.setenv("EXCEL_SEARCH_ROOTS", str(root))

    with pytest.raises(DataStoreError) as caught:
        discover_excel_file("REDE_B2B.xlsx", base_dir=tmp_path)

    assert caught.value.code == "EXCEL_NOT_FOUND"
    assert caught.value.details["searchedRoots"] == [str(root.resolve())]


@pytest.mark.parametrize("filename", ("", "../REDE_B2B.xlsx", "pasta/REDE_B2B.xlsx", "arquivo.csv", "~$REDE_B2B.xlsx"))
def test_discovery_validates_filename(filename, tmp_path, monkeypatch):
    root = tmp_path / "OneDrive"
    root.mkdir()
    monkeypatch.setenv("EXCEL_SEARCH_ROOTS", str(root))

    with pytest.raises(DataStoreError) as caught:
        discover_excel_file(filename, base_dir=tmp_path)

    assert caught.value.code == "EXCEL_FILENAME_INVALID"


def test_explicit_excel_path_still_has_priority(tmp_path, monkeypatch):
    explicit = _file(tmp_path / "manual" / "arquivo.xlsx")
    monkeypatch.setenv("EXCEL_PATH", str(explicit))
    monkeypatch.setenv("EXCEL_SEARCH_ROOTS", str(tmp_path / "inexistente"))

    repository = build_data_store(tmp_path / "app")

    assert repository.path == explicit.resolve()


def test_missing_excel_path_variable_keeps_demo_default(tmp_path, monkeypatch):
    demo = _file(tmp_path / "data" / "REDE_B2B_EXEMPLO.xlsx")
    monkeypatch.delenv("EXCEL_PATH", raising=False)

    repository = build_data_store(tmp_path)

    assert repository.path == demo.resolve()
