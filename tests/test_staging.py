import json

import pytest

from automation.staging_commands import enqueue


def test_command_is_complete_and_retry_reuses_same_file(tmp_path):
    op = "f19819c0-4f85-47ea-b22f-b84974e0ad3d"
    path = enqueue(tmp_path, op, "1", {"STATUS": "Concluído"})
    assert json.loads(path.read_text(encoding="utf-8"))["entityId"] == "1"
    original = path.read_bytes()
    assert enqueue(tmp_path, op, "1", {"STATUS": "Concluído"}) == path
    assert path.read_bytes() == original
    assert not list(tmp_path.glob("*.tmp"))
    with pytest.raises(ValueError, match="reutilizado"):
        enqueue(tmp_path, op, "1", {"STATUS": "Pendente"})


def test_command_rejects_invalid_identifier(tmp_path):
    with pytest.raises(ValueError):
        enqueue(tmp_path, "../unsafe", "1", {})
