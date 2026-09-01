"""Exemplo executável de PATCH; exige --apply para realmente gravar."""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from excel_safe import SafeExcel
from sharepoint_client import FIELDS, DATE_FIELDS


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", required=True)
    parser.add_argument("--sheet", default="Dados")
    parser.add_argument("--table", default="")
    parser.add_argument("--key", default="_ITEM_ID")
    parser.add_argument("--id", required=True)
    parser.add_argument("--patch-file", required=True)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()
    patch = json.loads(Path(args.patch_file).read_text(encoding="utf-8-sig"))
    adapter = SafeExcel(args.file, args.sheet, FIELDS, key=args.key,
                        table=args.table or None, date_fields=DATE_FIELDS)
    current = adapter.get(args.id)
    if not args.apply:
        print("Simulação: campos a alterar:", ", ".join(patch))
        print("Original não foi modificado. Use --apply após revisar.")
        return
    result = adapter.update(args.id, patch, expected_version=current["_etag"])
    print("Registro atualizado:", result["id"])


if __name__ == "__main__":
    main()
