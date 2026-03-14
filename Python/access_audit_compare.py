from __future__ import annotations

import argparse

from access_audit_core import (
    DEFAULT_IGNORE_COLUMNS,
    invoke_access_audit_comparison,
    write_report_json,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Compara dos bases de datos Access y genera un informe JSON."
    )
    parser.add_argument("--before", required=True, help="Ruta del archivo ACCDB antes.")
    parser.add_argument("--after", required=True, help="Ruta del archivo ACCDB después.")
    parser.add_argument(
        "--output",
        default="audit-report-python.json",
        help="Ruta del informe JSON de salida.",
    )
    parser.add_argument(
        "--ignore-columns",
        nargs="*",
        default=DEFAULT_IGNORE_COLUMNS,
        help="Columnas a ignorar en la comparación.",
    )
    parser.add_argument(
        "--include-unchanged",
        action="store_true",
        help="Incluye tablas sin cambios en el resultado.",
    )
    parser.add_argument(
        "--table",
        action="append",
        default=[],
        help="Limita la comparación a una tabla concreta. Se puede repetir.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    report = invoke_access_audit_comparison(
        before_path=args.before,
        after_path=args.after,
        ignore_columns=args.ignore_columns,
        include_unchanged_tables=args.include_unchanged,
        table_names=args.table,
    )
    output_path = write_report_json(report, args.output)

    changed_tables = report["tables"]
    if not changed_tables:
        print("No se detectaron cambios en las tablas analizadas.")
    else:
        print("Tablas con cambios:")
        for table in changed_tables:
            key_mode = "PK" if table["has_reliable_key"] else "sin PK fiable"
            print(
                f"- {table['table']}: +{table['inserted_count']} / "
                f"-{table['deleted_count']} / ~{table['modified_count']} ({key_mode})"
            )

    print()
    print(f"Informe JSON generado en: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
