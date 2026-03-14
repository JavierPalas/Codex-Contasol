from __future__ import annotations

import base64
import json
import os
import site
from collections import OrderedDict
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable


def _ensure_user_site_on_path() -> None:
    user_site = site.getusersitepackages()
    if user_site and user_site not in os.sys.path:
        os.sys.path.append(user_site)


_ensure_user_site_on_path()

import pyodbc  # noqa: E402


DEFAULT_IGNORE_COLUMNS = [
    "FechaModificacion",
    "Fecha_Modificacion",
    "FecMod",
    "UsuarioModificacion",
    "Usuario_Modificacion",
]


@dataclass
class TableReport:
    table: str
    primary_key_columns: list[str]
    tracked_columns: list[str]
    has_reliable_key: bool
    inserted_count: int
    deleted_count: int
    modified_count: int
    inserted: list[dict[str, Any]]
    deleted: list[dict[str, Any]]
    modified: list[dict[str, Any]]
    note: str | None = None

    def to_dict(self) -> dict[str, Any]:
        result = {
            "table": self.table,
            "primary_key_columns": self.primary_key_columns,
            "tracked_columns": self.tracked_columns,
            "has_reliable_key": self.has_reliable_key,
            "inserted_count": self.inserted_count,
            "deleted_count": self.deleted_count,
            "modified_count": self.modified_count,
            "inserted": self.inserted,
            "deleted": self.deleted,
            "modified": self.modified,
        }
        if self.note is not None:
            result["note"] = self.note
        return result


def resolve_required_path(path: str) -> str:
    resolved = Path(path).expanduser().resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"No existe el archivo: {path}")
    return str(resolved)


def connect_access(database_path: str) -> pyodbc.Connection:
    return pyodbc.connect(
        rf"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={database_path};"
    )


def get_user_tables(connection: pyodbc.Connection) -> list[str]:
    cursor = connection.cursor()
    tables = [
        row.table_name
        for row in cursor.tables(tableType="TABLE")
        if not row.table_name.startswith("MSys") and not row.table_name.startswith("~")
    ]
    return sorted(set(tables))


def get_access_table_names(database_path: str) -> list[str]:
    resolved_path = resolve_required_path(database_path)
    with connect_access(resolved_path) as connection:
        return get_user_tables(connection)


def get_column_names(connection: pyodbc.Connection, table_name: str) -> list[str]:
    cursor = connection.cursor()
    cursor.execute(f"SELECT * FROM [{table_name}] WHERE 1 = 0")
    return [column[0] for column in cursor.description]


def get_primary_key_columns(connection: pyodbc.Connection, table_name: str) -> list[str]:
    cursor = connection.cursor()
    try:
        rows = list(cursor.statistics(table_name, unique=True))
    except pyodbc.Error:
        return []

    primary_rows = [
        row
        for row in rows
        if getattr(row, "index_name", None) == "PrimaryKey"
        and getattr(row, "column_name", None)
    ]
    if primary_rows:
        primary_rows.sort(key=lambda row: getattr(row, "ordinal_position", 0) or 0)
        return [row.column_name for row in primary_rows]

    unique_index_name = None
    for row in rows:
        index_name = getattr(row, "index_name", None)
        column_name = getattr(row, "column_name", None)
        if index_name and column_name:
            unique_index_name = index_name
            break

    if not unique_index_name:
        return []

    unique_rows = [
        row
        for row in rows
        if getattr(row, "index_name", None) == unique_index_name
        and getattr(row, "column_name", None)
    ]
    unique_rows.sort(key=lambda row: getattr(row, "ordinal_position", 0) or 0)
    return [row.column_name for row in unique_rows]


def convert_field_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return datetime.combine(value, time()).isoformat()
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, bytes):
        return base64.b64encode(value).decode("ascii")
    return str(value)


def get_row_map(
    connection: pyodbc.Connection,
    table_name: str,
    columns: list[str],
    key_columns: list[str],
) -> dict[str, OrderedDict[str, Any]]:
    cursor = connection.cursor()
    cursor.execute(f"SELECT * FROM [{table_name}]")

    rows: dict[str, OrderedDict[str, Any]] = {}
    for row_number, record in enumerate(cursor.fetchall(), start=1):
        row = OrderedDict()
        for column in columns:
            row[column] = convert_field_value(getattr(record, column))

        if key_columns:
            key = "|".join(f"{column}={row[column]}" for column in key_columns)
        else:
            fingerprint = "|".join(f"{column}={row[column]}" for column in columns)
            key = f"ROW:{row_number}:{fingerprint}"
        rows[key] = row

    return rows


def compare_table_content(
    before_connection: pyodbc.Connection,
    after_connection: pyodbc.Connection,
    table_name: str,
    ignore_columns: Iterable[str] | None = None,
) -> TableReport:
    ignore_set = {column for column in (ignore_columns or [])}
    before_columns = get_column_names(before_connection, table_name)
    after_columns = get_column_names(after_connection, table_name)
    common_columns = [column for column in before_columns if column in after_columns]
    tracked_columns = [column for column in common_columns if column not in ignore_set]
    primary_key_columns = get_primary_key_columns(after_connection, table_name)

    before_rows = get_row_map(
        before_connection, table_name, tracked_columns, primary_key_columns
    )
    after_rows = get_row_map(
        after_connection, table_name, tracked_columns, primary_key_columns
    )

    all_keys = sorted(set(before_rows) | set(after_rows))
    inserted: list[dict[str, Any]] = []
    deleted: list[dict[str, Any]] = []
    modified: list[dict[str, Any]] = []

    for key in all_keys:
        in_before = key in before_rows
        in_after = key in after_rows

        if in_before and not in_after:
            deleted.append({"key": key, "row": before_rows[key]})
            continue
        if in_after and not in_before:
            inserted.append({"key": key, "row": after_rows[key]})
            continue

        before_row = before_rows[key]
        after_row = after_rows[key]
        changes = []
        for column in tracked_columns:
            if before_row[column] != after_row[column]:
                changes.append(
                    {
                        "column": column,
                        "before": before_row[column],
                        "after": after_row[column],
                    }
                )
        if changes:
            modified.append({"key": key, "changes": changes})

    return TableReport(
        table=table_name,
        primary_key_columns=primary_key_columns,
        tracked_columns=tracked_columns,
        has_reliable_key=bool(primary_key_columns),
        inserted_count=len(inserted),
        deleted_count=len(deleted),
        modified_count=len(modified),
        inserted=inserted,
        deleted=deleted,
        modified=modified,
    )


def invoke_access_audit_comparison(
    before_path: str,
    after_path: str,
    ignore_columns: Iterable[str] | None = None,
    include_unchanged_tables: bool = False,
    table_names: Iterable[str] | None = None,
) -> dict[str, Any]:
    resolved_before = resolve_required_path(before_path)
    resolved_after = resolve_required_path(after_path)
    ignore_columns = [str(column) for column in (ignore_columns or [])]
    selected_tables_filter = {str(name) for name in (table_names or [])}

    with connect_access(resolved_before) as before_connection, connect_access(
        resolved_after
    ) as after_connection:
        before_tables = get_user_tables(before_connection)
        after_tables = get_user_tables(after_connection)
        all_tables = sorted(set(before_tables) | set(after_tables))

        if selected_tables_filter:
            tables_to_compare = [
                table_name
                for table_name in all_tables
                if table_name in selected_tables_filter
            ]
        else:
            tables_to_compare = all_tables

        reports: list[TableReport] = []
        for table_name in tables_to_compare:
            if table_name in before_tables and table_name in after_tables:
                report = compare_table_content(
                    before_connection,
                    after_connection,
                    table_name,
                    ignore_columns=ignore_columns,
                )
            elif table_name in after_tables:
                report = TableReport(
                    table=table_name,
                    primary_key_columns=[],
                    tracked_columns=[],
                    has_reliable_key=False,
                    inserted_count=-1,
                    deleted_count=0,
                    modified_count=0,
                    inserted=[],
                    deleted=[],
                    modified=[],
                    note="La tabla solo existe en la base de datos DESPUES.",
                )
            else:
                report = TableReport(
                    table=table_name,
                    primary_key_columns=[],
                    tracked_columns=[],
                    has_reliable_key=False,
                    inserted_count=0,
                    deleted_count=-1,
                    modified_count=0,
                    inserted=[],
                    deleted=[],
                    modified=[],
                    note="La tabla solo existe en la base de datos ANTES.",
                )

            has_changes = (
                report.note is not None
                or report.inserted_count > 0
                or report.deleted_count > 0
                or report.modified_count > 0
            )
            if include_unchanged_tables or has_changes:
                reports.append(report)

    return {
        "generated_at": datetime.now().astimezone().isoformat(),
        "before_database": resolved_before,
        "after_database": resolved_after,
        "ignored_columns": ignore_columns,
        "table_count": len(reports),
        "tables": [report.to_dict() for report in reports],
    }


def write_report_json(report: dict[str, Any], output_path: str) -> str:
    resolved_output = Path(output_path).expanduser().resolve()
    resolved_output.parent.mkdir(parents=True, exist_ok=True)
    resolved_output.write_text(
        json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    return str(resolved_output)
