from __future__ import annotations

import json
import shutil
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from access_audit_core import (
    DEFAULT_IGNORE_COLUMNS,
    get_access_table_names,
    invoke_access_audit_comparison,
    write_report_json,
)


class AuditApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Auditoria de Access (Python)")
        self.root.geometry("1280x820")

        self.erp_db_var = tk.StringVar()
        self.snapshot_dir_var = tk.StringVar(value=str(Path(".\\copias").resolve()))

        self.before_var = tk.StringVar()
        self.after_var = tk.StringVar()
        self.before_table_var = tk.StringVar(value="(Todas)")
        self.after_table_var = tk.StringVar(value="(Todas)")
        self.ignore_var = tk.StringVar(value=",".join(DEFAULT_IGNORE_COLUMNS))
        self.include_unchanged_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(
            value="Selecciona las bases de datos y pulsa 'Cargar tablas'."
        )

        self.current_result: dict | None = None
        self.current_table_reports: list[dict] = []
        self.current_inserted: list[dict] = []
        self.current_deleted: list[dict] = []
        self.current_modified: list[dict] = []

        # Crear carpeta para contexto de IA si no existe
        Path(".\\contexto_ia").mkdir(exist_ok=True)

        self._build_ui()

    def _build_ui(self) -> None:
        main_container = ttk.Frame(self.root, padding=12)
        main_container.pack(fill="both", expand=True)

        erp_frame = ttk.LabelFrame(main_container, text=" Snapshots Automáticos (Base ERP) ", padding=8)
        erp_frame.pack(fill="x", pady=(0, 15))

        ttk.Label(erp_frame, text="Base de datos ERP:").grid(row=0, column=0, sticky="w")
        ttk.Entry(erp_frame, textvariable=self.erp_db_var, width=80).grid(row=0, column=1, sticky="w", padx=5)
        ttk.Button(erp_frame, text="Seleccionar...", command=self.pick_erp_db).grid(row=0, column=2, padx=5)

        ttk.Label(erp_frame, text="Carpeta copias:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(erp_frame, textvariable=self.snapshot_dir_var, width=80).grid(row=1, column=1, sticky="w", padx=5, pady=5)
        ttk.Button(erp_frame, text="Seleccionar...", command=self.pick_snapshot_dir).grid(row=1, column=2, padx=5, pady=5)

        ttk.Button(erp_frame, text="Snapshot ANTES (A)", command=self.create_snapshot_a).grid(row=0, column=3, rowspan=2, padx=(20, 5), sticky="ns", pady=5)
        ttk.Button(erp_frame, text="Snapshot DESPUÉS (B)", command=self.create_snapshot_b).grid(row=0, column=4, rowspan=2, padx=5, sticky="ns", pady=5)

        container = ttk.Frame(main_container)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="Base ANTES").grid(row=0, column=0, sticky="w")
        ttk.Entry(container, textvariable=self.before_var, width=95).grid(
            row=1, column=0, columnspan=3, sticky="ew", pady=(0, 8)
        )
        ttk.Button(container, text="Seleccionar...", command=self.pick_before).grid(
            row=1, column=3, padx=(8, 0), sticky="ew"
        )

        ttk.Label(container, text="Base DESPUES").grid(row=2, column=0, sticky="w")
        ttk.Entry(container, textvariable=self.after_var, width=95).grid(
            row=3, column=0, columnspan=3, sticky="ew", pady=(0, 8)
        )
        ttk.Button(container, text="Seleccionar...", command=self.pick_after).grid(
            row=3, column=3, padx=(8, 0), sticky="ew"
        )

        ttk.Button(container, text="Cargar tablas", command=self.load_tables).grid(
            row=1, column=4, padx=(12, 0), sticky="ew"
        )
        ttk.Button(
            container, text="Ejecutar auditoria", command=self.run_audit
        ).grid(row=1, column=5, padx=(12, 0), sticky="ew")
        
        self.ai_prompt_button = ttk.Button(
            container, text="Generar Prompt IA", command=self.generate_ai_prompt, state="disabled"
        )
        self.ai_prompt_button.grid(row=3, column=4, padx=(12, 0), sticky="ew")

        self.export_button = ttk.Button(
            container, text="Exportar JSON", command=self.export_json, state="disabled"
        )
        self.export_button.grid(row=3, column=5, padx=(12, 0), sticky="ew")

        ttk.Label(container, text="Tabla ANTES").grid(row=4, column=0, sticky="w")
        self.before_table_combo = ttk.Combobox(
            container,
            textvariable=self.before_table_var,
            values=["(Todas)"],
            state="readonly",
            width=42,
        )
        self.before_table_combo.grid(row=5, column=0, columnspan=2, sticky="ew")

        ttk.Label(container, text="Tabla DESPUES").grid(row=4, column=2, sticky="w")
        self.after_table_combo = ttk.Combobox(
            container,
            textvariable=self.after_table_var,
            values=["(Todas)"],
            state="readonly",
            width=42,
        )
        self.after_table_combo.grid(row=5, column=2, columnspan=2, sticky="ew")

        ttk.Label(container, text="Ignorar columnas").grid(
            row=6, column=0, sticky="w", pady=(12, 0)
        )
        ttk.Entry(container, textvariable=self.ignore_var, width=95).grid(
            row=7, column=0, columnspan=4, sticky="ew"
        )

        ttk.Checkbutton(
            container,
            text="Incluir tablas sin cambios",
            variable=self.include_unchanged_var,
        ).grid(row=8, column=0, columnspan=2, sticky="w", pady=(10, 14))

        ttk.Label(container, text="Resumen de cambios").grid(
            row=9, column=0, sticky="w"
        )
        ttk.Label(container, text="Detalle de la tabla seleccionada").grid(
            row=9, column=4, columnspan=2, sticky="w"
        )

        self.tree = ttk.Treeview(
            container,
            columns=("tabla", "altas", "bajas", "modificados", "clave", "nota"),
            show="headings",
            height=18,
        )
        for key, title, width in (
            ("tabla", "Tabla", 150),
            ("altas", "Altas", 65),
            ("bajas", "Bajas", 65),
            ("modificados", "Modificados", 95),
            ("clave", "Clave", 180),
            ("nota", "Nota", 260),
        ):
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor="w")
        self.tree.grid(row=10, column=0, columnspan=4, sticky="nsew", pady=(8, 0))
        self.tree.bind("<<TreeviewSelect>>", self.on_table_selected)

        detail_frame = ttk.Notebook(container)
        detail_frame.grid(row=10, column=4, columnspan=2, sticky="nsew", pady=(8, 0))

        self.inserted_list, self.inserted_text = self._build_detail_tab(detail_frame, "Altas")
        self.deleted_list, self.deleted_text = self._build_detail_tab(detail_frame, "Bajas")
        self.modified_list, self.modified_text = self._build_detail_tab(
            detail_frame, "Modificados"
        )

        status = ttk.Label(
            container,
            textvariable=self.status_var,
            relief="sunken",
            anchor="w",
            padding=(8, 6),
        )
        status.grid(row=11, column=0, columnspan=6, sticky="ew", pady=(10, 0))

        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)
        container.columnconfigure(2, weight=1)
        container.columnconfigure(3, weight=0)
        container.columnconfigure(4, weight=0)
        container.columnconfigure(5, weight=0)
        container.rowconfigure(10, weight=1)

    def _build_detail_tab(
        self, notebook: ttk.Notebook, title: str
    ) -> tuple[tk.Listbox, tk.Text]:
        frame = ttk.Frame(notebook, padding=6)
        notebook.add(frame, text=title)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(0, weight=1)

        listbox = tk.Listbox(frame, exportselection=False, width=40)
        listbox.grid(row=0, column=0, sticky="nsw")
        text = tk.Text(frame, wrap="word")
        text.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        text.configure(state="disabled")
        return listbox, text

    def pick_erp_db(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("Access database", "*.accdb *.mdb"), ("All files", "*.*")]
        )
        if path:
            self.erp_db_var.set(path)

    def pick_snapshot_dir(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.snapshot_dir_var.set(path)

    def _do_snapshot(self, target_name: str, var_to_update: tk.StringVar, status_msg: str) -> None:
        erp_path = self.erp_db_var.get().strip()
        snap_dir = self.snapshot_dir_var.get().strip()
        if not erp_path:
            messagebox.showerror("Error", "Selecciona primero la Base de datos ERP.")
            return
        if not snap_dir:
            messagebox.showerror("Error", "Selecciona una carpeta para las copias.")
            return
        
        try:
            Path(snap_dir).mkdir(parents=True, exist_ok=True)
            target_path = Path(snap_dir) / target_name
            shutil.copy2(erp_path, target_path)
            var_to_update.set(str(target_path.resolve()))
            self.status_var.set(status_msg)
        except Exception as e:
            messagebox.showerror("Error copiando", str(e))

    def create_snapshot_a(self) -> None:
        self._do_snapshot("A.accdb", self.before_var, "Snapshot ANTES (A) creado y asignado.")

    def create_snapshot_b(self) -> None:
        self._do_snapshot("B.accdb", self.after_var, "Snapshot DESPUÉS (B) creado y asignado.")

    def generate_ai_prompt(self) -> None:
        if not self.current_result:
            return

        prompt = [
            "Actúa como un experto contable y auditor de sistemas. Analiza los siguientes cambios en la base de datos de un programa contable tras una acción del usuario.",
            "Tu objetivo es deducir qué operación administrativa o contable se ha realizado (ej. Facturación, Contabilización, Cobro, etc.).",
            "",
            "### RESUMEN DE CAMBIOS DETECTADOS ###",
            ""
        ]

        for report in self.current_table_reports:
            table_name = report["table"]
            prompt.append(f"--- TABLA: {table_name} ---")
            prompt.append(f"  Altas: {report['inserted_count']}")
            prompt.append(f"  Bajas: {report['deleted_count']}")
            prompt.append(f"  Modificaciones: {report['modified_count']}")

            if report["inserted_count"] > 0 and report.get("inserted"):
                prompt.append("  [Detalle Altas (Primeros registros)]:")
                for i, entry in enumerate(report["inserted"][:2]):
                    prompt.append(f"    - Nueva Fila: {entry['key']}")
                    for k, v in list(entry.get("row", {}).items())[:10]:
                        prompt.append(f"        {k} = {v}")

            if report["modified_count"] > 0 and report.get("modified"):
                prompt.append("  [Detalle Modificaciones (Primeros registros)]:")
                for i, entry in enumerate(report["modified"][:3]):
                    prompt.append(f"    - Fila {entry['key']}:")
                    for change in entry.get("changes", []):
                        prompt.append(f"        Campo '{change['column']}': Antes='{change['before']}' -> Después='{change['after']}'")
            prompt.append("")

        prompt.extend([
            "### INSTRUCCIONES PARA EL ANÁLISIS ###",
            "1. Identifica qué tablas principales sufrieron cambios (ej. Cabeceras de facturas vs Líneas).",
            "2. Relaciona los campos modificados (ej. cambio de estado de pendiente a cobrado o totales).",
            "3. Concluye explicando qué ha hecho el usuario en el programa.",
            "4. Utiliza la documentación del manual del programa que tienes en tu contexto para dar nombres exactos si es posible."
        ])

        full_prompt = "\n".join(prompt)

        # Mostrar en ventana
        popup = tk.Toplevel(self.root)
        popup.title("Prompt para NotebookLM / IA")
        popup.geometry("800x600")
        
        text_area = tk.Text(popup, wrap="word", font=("Consolas", 10))
        text_area.pack(fill="both", expand=True, padx=10, pady=10)
        text_area.insert("1.0", full_prompt)
        text_area.configure(state="disabled")

        def copy_to_clipboard():
            self.root.clipboard_clear()
            self.root.clipboard_append(full_prompt)
            messagebox.showinfo("Copiado", "Prompt copiado al portapapeles. Pégalo en NotebookLM.", parent=popup)

        ttk.Button(popup, text="Copiar al Portapapeles", command=copy_to_clipboard).pack(pady=10)

    def pick_before(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("Access database", "*.accdb *.mdb"), ("All files", "*.*")]
        )
        if path:
            self.before_var.set(path)

    def pick_after(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("Access database", "*.accdb *.mdb"), ("All files", "*.*")]
        )
        if path:
            self.after_var.set(path)

    def load_tables(self) -> None:
        try:
            if not self.before_var.get().strip() or not self.after_var.get().strip():
                raise ValueError("Debes seleccionar la base ANTES y la base DESPUES.")

            before_tables = get_access_table_names(self.before_var.get())
            after_tables = get_access_table_names(self.after_var.get())
            self.before_table_combo["values"] = ["(Todas)", *before_tables]
            self.after_table_combo["values"] = ["(Todas)", *after_tables]
            self.before_table_var.set("(Todas)")
            self.after_table_var.set("(Todas)")
            self.status_var.set("Tablas cargadas. Ya puedes ejecutar la auditoria.")
        except Exception as exc:
            messagebox.showerror("Auditoria Access (Python)", str(exc))
            self.status_var.set("Error cargando tablas.")

    def get_selected_tables(self) -> list[str]:
        selected: list[str] = []
        if self.before_table_var.get() and self.before_table_var.get() != "(Todas)":
            selected.append(self.before_table_var.get())
        if (
            self.after_table_var.get()
            and self.after_table_var.get() != "(Todas)"
            and self.after_table_var.get() not in selected
        ):
            selected.append(self.after_table_var.get())
        return selected

    def run_audit(self) -> None:
        try:
            self.status_var.set("Ejecutando auditoria...")
            self.root.update_idletasks()

            ignore_columns = [
                value.strip()
                for value in self.ignore_var.get().split(",")
                if value.strip()
            ]

            self.current_result = invoke_access_audit_comparison(
                before_path=self.before_var.get(),
                after_path=self.after_var.get(),
                ignore_columns=ignore_columns,
                include_unchanged_tables=self.include_unchanged_var.get(),
                table_names=self.get_selected_tables(),
            )

            self.current_table_reports = list(self.current_result["tables"])
            self.bind_table_grid()
            self.reset_detail_views()
            self.export_button.configure(state="normal")
            self.ai_prompt_button.configure(state="normal")

            if not self.current_table_reports:
                self.status_var.set("No se detectaron tablas con cambios.")
            else:
                self.status_var.set(
                    "Auditoria completada. Selecciona una tabla para ver detalles."
                )
        except Exception as exc:
            messagebox.showerror("Auditoria Access (Python)", str(exc))
            self.status_var.set("La auditoria ha fallado.")

    def bind_table_grid(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        for report in self.current_table_reports:
            key_label = (
                ", ".join(report["primary_key_columns"])
                if report["has_reliable_key"]
                else "Sin PK fiable"
            )
            note = report.get("note", "")
            self.tree.insert(
                "",
                "end",
                values=(
                    report["table"],
                    report["inserted_count"],
                    report["deleted_count"],
                    report["modified_count"],
                    key_label,
                    note,
                ),
            )

    def reset_detail_views(self) -> None:
        self.current_inserted = []
        self.current_deleted = []
        self.current_modified = []
        for listbox in (self.inserted_list, self.deleted_list, self.modified_list):
            listbox.delete(0, "end")
        for text in (self.inserted_text, self.deleted_text, self.modified_text):
            text.configure(state="normal")
            text.delete("1.0", "end")
            text.configure(state="disabled")

    def on_table_selected(self, _event: object) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        item = self.tree.item(selection[0])
        table_name = item["values"][0]
        report = next(
            (candidate for candidate in self.current_table_reports if candidate["table"] == table_name),
            None,
        )
        if report is None:
            return

        self.reset_detail_views()
        self.current_inserted = list(report["inserted"])
        self.current_deleted = list(report["deleted"])
        self.current_modified = list(report["modified"])

        self._bind_detail_list(self.inserted_list, self.current_inserted, "inserted")
        self._bind_detail_list(self.deleted_list, self.current_deleted, "deleted")
        self._bind_detail_list(self.modified_list, self.current_modified, "modified")

    def _bind_detail_list(
        self, listbox: tk.Listbox, entries: list[dict], mode: str
    ) -> None:
        for entry in entries:
            listbox.insert("end", entry["key"])
        if entries:
            listbox.selection_set(0)
            self.show_entry_details(entries[0], mode)

        callback = lambda event, m=mode: self._on_detail_selected(event, m)
        listbox.bind("<<ListboxSelect>>", callback)

    def _on_detail_selected(self, event: tk.Event, mode: str) -> None:
        widget = event.widget
        selection = widget.curselection()
        if not selection:
            return
        index = selection[0]
        source = {
            "inserted": self.current_inserted,
            "deleted": self.current_deleted,
            "modified": self.current_modified,
        }[mode]
        self.show_entry_details(source[index], mode)

    def show_entry_details(self, entry: dict, mode: str) -> None:
        if mode == "inserted":
            target = self.inserted_text
        elif mode == "deleted":
            target = self.deleted_text
        else:
            target = self.modified_text

        lines = [f"Clave: {entry['key']}", ""]
        if mode == "modified":
            for change in entry["changes"]:
                lines.extend(
                    [
                        f"Campo: {change['column']}",
                        f"  Antes: {change['before']}",
                        f"  Despues: {change['after']}",
                        "",
                    ]
                )
        else:
            for key, value in entry["row"].items():
                lines.append(f"{key}: {value}")

        target.configure(state="normal")
        target.delete("1.0", "end")
        target.insert("1.0", "\n".join(lines).rstrip())
        target.configure(state="disabled")

    def export_json(self) -> None:
        if not self.current_result:
            messagebox.showerror(
                "Auditoria Access (Python)", "Todavia no hay una auditoria ejecutada."
            )
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            initialfile="audit-report-python.json",
            filetypes=[("JSON", "*.json")],
        )
        if not output_path:
            return

        resolved = write_report_json(self.current_result, output_path)
        messagebox.showinfo(
            "Auditoria Access (Python)", f"Informe exportado a:\n{resolved}"
        )
        self.status_var.set("Informe exportado correctamente.")


def main() -> int:
    root = tk.Tk()
    app = AuditApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
