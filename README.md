# Auditor de cambios para bases de datos Access (`ACCDB`)

Este proyecto compara dos copias de una base de datos Access:

- estado **antes** de ejecutar un proceso,
- estado **despues** de ejecutarlo.

La solucion tiene dos modos:

- modo comando con `PowerShell\access_audit_compare.ps1`,
- modo interfaz con `PowerShell\access_audit_gui.ps1`.

Tambien tienes una version en Python:

- modo comando con `Python\access_audit_compare.py`,
- modo interfaz con `Python\access_audit_gui.py`.

El comparador detecta:

- tablas que han cambiado,
- registros insertados,
- registros borrados,
- registros modificados,
- campos concretos que cambiaron en cada registro modificado.

## Requisitos

- Windows
- Driver ODBC de Microsoft Access instalado
- PowerShell 5+ o PowerShell 7+

En este equipo ya aparece instalado `Microsoft Access Driver (*.mdb, *.accdb)`.

## Uso con interfaz

La forma mas comoda es abrir la interfaz:

```powershell
powershell -ExecutionPolicy Bypass -File .\PowerShell\access_audit_gui.ps1
```

Desde esa ventana puedes:

- elegir la base `ANTES`,
- elegir la base `DESPUES`,
- cargar las tablas de cada base,
- seleccionar una tabla concreta o dejar `(Todas)`,
- ejecutar la auditoria,
- ver el resumen por tabla,
- abrir el detalle de altas, bajas y modificados,
- exportar el resultado a JSON.

## Uso con interfaz Python

Si prefieres Python:

```powershell
python .\Python\access_audit_gui.py
```

## Uso por linea de comandos en Python

```powershell
python .\Python\access_audit_compare.py `
  --before "C:\ruta\empresa_antes.accdb" `
  --after "C:\ruta\empresa_despues.accdb" `
  --output ".\salida\informe-python.json"
```

## Uso basico por linea de comandos

```powershell
powershell -ExecutionPolicy Bypass -File .\PowerShell\access_audit_compare.ps1 `
  -BeforePath "C:\ruta\empresa_antes.accdb" `
  -AfterPath "C:\ruta\empresa_despues.accdb" `
  -OutputPath ".\salida\informe.json"
```

## Ignorar columnas de ruido

Si hay campos que cambian siempre, puedes ignorarlos:

```powershell
powershell -ExecutionPolicy Bypass -File .\PowerShell\access_audit_compare.ps1 `
  -BeforePath "C:\ruta\antes.accdb" `
  -AfterPath "C:\ruta\despues.accdb" `
  -OutputPath ".\salida\informe.json" `
  -IgnoreColumns FechaModificacion,UsuarioModificacion,ContadorInterno
```

## Salida

El script genera un JSON con esta estructura:

```json
{
  "generated_at": "2026-03-14T15:00:00.0000000+01:00",
  "before_database": "C:\\ruta\\antes.accdb",
  "after_database": "C:\\ruta\\despues.accdb",
  "ignored_columns": ["FechaModificacion"],
  "table_count": 2,
  "tables": [
    {
      "table": "FACTURAS",
      "primary_key_columns": ["IDFACTURA"],
      "has_reliable_key": true,
      "inserted_count": 1,
      "deleted_count": 0,
      "modified_count": 2,
      "modified": [
        {
          "key": "IDFACTURA=1052",
          "changes": [
            {
              "column": "TOTAL",
              "before": "1250",
              "after": "1475"
            }
          ]
        }
      ]
    }
  ]
}
```

## Limitaciones

- Si una tabla no tiene clave primaria detectable, el script puede listar altas y bajas, pero la deteccion de modificaciones no es tan fiable.
- Si el programa contable cambia la estructura de la base entre el antes y el despues, el informe lo reflejara a nivel de tabla, pero no hace una auditoria completa del esquema.
- Los campos binarios o complejos se serializan como texto/base64 para poder compararlos.

## Siguiente mejora recomendada

La siguiente evolucion natural seria anadir:

- filtro por rango de claves o fechas,
- exportacion a Excel,
- vista previa de solo tablas con cambios,
- perfiles guardados para repetir auditorias.
