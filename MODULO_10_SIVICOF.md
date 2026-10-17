# Módulo 10 — Reportes SIVICOF

Sistema de Gestión Contractual · https://richbello.github.io/sistema-gestion-contractual/
Repositorio: richbello/sistema-gestion-contractual · commits d96043c y 6e9ea79 (rama main)

## 1. Qué hace

Genera el paquete mensual que Presupuesto del FDL Usme envía a la Contraloría de Bogotá
(SIVICOF) **a más tardar el día 9 del mes siguiente al corte**:

```
SIVICOF <MES>.zip
└── SIVICOF <MES>/
    ├── 01_000000005_AAAAMMDD.xlsx   Formulario 400 · CB-0104 (reservas / obligaciones por pagar)
    ├── CBN-0001_<Mes>_<AAAA>.pdf     No existencia de recursos de tesorería
    ├── CBN-1093_<Mes>_<AAAA>.pdf     Modificaciones presupuestales
    ├── CBN-1098_<Mes>_<AAAA>.pdf     Patrimonios autónomos / encargos fiduciarios
    ├── CBN-1109_<Mes>_<AAAA>.pdf     Egresos de patrimonios autónomos
    └── <Decreto>.pdf                 solo si hubo modificación presupuestal
```

Ejemplo: corte 30/09/2026 → `01_000000005_20260930.xlsx`, `CBN-1093_Sept_2026.pdf`,
`SIVICOF SEPTIEMBRE.zip`.

## 2. Entradas

| Entrada | Origen | Para qué |
|---|---|---|
| Mes a reportar | Pantalla | Fecha de corte = último día del mes (C5 y nombre del Excel) |
| Fecha de firma | Pantalla (por defecto el día 9 del mes siguiente) | Texto "Dado en Bogotá D.C. a los …" de los 4 certificados |
| ¿Hubo modificación? + No. y fecha del decreto + PDF | Pantalla | Texto del CBN-1093 y anexo del decreto |
| Reporte de Ejecución de Gastos | SAP / BogData, corte último día del mes | E, H, J, K y controles |
| Reporte de Anulaciones de CRP | SAP / BogData | F y G |

## 3. Lógica del CB-0104

Solo se diligencian **C5** y las columnas **E, F, G, H, J, K** de las filas:

| Fila | Rubro (columna CODIGO) |
|---|---|
| 11 | O219001 Obligaciones por Pagar Funcionamiento Vigencia Anterior |
| 12 | O219002 Obligaciones por Pagar Funcionamiento Otras Vigencias |
| 19 | O230689 Obligaciones por pagar Inversión vigencia anterior |
| 20 | O230690 Obligaciones por pagar Inversión vigencias anteriores |

Las columnas **I** (% participación), **L** (% ejecución giro) y **M** (saldo, en verde) **no se
diligencian**: son fórmulas del formato (`=+H/E`, `=K/H`, `=H-K`). El sistema solo recalcula su
valor guardado para que el archivo se vea bien aunque se abra en un visor que no recalcula.

| Celda | Fuente | Regla |
|---|---|---|
| C5 | Mes elegido | Último día del mes, formato aaaa/mm/dd |
| E Reservas constituidas | Ejecución | Columna "Apropiación Vigente // RESERVAS CONSTITUIDAS" del rubro |
| H Reservas definitivas | Ejecución | Columna "CDP Acumulado // RESERVAS DEFINITIVAS" |
| J Autorización giro del mes | Ejecución | Columna "Giro Mes Presupuestal // AUTORIZACION DE GIRO DEL MES" |
| K Autorización giro acumulada | Ejecución | Columna "Giros Acumulados Ppto // AUTORIZACION DE GIRO ACUMULADA" |
| **F Anulaciones del mes** | Anulaciones CRP | Σ (Liberación total + Liberación parcial) del rubro con **Fecha Liberación dentro del mes del corte** |
| **G Anulaciones acumuladas** | Anulaciones CRP | Σ (Liberación total + Liberación parcial) del rubro con **Fecha Liberación en los meses anteriores** del mismo año (enero → mes anterior), **sin importar el motivo** ni si tiene acta |

En el reporte de Ejecución el rubro se busca por el primer texto de la columna A
(`O230689   Obligaciones por pagar…` → `O230689`). Las columnas se ubican por el **nombre**
del encabezado, no por posición. En Anulaciones se usan las columnas `Rubro`,
`Liberación total`, `Liberación parcial` y `Fecha Liberación`; `Importe Original` no se usa.

Ejemplo O230689 (reporte de anulaciones a 30/09/2026):

| Corte | F (mes) | G (meses anteriores) |
|---|---|---|
| Junio | 41.170.000 | 811.056.567 (ene + feb + abr + may) |
| Julio | 4.490.001 | 852.226.567 (ene → jun) |
| Septiembre | 0 | 856.716.568 (ene → ago) |

## 4. Controles (avisos, no bloquean)

Se muestran en pantalla antes de generar:

1. **F vs CDP Mes de la Ejecución** (con signo contrario). Marca ✓ o "≠ SAP".
2. **F + G vs "Saldo Apr.Disponible"** de la Ejecución (= E − H). Marca ✓ o "≠ SAP" y explica
   la diferencia. Si la diferencia es igual al "Saldo p. Comprometer", indica que es un CRP
   anulado cuyo saldo quedó en el CDP (no rebaja la reserva definitiva).

Con septiembre 2026 aparecen dos avisos: O219001 (713.692) y O230689 (732.651.444, que son las
correcciones de enero y febrero "ERROR INVOLUNTARIO EN CARGUE").

## 5. Certificados CBN

Diseño igual al Word oficial (logo Bogotá–Secretaría de Gobierno, firma, "Proyectó / Revisó",
pie de la Alcaldía Local, código GDI-GPD-F054 y escudo), generados con reportlab y fuente
EB Garamond (licencia libre OFL, incluida).

- Fecha (los 4): `Dado en Bogotá D.C. a los nueve (09) días del mes de octubre de dos mil veintiséis (2026).`
  Día 1: "al primer (01) día…".
- CBN-1093 con modificación: `…del Fondo de Desarrollo Local de Usme se realizó modificación
  presupuestal mediante Decreto Local No. 11 de septiembre de 2026.`
  Si se da la fecha del decreto: `…No. 04 del 06 de mayo de 2026.`
- CBN-1093 sin modificación: `…no se realizaron modificaciones presupuestales de ingresos y
  gastos durante el mes de octubre de 2026.`

## 6. Correo

La pantalla deja listos Para (jorgea.gonzalez@gobiernobogota.gov.co), Asunto
(`Reporte SIVICOF  Septiembre 2026_FDLU`) y Mensaje, con botones Copiar y un enlace "Abrir
borrador de correo". El ZIP se adjunta a mano (los navegadores no permiten adjuntar por mailto).

## 7. Arquitectura y archivos

| Archivo | Rol |
|---|---|
| `backend/app/services/sivicof_service.py` | Lectura de reportes, cálculo, Excel, PDF, ZIP. Datos de la entidad al inicio (FDL, firmante, pie, código 000000005) |
| `backend/app/routes/sivicof.py` | Blueprint `/api/sivicof` |
| `backend/app/plantillas/sivicof/` | `CB-0104.xlsx` (formato con logo), logos, fuentes |
| `frontend/static/js/modulo-10-sivicof.js` | Pantalla |
| `frontend/index.html` | Botón "10 Reportes SIVICOF" (`fragmentos/boton-menu.html`), vista (`fragmentos/vista-sivicof.html`) y `<script>` |
| `frontend/static/js/main.js` | Entrada `sivicof` en `MODULOS` |
| `backend/app/__init__.py` | `register_blueprint(sivicof_bp, url_prefix="/api/sivicof")` |

Endpoints (multipart/form-data, todo en memoria, nada se guarda en el servidor):

- `POST /api/sivicof/calcular` → JSON con filas E…M, controles, alertas, textos y correo.
  Campos: `anio`, `mes`, `fecha_firma`, `hubo_modificacion`, `numero_decreto`, `fecha_decreto`,
  archivos `ejecucion`, `anulaciones`.
- `POST /api/sivicof/generar` → `SIVICOF <MES>.zip`. Mismos campos + archivo `decreto`
  (obligatorio si `hubo_modificacion=1`).

El Excel se arma editando el XML de la plantilla (no con openpyxl) para conservar el logo,
estilos y fórmulas; se marca `fullCalcOnLoad` para que Excel recalcule al abrir.
Sin dependencias nuevas: pandas, python-calamine y reportlab ya están en `requirements.txt`.

## 8. Para otra alcaldía

Cambiar en `sivicof_service.py`: `ENTIDAD_CODIGO`, `FONDO`, `RESPONSABLE_*`, `PROYECTO*`,
`REVISO`, `PIE_DIRECCION`, y reemplazar logos/plantilla en `plantillas/sivicof/`.
En `modulo-10-sivicof.js`: `CORREO_JEFE`.
