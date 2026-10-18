# Módulo 11 · CDP_CRP Vigencia

Elaboración de **CRP** y **CDP** para adiciones y prórrogas, a partir de los memorandos en PDF y el Reporte CDP de BOGDATA.
Todo se procesa en el navegador (PDF.js + SheetJS + ExcelJS); ni los PDF ni el reporte se envían al backend.

- JS: `docs/static/js/modulo-11-cdp-crp-vigencia.js` (copia en `frontend/`)
- Vista: `docs/index.html`, entre `VISTA 11` y `FIN VISTA 11`
- Instalador idempotente: `integracion_modulo_11/instalar_modulo_11.sh` (ejecutar desde la raíz)

## Pestaña 1 · Elaboración de CRP
Entrada: memorando "Solicitud expedición de CRP" (PDF) + Reporte CDP. Cruce por **No. CDP** (col. H).

| Columna | Origen |
|---|---|
| CONTRATO, CONTRATISTA, CC, RUBRO, No. CDP, VALOR CDP, VALOR ADICION, RADICADO | PDF |
| IMPORTE / ANULACIONES | Reporte CDP: Valor (col. Q) / Anulaciones |
| reserva / Pos | No.Interno CDP (col. U) / No.Posición CDP |
| OBJETO / FONDOS / POS PRES | Objeto / Fondo / Cod.Concepto Gasto |
| CÓDIGO | 145 profesionales/especializados · 148 apoyo/técnicos/gestión |

Salida: `MEMORANDO_SOLICITUD_CRP_<radicado>.xlsx`, con las hojas `base1` y **`DATOS_SAP`**.

### Hoja DATOS_SAP (para la macro LlenarSAP)
Las mismas 13 columnas que la plantilla `Plantilla_Llenar_SAP_v2.xlsm`, hoja DATOS:
Tipo de compromiso = CÓDIGO · Número de Compromiso = contrato · Fecha Inicial = hoy · Fecha Final = 31.12 del año ·
Tipo de pago (02 por defecto) y Mod. de selección (10 por defecto), editables en pantalla · BP vacío (SAP lo calcula) ·
Tipo Doc. = CC · Identif. = CC del memorando · ID Responsable / ID Solicitante = Reporte CDP.

## Pestaña 2 · Elaboración de CDP
Entrada: memorando "Solicitud expedición CDP para Prórroga y Adición" (PDF) + Reporte CDP.
La tabla del PDF se lee por sus bordes. Cruce por **No. de contrato** buscado en el Objeto del reporte CDP.

| Columna | Origen |
|---|---|
| N° PROYECTO, META, CONTRATO, CONTRATISTA, VALOR ADICION, OBJETO, RADICADO | PDF |
| CÓDIGO | Del OBJETO (145/148) |
| IMPORTE | = Valor adición |
| POS PRES / FONDOS / ELEMENTO PEP | Reporte CDP (si no hay CDP: PEP = PM/0005/0101/4599000 + proyecto) |

Validado el 18-oct-2026 con el radicado 20265520010823: 8 contratos, 117 de 117 celdas iguales al formato manual.

## Llenado en BogData (SAP GUI)
El scripting de SAP GUI está **deshabilitado por el servidor** en HEC-SDH S/4-PRD (DisabledByServer, verificado el 18-oct-2026).
Por eso se usa la macro de Excel `LlenarSAP` (SendKeys), que escribe en "Campos adicionales de cabecera":
el orden de los campos está en la hoja ORDEN_TAB y la configuración en CONFIG. La macro nunca guarda: el usuario revisa y confirma.

## OBSERVACIONES (validaciones)
CDP o contrato no encontrado · valor distinto · anulaciones > 0 · adición > saldo · rubro/proyecto no coincide · CDP duplicado · código no determinado.
