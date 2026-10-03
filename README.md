# Sistema de Cruce SIATH · LSI · DATA — JESEP

API en Google Apps Script que cruza tres bases de Excel de la Jefatura Nacional de
Servicio de Policía (JESEP), filtra por la palabra **JESEP**, produce los conteos y
listados solicitados, valida si las novedades se encuentran grabadas y genera el
reporte en **Word (.docx)**.

---

## 1. Archivos del proyecto

| Archivo | Contenido |
|---|---|
| `Code.gs` | API completa (lectura de XLSX, cruce, conteos, generación de DOCX) |
| `Index.html` | Interfaz institucional con fondo holográfico |
| `appsscript.json` | Manifiesto (V8, zona horaria y permisos) |

> Los archivos quedaron en la carpeta `jesep_cruce/` para **no tocar** el proyecto
> anterior (`codigo2.gs` / `index*.html`), que sigue intacto.

---

## 2. Despliegue

1. Abra <https://script.google.com> y cree un proyecto nuevo (**sin** vincularlo a una hoja).
2. En el editor, cree los archivos de scripted (`Code.gs`) y HTML (`Index`),
   o active *Mostrar el archivo de manifiesto* y reemplace `appsscript.json`.
3. Copie el contenido de los tres archivos de esta carpeta.
4. Guarde y despliegue: **Implementar → Nueva implementación → Aplicación web**.
   - *Ejecutar como:* Myself / Usuario que despliega.
   - *Quién puede acceder:* Cualquier persona (o el dominio, si se requiere uso interno).
5. Autorice el acceso a Drive la primera vez y **copie la URL /xxxxx/exec**.

### Permiso de Drive
El manifiesto pide `drive.file`, el más restringido: el script solo ve los archivos
que él mismo crea. Si al institutionalizar apareciera un error de permisos,
cámbielo por `https://www.googleapis.com/auth/drive`.

---

## 3. Uso

1. **BASE SIATH** — `Personal Nominado--JESEP.xlsx`
   Trae `UNIDAD` (que contiene el literal `JESEP`), `ORDEN_PUBLICO_AUTOMATICO`,
   `EMBARGOS_APLICADOS` y `EMB_FAMILIA`.
2. **BASE LSI** — `NOVEDADES GENERALES PRENOMINA DETALL 102026.xlsx`
   Trae `DESCRIPCION`, `ANO_MES`, `NOMBRES` / `APELLIDOS`, documentos de disposición.
3. **BASE DATA** — `REPORTE INCAPACIDADES 10-2026.xlsx`
   Trae `CATEGORIA`, `NUMERO_DIAS`, fechas de inicio y fin.

Pulse **Analizar cruce**. Si un libro tiene varias hojas, el sistema elige la que
contenga `JESEP` y ofrece un selector para cambiar de hoja.

### Sección NOVEDADES
Cargue el archivo de novedades realizadas y pulse **Cruzar novedades con las bases**.
El sistema informa si están **grabadas o no** y lista **quiénes no están**.

### Botón Generar Word
Descarga `Reporte_Cruce_SIATH_LSI_DATA_JESEP.docx` con el reporte completo.

---

## 4. Personal nominado y traslados a otras unidades

### 4.1 Cantidad de personal nominado
El informe (Word y pantalla) indica de forma destacada el total de
**PERSONAL NOMINADO** tomado de BASE SIATH, con el detalle de:
- filas leídas y filtradas, y la hoja utilizada;
- composición por dependencia/cargo (sección **1.1** del Word).

### 4.2 Traslados:ATALMENTE en la unidad filtrada, ahora en otra unidad
Se cruzan las **cédulas** de cada registro `TRASLADO` contra el padrón de personal
nominado de BASE SIATH. El criterio de "cambio de unidad" es:

| Dato | Columna en BASE LSI |
|---|---|
| Unidad actual del funcionario | `EMPL_CONSECUTIVO` |
| Unidad de destino (nómina) | `UNDE_CONSECUTIVO_TRASLA_NOMI` |
| Unidad de destino (laboral) | `UNDE_CONSECUTIVO_TRASLA_LABOR` |

Si el consecutivo de destino difiere del actual, el funcionario **está nominalmente
en la unidad filtrada (JESEP) pero quedó asignado a otra unidad**. Se reportan:
- total de traslados del padrón;
- cuántos quedan en otras unidades y cuántos funcionarios distintos son;
- traslados donde el destino *nómina* difiere del *laboral*;
- distribución por unidad de destino;
- detalle con identificación, nombres, apellidos, unidad de origen, unidad de
  destino, códigos, periodo y documento.

> **Las unidades llegan solo como código.** Para que el informe muestre nombres en
> lugar de `Unidad código 66899`, use el campo **“Catálogo de unidades”** de la barra
> de configuración, con una línea por unidad:
> ```
> 65160=DIRECCION METROPOLITANA
> 66559=COMISARIA NORTE
> 66899=UNIDAD ESPECIAL
> ```

---

## 5. Criterios de búsqueda

| Concepto | Dónde se busca |
|---|---|
| Filtro `JESEP` | Cualquier celda de las 3 bases |
| Adicionales de nómina | Columna que contenga `ADICIONAL` con valor verdadero, o `DESCRIPCION` con el marcador configurado |
| Incapacidades | `DESCRIPCION` / `CATEGORIA` contiene `INCAPACIDAD` |
| Disfrute de vacaciones | `DESCRIPCION` / `CATEGORIA` contiene `VACACION` |
| `EMBARGOS_APLICADOS`, `EMB_FAMILIA` | BASE SIATH, valor mayor que cero |
| NO en `ORDEN_PUBLICO_AUTOMATICO` | BASE SIATH, valor `NO` |
| Comisiones en el exterior | `DESCRIPCION` contiene `EXTERIOR` o `TIPO_COMISION_EXTERIOR` |

Las columnas se localizan **por nombre de encabezado**, no por posición, así que el
orden de columnas puede cambiar sin romper el proceso.

---

## 6. Resultados verificados con los archivos entregados

Ejecutado sobre los tres archivos reales de octubre de 2026:

| Concepto | Resultado |
|---|---|
| **PERSONAL NOMINADO (BASE SIATH, filtro JESEP)** | **1.382** |
| Filas BASE LSI leídas / cruzadas | 1.161 / **1.154** |
| Filas BASE DATA leídas / cruzadas | 15.310 / **146** |
| Incapacidades | 80 registros · 54 funcionarios · 1.173 días |
| Disfrute de vacaciones | 285 registros · 283 funcionarios · 3.517 días |
| `EMBARGOS_APLICADOS` | 56 funcionarios |
| `EMB_FAMILIA` | 17 funcionarios |
| NO en `ORDEN_PUBLICO_AUTOMATICO` | 34 funcionarios |
| Comisiones en el exterior | 5 registros |
| Adicionales de nómina | **0** (ver nota) |

Detalle por concepto:

| Grupo | Concepto | Registros |
|---|---|---|
| Subsidios | Disminución del_subsidio familiar | 4 |
| Subsidios | Extinción del_subsidio familiar | 5 |
| Subsidios | Extinción de la bonificación para la asistencia familiar | 6 |
| Subsidios | Bonificación para la asistencia familiar | 23 |
| Operativos | Traslado | 666 |
| Operativos | Licencia Ley María (paternidad) | 3 |
| Operativos | Licencia remunerada por luto | 10 |
| Operativos | Bonificación para la asistencia familiar | 29 |
| Operativos | Alta titular / sustituto pensionado | **0** |

### Traslados del padrón (sección 4.2)

| Indicador | Resultado |
|---|---|
| Traslados del padrón JESEP | **666** |
| Quedaron en **otras unidades** | **666** (100 %) |
| Funcionarios distintos afectados | **666** |
| Traslados que **no** cambian de unidad | 0 |
| Destino *nómina* distinto del *laboral* | 85 |
| Traslados sin unidad de destino | 0 |

Destinos más frecuentes: `66899` → 573 funcionarios, `66942` → 22, `66913` → 21.

> **Interpretación:** los 666 traslados del padrón JESEP apuntan a una unidad distinta
> de la actual. No es un error del cruce: en la fuente, `EMPL_CONSECUTIVO` y
> `UNDE_CONSECUTIVO_TRASLA_*` nunca coinciden en este lote. Revise con la fuente si
> espera traslados internos; con el catálogo de unidades el destino queda legible.

### Notas importantes sobre los datos entregados

- **Alta de pensionado = 0 dentro del padrón.** El archivo LSI sí trae
  `ALTA TITULAR PENSIONADO` (1) y `ALTA SUSTITUTO PENSIONADO` (4), pero esos
  funcionarios **no están en el padrón de BASE SIATH filtrado por JESEP**
  (identificaciones `1090395153` y `72275526`). El cruce es correcto; conviene
  confirmarlo con la fuente.
- **Adicionales de nómina = 0.** El archivo LSI no trae la columna
  `NOVEDADES_ADICIONALES_NOMINA` ni descripciones con la palabra `ADICIONAL`.
  El criterio queda implementado y se activa solo en cuanto el archivo real la
  traiga; también se puede cambiar el texto en la barra de configuración.

---

## 6. Notas técnicas

- Los `.xlsx` se descomprimen y leen sin bibliotecas externas: se procesa el XML
  de la hoja fila por fila y se resuelven los `sharedStrings`, para no agotar memoria.
- El `.docx` se construye con un escritor **ZIP propio** (método *store* + CRC-32),
  por lo que no depende de librerías ni de plantillas externas.
- Los archivos subidos se guardan en una carpeta de Drive llamada
  `JESEP_CRUCE_TMP`. El botón **Limpiar temporal** la vacía.
- Límites configurables al inicio de `Code.gs` (`CFG`): tamaño de archivo, tamaño
  de hoja, filas máximas y filas de detalle por concepto.