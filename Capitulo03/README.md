---LAB_START---
LAB_ID: 03-00-01
---MARKDOWN---
# Estrategias para integrar datos extraídos de Oracle con funciones de Copilot en Excel. Limpieza, transformación y estructuración de grandes volúmenes de datos financieros.

## Metadatos del Laboratorio

| Atributo | Detalle |
|---|---|
| **ID del Laboratorio** | 03-00-01 |
| **Módulo** | Módulo 3 — Integración y preparación de datos financieros desde Oracle |
| **Duración estimada** | 90 minutos |
| **Complejidad** | Alta (Hard) |
| **Nivel Bloom** | Aplicar (Apply) |
| **Modalidad** | Presencial / Virtual asistida |
| **Instructor** | Proporciona archivos CSV de práctica |

---

## Descripción General

Este laboratorio aborda el desafío más frecuente en la contraloría corporativa: la integración, limpieza y estructuración de datos provenientes de Oracle ERP para su análisis en Excel. Los participantes trabajarán con cuatro archivos CSV que replican exportaciones reales de Oracle (movimientos contables, catálogo de cuentas, auxiliar de proveedores y saldos mensuales), incluyendo los problemas típicos de calidad que presentan estas exportaciones en entornos productivos. Utilizando Microsoft 365 Copilot como asistente inteligente, Power Query como motor de transformación y macros VBA generadas por IA, los participantes construirán un pipeline de limpieza documentado, reproducible y alineado con los estándares de la contraloría corporativa sobre un conjunto de más de 10,000 registros financieros.

---

## Objetivos de Aprendizaje

Al finalizar este laboratorio, el participante será capaz de:

- [ ] **Aplicar** estrategias de integración de archivos CSV exportados desde Oracle ERP hacia Excel, utilizando Copilot para asistir en la normalización y estandarización del formato de los datos financieros.
- [ ] **Utilizar** Copilot en Excel para identificar y corregir problemas comunes de calidad de datos: valores nulos, duplicados, formatos de fecha inconsistentes, codificaciones de cuentas incorrectas y caracteres especiales.
- [ ] **Transformar** y estructurar más de 10,000 registros financieros usando Power Query asistido por Copilot y macros VBA generadas por IA, preparando los datos para análisis contable.
- [ ] **Construir** una estructura de datos financieros limpia, estandarizada y auditada, documentando cada paso de transformación realizado en un registro de cambios.

---

## Prerrequisitos

### Conocimientos Previos

- Haber completado las Prácticas 1 y 2 de los módulos anteriores, o tener experiencia equivalente con Copilot en Excel.
- Conocimiento básico de Power Query: acceso desde **Datos → Obtener y transformar datos → Obtener datos**.
- Familiaridad con conceptos contables básicos: libro mayor, auxiliares, pólizas, catálogo de cuentas.
- Comprensión del flujo de integración Oracle → Excel descrito en la Lección 3.1 (ODBC, Power Query, exportación CSV).

### Acceso y Licencias

- Licencia **Microsoft 365 Copilot** activa y verificada en el tenant corporativo.
- Excel 365 actualizado al **canal mensual** (versión 2308 o superior) con Copilot habilitado en el panel lateral.
- Los cuatro archivos CSV de práctica proporcionados por el instructor:
  - `M3_Oracle_Movimientos_Contables.csv` (10,000+ registros)
  - `M3_Oracle_Catalogo_Cuentas.csv`
  - `M3_Oracle_Auxiliar_Proveedores.csv`
  - `M3_Oracle_Saldos_Mensuales.csv`
- Carpeta de trabajo creada en: `C:\LabM3\` (o equivalente en macOS: `~/LabM3/`)

> ⚠️ **Nota de seguridad:** Todos los archivos utilizados son datos financieros **simulados**. No cargar datos financieros reales de la empresa en el entorno de Copilot sin autorización explícita del área de seguridad de la información.

---

## Entorno del Laboratorio

### Hardware Requerido

| Componente | Mínimo | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 8ª gen / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 8 GB | 16 GB (recomendado para 10,000+ registros) |
| Almacenamiento disponible | 500 MB libres para este lab | 2 GB libres |
| Conexión a Internet | 10 Mbps estable | 25 Mbps o superior |
| Resolución de pantalla | 1920×1080 | Monitor dual 1920×1080 |

### Software Requerido

| Software | Versión | Verificación |
|---|---|---|
| Microsoft Excel | 365 versión 2308+ (canal mensual) | `Archivo → Cuenta → Acerca de Excel` |
| Microsoft 365 Copilot | Versión empresarial 2024 | Panel Copilot visible en Excel |
| Microsoft Edge / Chrome | Versión más reciente | Para verificar licencia en portal M365 |
| Archivos CSV del instructor | Módulo 3 — set completo | 4 archivos en `C:\LabM3\` |

### Preparación del Entorno (ejecutar antes de comenzar)

**1. Verificar la licencia de Copilot:**
```text
1. Abrir un navegador y acceder a: https://portal.office.com
2. Iniciar sesión con las credenciales corporativas
3. Ir a: Configuración → Aplicaciones de Microsoft 365 → Copilot
4. Confirmar que el estado muestra "Activo" o "Habilitado"
```

**2. Crear la carpeta de trabajo y colocar los archivos:**
```cmd
REM Windows — ejecutar en símbolo del sistema
mkdir C:\LabM3
REM Copiar los 4 archivos CSV proporcionados por el instructor a C:\LabM3\
```

**3. Verificar la versión de Excel:**
```text
Excel → Archivo → Cuenta → Información de producto
Confirmar: "Microsoft 365 Apps para empresas" y canal "Canal mensual"
```

**4. Habilitar el panel de Copilot en Excel:**
```text
Excel → pestaña "Inicio" → botón "Copilot" (extremo derecho de la cinta)
Si no aparece: Archivo → Opciones → General → verificar "Mostrar características de Copilot"
```

---

## Pasos del Laboratorio

---

### Paso 1: Importar el archivo de movimientos contables de Oracle a Excel mediante Power Query

**Objetivo:** Cargar correctamente el archivo CSV principal (`M3_Oracle_Movimientos_Contables.csv`) en Excel usando Power Query, identificando desde el inicio los problemas de estructura y codificación que presenta la exportación de Oracle.

#### Instrucciones

1. Abrir **Excel 365** con un libro en blanco. Guardar inmediatamente como:
   ```text
   Nombre: M3_Pipeline_Datos_Oracle.xlsx
   Ubicación: C:\LabM3\
   ```

2. Ir a la pestaña **Datos** → **Obtener datos** → **Desde un archivo** → **Desde texto/CSV**.

3. Navegar a `C:\LabM3\` y seleccionar `M3_Oracle_Movimientos_Contables.csv`. Hacer clic en **Importar**.

4. En el cuadro de diálogo de vista previa, **no hacer clic en "Cargar"** todavía. En su lugar, hacer clic en **Transformar datos** para abrir el **Editor de Power Query**.

5. En el Editor de Power Query, observar y documentar los siguientes problemas visibles en la vista previa:
   ```text
   - Columnas con nombres en inglés técnico de Oracle (ej. ACCOUNTED_DR, PERIOD_NAME, CODE_COMBINATION_ID)
   - Columna de fecha con formato mixto (algunas filas: DD/MM/YYYY, otras: YYYY-MM-DD)
   - Columna de montos con símbolos de moneda pegados al número (ej. "$1,234.56" o "MXN 5000")
   - Registros con valores nulos en columnas críticas
   - Posibles duplicados por reruns de Oracle (mismo JE_HEADER_ID repetido)
   ```

6. En el panel de **Configuración de la consulta** (derecha), renombrar la consulta de `M3_Oracle_Movimientos_Contables` a:
   ```text
   Oracle_Movimientos_RAW
   ```

7. **Sin cerrar el Editor de Power Query**, proceder al Paso 2.

#### Resultado Esperado

El Editor de Power Query muestra la tabla `Oracle_Movimientos_RAW` con los 10,000+ registros cargados, columnas visibles con nombres en inglés de Oracle, y al menos 3 indicadores de tipo de dato incorrectos (ícono de advertencia en el encabezado de columna).

#### Verificación

```text
✓ El contador de filas en la barra inferior del Editor de Power Query muestra "10,000+ filas"
✓ Se identifican al menos 5 columnas con nombres técnicos de Oracle (en inglés)
✓ Al menos 2 columnas muestran el ícono de tipo de dato incorrecto (ABC en lugar de 123 o fecha)
✓ La consulta está renombrada como "Oracle_Movimientos_RAW"
```

---

### Paso 2: Usar Copilot para diagnosticar los problemas de calidad de datos

**Objetivo:** Aprovechar Copilot en Excel para obtener un diagnóstico automático de los problemas de calidad en los datos importados, generando una lista priorizada de transformaciones necesarias antes de proceder a la limpieza manual.

#### Instrucciones

1. Cerrar el Editor de Power Query haciendo clic en **Cerrar y cargar en...** → seleccionar **Solo crear conexión** → **Aceptar**. Esto carga los datos como conexión sin volcarlos aún a la hoja.

2. En el libro de Excel, ir a la pestaña **Datos** → **Consultas y conexiones**. Confirmar que `Oracle_Movimientos_RAW` aparece en el panel.

3. Hacer clic derecho sobre la consulta `Oracle_Movimientos_RAW` → **Cargar en...** → **Tabla** → seleccionar **Nueva hoja de cálculo** → **Aceptar**. Renombrar la hoja como `RAW_Movimientos`.

4. Seleccionar cualquier celda dentro de la tabla cargada. Abrir el **panel de Copilot** (botón en la pestaña Inicio o combinación de teclas según configuración del tenant).

5. Escribir el siguiente prompt en el panel de Copilot:
   ```text
   Analiza esta tabla de datos financieros exportada desde Oracle ERP. 
   Identifica todos los problemas de calidad de datos que encuentres: 
   valores nulos, duplicados, formatos inconsistentes en fechas y montos, 
   nombres de columnas en inglés que deban traducirse, y cualquier anomalía 
   que impida el análisis contable. Presenta los resultados como una lista 
   priorizada con el tipo de problema, la columna afectada y la acción 
   correctiva recomendada.
   ```

6. Revisar la respuesta de Copilot. Copiar el diagnóstico generado y pegarlo en una **nueva hoja** llamada `Log_Transformaciones`. En la celda A1 escribir:
   ```text
   REGISTRO DE TRANSFORMACIONES — Pipeline Oracle → Excel
   Fecha: [fecha actual]    Responsable: [nombre del participante]
   ```

7. Debajo del encabezado, pegar el diagnóstico de Copilot. Este documento será el **registro de auditoría** del pipeline de limpieza.

8. Volver a la hoja `RAW_Movimientos` y escribir un segundo prompt en Copilot:
   ```text
   ¿Cuántos registros duplicados existen en esta tabla considerando como 
   clave única la combinación de las columnas JE_HEADER_ID, JE_LINE_NUM 
   y PERIOD_NAME? Muéstrame el conteo y los primeros 5 ejemplos de duplicados.
   ```

9. Anotar el número de duplicados reportado por Copilot en la hoja `Log_Transformaciones`, columna correspondiente.

#### Resultado Esperado

Copilot genera una lista diagnóstica con al menos 6 categorías de problemas identificados, incluyendo: columnas con nombres en inglés, formatos de fecha mixtos, montos con caracteres no numéricos, valores nulos en columnas clave, duplicados por rerun de Oracle, y posibles problemas de codificación de caracteres. El registro `Log_Transformaciones` contiene el diagnóstico inicial documentado.

#### Verificación

```text
✓ El panel de Copilot muestra una respuesta estructurada con categorías de problemas
✓ La hoja "Log_Transformaciones" contiene el diagnóstico con fecha y responsable
✓ Se identificó el número exacto de duplicados en la tabla RAW
✓ Copilot sugiere al menos 4 acciones correctivas específicas con nombres de columna
```

---

### Paso 3: Limpiar y transformar datos en Power Query — Fase 1 (Estructura y nombres)

**Objetivo:** Aplicar las primeras transformaciones estructurales en Power Query: renombrar columnas al español, corregir tipos de datos y eliminar columnas innecesarias, usando las sugerencias de Copilot como guía.

#### Instrucciones

1. En el panel **Consultas y conexiones**, hacer doble clic en `Oracle_Movimientos_RAW` para abrir el Editor de Power Query.

2. **Renombrar columnas al español corporativo.** Hacer doble clic en cada encabezado de columna y renombrar según la siguiente tabla estándar:

   | Nombre Original (Oracle) | Nombre Nuevo (Español) |
   |---|---|
   | `PERIOD_NAME` | `Periodo_Contable` |
   | `JE_HEADER_ID` | `ID_Poliza` |
   | `JE_LINE_NUM` | `Num_Linea` |
   | `CODE_COMBINATION_ID` | `ID_Combinacion_Cuenta` |
   | `SEGMENT1` | `Cuenta_Contable` |
   | `SEGMENT2` | `Centro_Costo` |
   | `ACCOUNTED_DR` | `Debito` |
   | `ACCOUNTED_CR` | `Credito` |
   | `DESCRIPTION` | `Descripcion_Asiento` |
   | `CREATION_DATE` | `Fecha_Creacion` |
   | `CREATED_BY` | `Usuario_Creacion` |
   | `STATUS` | `Estado_Poliza` |

3. **Eliminar columnas técnicas innecesarias para el análisis contable.** Seleccionar las columnas que no están en la tabla anterior (columnas de metadatos internos de Oracle como `LAST_UPDATE_DATE`, `LAST_UPDATED_BY`, `REQUEST_ID`). Hacer clic derecho → **Quitar columnas**.

4. **Corregir el tipo de dato de la columna `Periodo_Contable`:** Hacer clic en el ícono de tipo de dato de la columna → seleccionar **Texto** (los períodos de Oracle como "MAR-25" deben mantenerse como texto, no convertirse a fecha).

5. **Corregir el tipo de dato de `Fecha_Creacion`:** Hacer clic en el ícono → seleccionar **Fecha** → si aparece un error de conversión, hacer clic en **Agregar nuevo paso** y usar la opción de reemplazar errores con `null`.

6. **Verificar los tipos de dato de `Debito` y `Credito`:** Deben ser **Número decimal**. Si aparecen como texto, hacer clic en el ícono de tipo → **Número decimal**.

7. En el panel **Pasos aplicados** (derecha), verificar que cada acción aparece registrada. Estos pasos son la trazabilidad del pipeline.

8. Hacer clic en **Cerrar y cargar** → **Cerrar y cargar en...** → seleccionar la hoja `RAW_Movimientos` (reemplazar). Confirmar con **Aceptar**.

#### Resultado Esperado

La tabla en la hoja `RAW_Movimientos` muestra todas las columnas con nombres en español, tipos de dato correctos (sin íconos de advertencia en encabezados), y sin columnas de metadatos técnicos de Oracle. El panel de **Pasos aplicados** en Power Query muestra un mínimo de 8 pasos documentados.

#### Verificación

```text
✓ Todas las columnas tienen nombres en español según la tabla de mapeo
✓ La columna "Periodo_Contable" es de tipo Texto (no Fecha)
✓ Las columnas "Debito" y "Credito" son de tipo Número decimal
✓ No existen columnas con nombres en inglés de Oracle en la tabla final
✓ El panel "Pasos aplicados" registra mínimo 8 transformaciones
```

---

### Paso 4: Limpiar y transformar datos en Power Query — Fase 2 (Calidad de datos)

**Objetivo:** Eliminar duplicados, corregir formatos de fecha inconsistentes, limpiar montos con caracteres especiales y tratar valores nulos, implementando las correcciones identificadas en el diagnóstico de Copilot del Paso 2.

#### Instrucciones

1. Abrir nuevamente el Editor de Power Query haciendo doble clic en la consulta `Oracle_Movimientos_RAW` desde el panel de consultas.

2. **Eliminar registros duplicados por rerun de Oracle:**
   - Seleccionar simultáneamente las columnas `ID_Poliza`, `Num_Linea` y `Periodo_Contable` (Ctrl+clic en cada encabezado).
   - En la pestaña **Inicio** del Editor → **Quitar filas** → **Quitar duplicados**.
   - Anotar en la hoja `Log_Transformaciones`: número de filas antes y después de eliminar duplicados.

3. **Estandarizar el formato de fecha en `Fecha_Creacion`:**
   
   En el Editor de Power Query, agregar una columna personalizada. Ir a **Agregar columna** → **Columna personalizada**. Nombrarla `Fecha_Creacion_Limpia` e ingresar la siguiente fórmula M:
   ```m
   = try Date.From([Fecha_Creacion]) otherwise 
     try Date.FromText([Fecha_Creacion], "dd/MM/yyyy") otherwise
     try Date.FromText([Fecha_Creacion], "yyyy-MM-dd") otherwise
     null
   ```
   
   Una vez creada la columna nueva, eliminar la columna `Fecha_Creacion` original y renombrar `Fecha_Creacion_Limpia` como `Fecha_Creacion`.

4. **Limpiar la columna `Debito` de símbolos de moneda:**
   
   Seleccionar la columna `Debito`. Ir a **Transformar** → **Reemplazar valores**:
   ```text
   Valor a buscar: $
   Reemplazar por: (vacío)
   ```
   Repetir para los siguientes caracteres: `MXN`, `USD`, `,` (coma como separador de miles), y espacios en blanco al inicio/fin.
   
   Alternativamente, usar una columna personalizada con fórmula M:
   ```m
   = Number.From(
       Text.Trim(
         Text.Replace(
           Text.Replace(
             Text.Replace([Debito], "$", ""),
           "MXN", ""),
         ",", "")
       )
     )
   ```
   Repetir el mismo proceso para la columna `Credito`.

5. **Tratar valores nulos en columnas críticas:**
   - Columna `Cuenta_Contable`: Seleccionar → **Transformar** → **Reemplazar valores** → buscar `null` → reemplazar con `"0000"` (código de cuenta desconocida para auditoría).
   - Columna `Centro_Costo`: Reemplazar `null` con `"SINCC"` (sin centro de costo).
   - Columna `Descripcion_Asiento`: Reemplazar `null` con `"SIN DESCRIPCION"`.
   - Columna `Debito`: Reemplazar `null` con `0`.
   - Columna `Credito`: Reemplazar `null` con `0`.

6. **Filtrar solo pólizas publicadas (Status = 'P'):**
   - Columna `Estado_Poliza` → clic en la flecha de filtro → desmarcar todos los valores excepto `P` (Published/Publicado).
   - Anotar en `Log_Transformaciones` cuántos registros fueron excluidos por este filtro.

7. Hacer clic en **Cerrar y cargar** → **Cerrar y cargar en...** → reemplazar la hoja `RAW_Movimientos`.

#### Resultado Esperado

La tabla `RAW_Movimientos` contiene datos sin duplicados, con fechas en formato uniforme, montos como valores numéricos puros (sin símbolos), valores nulos reemplazados por valores estándar de auditoría, y únicamente pólizas con estado publicado. El registro en `Log_Transformaciones` documenta el impacto de cada transformación (filas afectadas).

#### Verificación

```text
✓ Cero duplicados en la combinación ID_Poliza + Num_Linea + Periodo_Contable
✓ La columna "Fecha_Creacion" muestra fechas en formato uniforme (DD/MM/YYYY o según configuración regional)
✓ Las columnas "Debito" y "Credito" no contienen símbolos de moneda ni comas
✓ No existen celdas vacías en las 5 columnas críticas tratadas
✓ La tabla solo contiene registros con Estado_Poliza = "P"
✓ Log_Transformaciones registra el conteo de filas antes y después de cada transformación
```

---

### Paso 5: Importar y cruzar el Catálogo de Cuentas de Oracle

**Objetivo:** Cargar el archivo `M3_Oracle_Catalogo_Cuentas.csv`, limpiarlo y combinarlo con la tabla de movimientos para enriquecer los datos con la descripción oficial de cada cuenta contable.

#### Instrucciones

1. Ir a **Datos** → **Obtener datos** → **Desde un archivo** → **Desde texto/CSV**.

2. Seleccionar `M3_Oracle_Catalogo_Cuentas.csv` desde `C:\LabM3\`. Hacer clic en **Transformar datos**.

3. En el Editor de Power Query, realizar las siguientes transformaciones básicas:
   - Renombrar la consulta como `Oracle_Catalogo_Cuentas`.
   - Identificar la columna de código de cuenta (probablemente `SEGMENT_VALUE` o `FLEX_VALUE`) y renombrarla como `Cuenta_Contable`.
   - Identificar la columna de descripción (probablemente `DESCRIPTION` o `SEGMENT_VALUE_DESC`) y renombrarla como `Descripcion_Cuenta`.
   - Eliminar columnas innecesarias (metadatos de Oracle).

4. **Problema crítico de calidad — Ceros a la izquierda:** En Oracle, los códigos de cuenta frecuentemente tienen ceros a la izquierda (ej. `"0510"`) que se pierden al importar como número. Verificar que la columna `Cuenta_Contable` sea de tipo **Texto** (no número). Si es número, cambiar el tipo a Texto. Esto garantiza que `"0510"` no se convierta en `510`.

5. Hacer clic en **Cerrar y cargar** → **Solo crear conexión**.

6. Ahora combinar las dos consultas. En el panel **Consultas y conexiones**, hacer doble clic en `Oracle_Movimientos_RAW` para abrir el Editor de Power Query.

7. En el Editor, ir a **Inicio** → **Combinar consultas** → **Combinar consultas como nueva**:
   ```text
   Tabla izquierda: Oracle_Movimientos_RAW
   Columna de combinación izquierda: Cuenta_Contable
   
   Tabla derecha: Oracle_Catalogo_Cuentas
   Columna de combinación derecha: Cuenta_Contable
   
   Tipo de combinación: Combinación externa izquierda (todos los registros de movimientos)
   ```

8. En la nueva consulta combinada, expandir la columna del catálogo haciendo clic en el ícono de expansión (⇔) junto al encabezado `Oracle_Catalogo_Cuentas`. Seleccionar únicamente la columna `Descripcion_Cuenta`. Desmarcar "Usar nombre de columna original como prefijo".

9. Renombrar la consulta combinada como `Movimientos_Enriquecidos`.

10. Hacer clic en **Cerrar y cargar** → **Nueva hoja de cálculo**. Renombrar la hoja como `Movimientos_Limpios`.

#### Resultado Esperado

La hoja `Movimientos_Limpios` contiene la tabla enriquecida con todos los movimientos contables limpios y una columna adicional `Descripcion_Cuenta` que muestra el nombre oficial de cada cuenta según el catálogo de Oracle. Los códigos de cuenta mantienen sus ceros a la izquierda.

#### Verificación

```text
✓ La hoja "Movimientos_Limpios" existe con la tabla "Movimientos_Enriquecidos"
✓ La columna "Descripcion_Cuenta" está presente y tiene valores para la mayoría de registros
✓ Los códigos de cuenta con ceros a la izquierda (ej. "0510") se mantienen como texto, no como número
✓ El tipo de combinación "externa izquierda" preserva el 100% de los movimientos originales
✓ Los registros sin coincidencia en el catálogo muestran "null" en Descripcion_Cuenta (identificables para auditoría)
```

---

### Paso 6: Generar una macro VBA de limpieza automatizada con Copilot

**Objetivo:** Usar Copilot para generar una macro VBA que automatice el proceso de limpieza de datos en futuras importaciones de Oracle, creando un script reutilizable y documentado que el equipo de contraloría pueda ejecutar mensualmente.

#### Instrucciones

1. Asegurarse de estar en la hoja `Movimientos_Limpios`. Abrir el panel de **Copilot**.

2. Escribir el siguiente prompt en Copilot:
   ```text
   Genera una macro VBA para Excel que realice las siguientes tareas de 
   limpieza en una tabla llamada "Movimientos_Enriquecidos":
   
   1. Elimine filas duplicadas basándose en las columnas "ID_Poliza", 
      "Num_Linea" y "Periodo_Contable"
   2. Reemplace valores nulos en la columna "Cuenta_Contable" con "0000"
   3. Reemplace valores nulos en la columna "Centro_Costo" con "SINCC"
   4. Reemplace valores nulos en "Debito" y "Credito" con 0
   5. Elimine espacios en blanco al inicio y final de todas las columnas de texto
   6. Agregue una columna "Fecha_Procesamiento" con la fecha y hora actual
   7. Muestre un mensaje al finalizar indicando cuántas filas fueron procesadas
   
   La macro debe incluir manejo de errores y comentarios explicativos en español.
   Nómbrala "LimpiarDatosOracle".
   ```

3. Copilot generará el código VBA. Revisarlo cuidadosamente. El código debe tener una estructura similar a la siguiente (el código exacto lo genera Copilot):

   ```vba
   Sub LimpiarDatosOracle()
   ' ============================================================
   ' MACRO: LimpiarDatosOracle
   ' PROPÓSITO: Limpieza automatizada de exportaciones Oracle ERP
   ' GENERADA CON: Microsoft 365 Copilot
   ' FECHA DE GENERACIÓN: [fecha]
   ' RESPONSABLE: [nombre]
   ' ============================================================
   
       Dim ws As Worksheet
       Dim tbl As ListObject
       Dim filasProcesadas As Long
       
       On Error GoTo ManejadorError
       
       ' Identificar la hoja y tabla de trabajo
       Set ws = ThisWorkbook.Sheets("Movimientos_Limpios")
       Set tbl = ws.ListObjects("Movimientos_Enriquecidos")
       
       filasProcesadas = tbl.DataBodyRange.Rows.Count
       
       ' PASO 1: Eliminar duplicados
       ' [Copilot genera el código específico aquí]
       
       ' PASO 2-4: Reemplazar nulos
       ' [Copilot genera el código específico aquí]
       
       ' PASO 5: Limpiar espacios en blanco
       ' [Copilot genera el código específico aquí]
       
       ' PASO 6: Agregar columna de fecha de procesamiento
       ' [Copilot genera el código específico aquí]
       
       MsgBox "Limpieza completada. Filas procesadas: " & filasProcesadas, _
              vbInformation, "Pipeline Oracle - Contraloría"
       Exit Sub
       
   ManejadorError:
       MsgBox "Error en la limpieza: " & Err.Description, vbCritical
   End Sub
   ```

4. Hacer clic en el botón **Insertar código** que aparece en la respuesta de Copilot (o copiar manualmente el código). Si Copilot ofrece la opción de insertar directamente en el módulo VBA, utilizarla.

5. Si el código no se inserta automáticamente, abrirlo manualmente: **Alt + F11** para abrir el Editor de VBA → **Insertar** → **Módulo** → pegar el código.

6. En el Editor de VBA, revisar el código línea por línea. Si alguna parte no está clara, volver al panel de Copilot y preguntar:
   ```text
   Explícame qué hace la sección de "Eliminar duplicados" de la macro 
   que generaste. ¿Cómo funciona el método RemoveDuplicates en VBA?
   ```

7. Ejecutar la macro presionando **F5** dentro del Editor de VBA, o desde Excel: **Desarrollador** → **Macros** → seleccionar `LimpiarDatosOracle` → **Ejecutar**.

   > 📌 **Nota:** Si la pestaña **Desarrollador** no está visible en Excel, habilitarla en: **Archivo → Opciones → Personalizar cinta de opciones → marcar "Desarrollador"**.

8. Verificar que el mensaje de confirmación aparece con el número correcto de filas procesadas.

9. Guardar el archivo como **Excel con macros habilitadas**: **Archivo → Guardar como → Tipo: Libro de Excel habilitado para macros (.xlsm)**. Renombrar como `M3_Pipeline_Datos_Oracle.xlsm`.

#### Resultado Esperado

Una macro VBA funcional llamada `LimpiarDatosOracle` está almacenada en el módulo del libro, ejecuta sin errores, muestra el mensaje de confirmación con el conteo de filas procesadas, y el archivo se guarda en formato `.xlsm`. El código incluye comentarios en español y manejo de errores.

#### Verificación

```text
✓ La macro "LimpiarDatosOracle" existe en el módulo VBA del libro
✓ La macro se ejecuta sin errores (sin cuadros de error rojos)
✓ El mensaje de confirmación muestra un número de filas > 0
✓ La columna "Fecha_Procesamiento" fue agregada a la tabla con la fecha y hora actual
✓ El archivo se guardó como .xlsm (con macros habilitadas)
✓ El código VBA incluye comentarios en español
```

---

### Paso 7: Construir el reporte de estructura de datos limpia con Copilot

**Objetivo:** Utilizar Copilot para generar un resumen ejecutivo de la calidad de los datos procesados, crear una tabla dinámica de validación y construir la estructura final de datos lista para análisis contable, documentando el pipeline completo.

#### Instrucciones

1. Ir a la hoja `Movimientos_Limpios`. Seleccionar cualquier celda dentro de la tabla `Movimientos_Enriquecidos`.

2. Abrir el panel de **Copilot** y escribir:
   ```text
   Analiza la tabla "Movimientos_Enriquecidos" y proporciona:
   1. Estadísticas generales: total de registros, rangos de fechas, 
      períodos contables presentes
   2. Suma total de débitos y créditos por período contable
   3. Top 10 cuentas contables con mayor movimiento (suma de débitos)
   4. Identificación de cuentas sin descripción en el catálogo (Descripcion_Cuenta vacía)
   5. Verificación de cuadre contable: ¿el total de débitos es igual al total de créditos?
   ```

3. Copilot generará el análisis. Si ofrece crear una tabla dinámica o insertar fórmulas, aceptar la sugerencia.

4. Crear una nueva hoja llamada `Resumen_Calidad`. En esta hoja, con ayuda de Copilot, construir un tablero de calidad de datos con la siguiente estructura:

   ```text
   SECCIÓN A — Métricas del Pipeline
   - Total registros importados (antes de limpieza)
   - Total registros después de eliminar duplicados
   - Registros excluidos por estado (no publicados)
   - Porcentaje de registros con descripción de cuenta completa
   
   SECCIÓN B — Validación Contable
   - Total débitos del período
   - Total créditos del período
   - Diferencia (debe ser 0 o mínima)
   - Períodos contables presentes en los datos
   
   SECCIÓN C — Alertas de Calidad
   - Cuentas sin descripción en catálogo
   - Registros con Centro_Costo = "SINCC"
   - Registros con Cuenta_Contable = "0000"
   ```

5. Para cada métrica, escribir un prompt específico a Copilot. Por ejemplo:
   ```text
   Calcula el porcentaje de registros en la tabla "Movimientos_Enriquecidos" 
   que tienen un valor no nulo y diferente de "0000" en la columna 
   "Cuenta_Contable". Muestra el resultado como porcentaje con 2 decimales.
   ```

6. Completar la hoja `Log_Transformaciones` con un resumen final del pipeline:

   | Paso | Transformación Aplicada | Filas Antes | Filas Después | Diferencia | Responsable |
   |---|---|---|---|---|---|
   | 1 | Importación CSV Oracle | — | [valor] | — | [nombre] |
   | 2 | Diagnóstico Copilot | [valor] | [valor] | — | Copilot |
   | 3 | Renombrado de columnas | [valor] | [valor] | 0 | Power Query |
   | 4 | Eliminación de duplicados | [valor] | [valor] | [valor] | Power Query |
   | 5 | Limpieza de nulos | [valor] | [valor] | 0 | Power Query |
   | 6 | Filtro Estado='P' | [valor] | [valor] | [valor] | Power Query |
   | 7 | Combinación con catálogo | [valor] | [valor] | 0 | Power Query |
   | 8 | Ejecución macro VBA | [valor] | [valor] | 0 | VBA |

7. Guardar el archivo.

#### Resultado Esperado

La hoja `Resumen_Calidad` presenta un tablero completo con las tres secciones (Métricas del Pipeline, Validación Contable, Alertas de Calidad) con valores calculados. La hoja `Log_Transformaciones` tiene el registro completo del pipeline con todos los conteos documentados. El cuadre contable (débitos = créditos) está verificado.

#### Verificación

```text
✓ La hoja "Resumen_Calidad" existe con las tres secciones completadas
✓ El total de débitos y créditos está calculado y visible
✓ Se identificó el número de cuentas sin descripción en el catálogo
✓ El Log_Transformaciones tiene todas las filas de la tabla de resumen completadas
✓ La diferencia entre débitos y créditos es $0 o se documenta la razón de la diferencia
✓ El archivo está guardado como .xlsm
```

---

## Validación y Pruebas Finales

Al concluir todos los pasos, realizar las siguientes verificaciones integrales para confirmar que el pipeline de datos está completo y funcional:

### Lista de Verificación Final

```text
ESTRUCTURA DEL LIBRO
[ ] El archivo se llama "M3_Pipeline_Datos_Oracle.xlsm"
[ ] El libro contiene las siguientes hojas: RAW_Movimientos, Movimientos_Limpios, 
    Resumen_Calidad, Log_Transformaciones
[ ] El libro contiene al menos 2 consultas en Power Query: 
    Oracle_Movimientos_RAW y Oracle_Catalogo_Cuentas

CALIDAD DE DATOS
[ ] La tabla "Movimientos_Enriquecidos" no contiene duplicados en la clave 
    ID_Poliza + Num_Linea + Periodo_Contable
[ ] Las columnas Debito y Credito son numéricas (sin símbolos de moneda)
[ ] La columna Fecha_Creacion está en formato fecha uniforme
[ ] Los códigos de cuenta con ceros a la izquierda se preservan como texto
[ ] No existen celdas completamente vacías en columnas críticas

AUTOMATIZACIÓN
[ ] La macro "LimpiarDatosOracle" existe y se ejecuta sin errores
[ ] La macro agrega la columna "Fecha_Procesamiento" al ejecutarse
[ ] El archivo está en formato .xlsm (habilitado para macros)

DOCUMENTACIÓN
[ ] Log_Transformaciones contiene el diagnóstico inicial de Copilot
[ ] Log_Transformaciones tiene la tabla de resumen del pipeline completa
[ ] Resumen_Calidad muestra validación de cuadre contable
```

### Prueba de Regresión: Simular una Nueva Importación

Para verificar que el pipeline es reutilizable, realizar la siguiente prueba rápida:

1. Duplicar la hoja `RAW_Movimientos` y renombrarla `TEST_Reimportacion`.
2. Introducir manualmente 3 filas duplicadas al final de la tabla en la hoja `TEST_Reimportacion`.
3. Ejecutar la macro `LimpiarDatosOracle` desde **Desarrollador → Macros**.
4. Verificar que las 3 filas duplicadas fueron eliminadas y el mensaje de confirmación refleja el conteo correcto.

```text
✓ PRUEBA APROBADA: La macro eliminó los 3 duplicados introducidos manualmente
✓ El mensaje de confirmación muestra el número correcto de filas procesadas
✓ La columna "Fecha_Procesamiento" se actualizó con la fecha y hora de la ejecución
```

---

## Solución de Problemas

### Problema 1: Copilot no analiza la tabla y responde "No puedo acceder a los datos de esta hoja"

**Síntoma:** Al escribir prompts en el panel de Copilot referentes a la tabla de movimientos, Copilot responde que no puede acceder a los datos, que la tabla está vacía, o que no encuentra el rango seleccionado. El panel de Copilot se abre pero no genera análisis basados en los datos de la hoja.

**Causa raíz:** Copilot en Excel requiere que los datos estén en una **tabla estructurada de Excel** (formato de tabla con encabezados, no simplemente un rango de celdas). Si los datos se cargaron desde Power Query como rango normal, o si la tabla tiene el cursor posicionado fuera de ella, Copilot no puede interpretarla. Adicionalmente, Copilot puede no funcionar si el archivo está guardado localmente sin sincronización con OneDrive/SharePoint.

**Solución:**
```text
1. Verificar que los datos están en formato de tabla:
   - Hacer clic dentro del rango de datos
   - Ir a: Insertar → Tabla → confirmar que "La tabla tiene encabezados" está marcado
   - La tabla debe tener un nombre (ej. "Movimientos_Enriquecidos"), visible en 
     la pestaña "Diseño de tabla" cuando se selecciona

2. Guardar el archivo en OneDrive o SharePoint corporativo:
   - Archivo → Guardar como → OneDrive - [Nombre de la organización]
   - Esperar a que la sincronización se complete (ícono de nube en la barra de título)
   - Reabrir el archivo desde OneDrive y volver a intentar con Copilot

3. Si el problema persiste, cerrar y reabrir Excel completamente, 
   asegurarse de tener conexión a Internet activa y reintentar.
```

---

### Problema 2: Power Query pierde los ceros a la izquierda en los códigos de cuenta al importar el CSV

**Síntoma:** Al importar `M3_Oracle_Catalogo_Cuentas.csv` o `M3_Oracle_Movimientos_Contables.csv`, los códigos de cuenta que en Oracle tienen ceros a la izquierda (ej. `"0510"`, `"0100"`, `"0050"`) aparecen en Excel como números sin ceros (ej. `510`, `100`, `50`). Esto provoca que la combinación de consultas en el Paso 5 no encuentre coincidencias y la columna `Descripcion_Cuenta` quede completamente vacía.

**Causa raíz:** Al importar archivos CSV, Excel y Power Query intentan inferir automáticamente el tipo de dato de cada columna. Los códigos de cuenta que contienen solo dígitos son interpretados como números enteros, eliminando los ceros a la izquierda. Esto rompe la integridad referencial entre la tabla de movimientos y el catálogo de cuentas, ya que `"0510"` (texto) no es igual a `510` (número).

**Solución:**
```text
1. En el Editor de Power Query, ANTES de hacer cualquier transformación:
   - Seleccionar la columna de código de cuenta (Cuenta_Contable o SEGMENT_VALUE)
   - Hacer clic en el ícono de tipo de dato en el encabezado de la columna
   - Cambiar el tipo a "Texto" (no "Número entero" ni "Número decimal")
   - Si Power Query ya convirtió el tipo, hacer clic en "Reemplazar paso actual"

2. Si los ceros ya se perdieron (columna muestra 510 en lugar de 0510):
   - Agregar una columna personalizada con la fórmula M:
     = Text.PadStart(Text.From([Cuenta_Contable]), 4, "0")
     (reemplazar 4 con la longitud correcta del código de cuenta en Oracle)
   - Esto restaura los ceros a la izquierda hasta la longitud estándar

3. Para prevenir el problema en futuras importaciones:
   - Al hacer clic en "Transformar datos" (no "Cargar"), ir inmediatamente a 
     la columna de código de cuenta y cambiar su tipo a Texto ANTES de 
     cualquier otro paso. Este es el primer paso que debe aparecer en 
     "Pasos aplicados".

4. Verificar que ambas tablas (movimientos y catálogo) tienen el código 
   de cuenta como tipo Texto con la misma longitud antes de realizar 
   la combinación de consultas.
```

---

## Limpieza del Entorno

Al finalizar el laboratorio, realizar las siguientes acciones para dejar el entorno en orden:

### Archivos a Conservar

```text
✓ C:\LabM3\M3_Pipeline_Datos_Oracle.xlsm  ← Archivo principal del laboratorio
✓ C:\LabM3\M3_Oracle_Movimientos_Contables.csv  ← Archivo fuente (conservar)
✓ C:\LabM3\M3_Oracle_Catalogo_Cuentas.csv  ← Archivo fuente (conservar)
✓ C:\LabM3\M3_Oracle_Auxiliar_Proveedores.csv  ← Para práctica extendida
✓ C:\LabM3\M3_Oracle_Saldos_Mensuales.csv  ← Para práctica extendida
```

### Acciones de Cierre

1. **Guardar el archivo final** con todos los cambios: `Ctrl + S`.

2. **Verificar que el archivo está en formato .xlsm** (no .xlsx). Si se guardó como .xlsx, las macros se habrán perdido. Volver a guardar como `.xlsm`.

3. **Cerrar las conexiones de Power Query activas** para liberar memoria:
   ```text
   Datos → Consultas y conexiones → clic derecho en cada consulta → 
   "Propiedades" → desmarcar "Actualizar al abrir el archivo" 
   (opcional, para entornos sin acceso a Oracle)
   ```

4. **Eliminar la hoja de prueba** creada en la validación final:
   ```text
   Clic derecho en la pestaña "TEST_Reimportacion" → Eliminar → Confirmar
   ```

5. **Cerrar el Editor de VBA** si está abierto: `Alt + F4` dentro del Editor de VBA (no de Excel).

6. **Confirmar con el instructor** que el archivo `M3_Pipeline_Datos_Oracle.xlsm` es visible y accesible antes de cerrar Excel.

> ⚠️ **Recordatorio de seguridad:** Si se utilizó OneDrive para sincronizar el archivo, confirmar que los datos simulados del laboratorio no quedan compartidos en carpetas de acceso público dentro del tenant corporativo.

---

## Resumen

### Lo que se logró en este laboratorio

En este laboratorio se construyó un **pipeline completo de integración y limpieza de datos Oracle → Excel**, abordando los desafíos reales de calidad de datos que enfrenta la contraloría corporativa en su operación diaria. Los participantes aplicaron las siguientes capacidades:

| Capacidad Desarrollada | Herramienta Utilizada | Resultado |
|---|---|---|
| Importación de datos CSV de Oracle | Power Query | Tabla RAW con 10,000+ registros cargados |
| Diagnóstico de calidad de datos | Copilot en Excel | Lista priorizada de problemas identificados |
| Renombrado y tipado de columnas | Power Query | Estructura estandarizada en español |
| Eliminación de duplicados Oracle | Power Query | Datos sin reruns de Oracle |
| Estandarización de fechas y montos | Power Query (fórmulas M) | Formatos uniformes y numéricos |
| Cruce con catálogo de cuentas | Power Query (Combinar) | Datos enriquecidos con descripciones |
| Automatización de limpieza | VBA generado por Copilot | Macro reutilizable mensualmente |
| Validación y cuadre contable | Copilot + fórmulas Excel | Reporte de calidad documentado |
| Trazabilidad del pipeline | Log_Transformaciones | Auditoría completa de cada paso |

### Conceptos Clave Reforzados

- **Copilot no se conecta directamente a Oracle**: actúa sobre datos ya disponibles en Excel, potenciando el análisis y la generación de código de transformación.
- **Power Query es el motor ETL nativo de Excel**: permite construir pipelines reproducibles con trazabilidad completa en el panel de "Pasos aplicados".
- **Los ceros a la izquierda en códigos de cuenta son críticos**: deben preservarse como texto desde el primer paso de importación para garantizar la integridad referencial.
- **La documentación del pipeline es tan importante como el pipeline mismo**: el `Log_Transformaciones` es la evidencia de auditoría que respalda la confiabilidad de los datos.
- **Las macros VBA generadas por Copilot son el puente hacia la automatización**: permiten que el equipo de contraloría ejecute procesos complejos de limpieza con un solo clic, sin conocimientos de programación.

### Recursos Adicionales

- [Microsoft Learn — Power Query: Conectar a una base de datos Oracle](https://learn.microsoft.com/es-es/power-query/connectors/oracle-database)
- [Microsoft Learn — Combinar consultas en Power Query](https://learn.microsoft.com/es-es/power-query/merge-queries-overview)
- [Microsoft Support — Introducción a Copilot en Excel](https://support.microsoft.com/es-es/office/introducción-a-copilot-en-excel-d7110502-0334-4b4f-a175-a73abdfc118a)
- [Oracle Docs — GL_JE_LINES y GL_JE_HEADERS: Referencia de tablas](https://docs.oracle.com/cd/E18727_01/doc.121/e13483/toc.htm)
- [Microsoft Learn — Introducción a VBA en Excel](https://learn.microsoft.com/es-es/office/vba/library-reference/concepts/getting-started-with-vba-in-office)

---

> 📋 **Próximo Laboratorio:** En el **Lab 03-00-02** aplicaremos los datos limpios generados en este pipeline para construir el cruce de cuentas contables, la detección automática de discrepancias y la generación del reporte de excepción para la dirección financiera, utilizando las capacidades avanzadas de análisis de Copilot en Excel.

---
LAB_END---
