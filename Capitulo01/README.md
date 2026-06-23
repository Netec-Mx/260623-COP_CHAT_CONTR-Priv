# Uso de Copilot para escribir, depurar y optimizar código VBA en Excel. Creación de automatizaciones personalizadas para tareas recurrentes sin conocimientos previos de programación.

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 60 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Aplicar (*Apply*) |
| **Módulo** | Módulo 1 — Automatización de tareas contables con Copilot y VBA |
| **Archivos requeridos** | `M1_Catalogo_Cuentas.xlsx`, `M1_Polizas_Enero.xlsx`, `M1_Balance_General.xlsx` |
| **Instructor** | Proporciona archivos de práctica antes del inicio del laboratorio |

---

## 2. Descripción General

En este laboratorio, los participantes utilizarán **Microsoft 365 Copilot integrado en Excel** para generar, depurar y optimizar código VBA orientado a tareas reales de contraloría corporativa, sin necesidad de conocimientos previos de programación. Se practicará el ciclo completo de desarrollo asistido por IA: redactar un prompt en lenguaje natural → obtener código VBA funcional → probarlo en el Editor de Visual Basic → identificar y corregir errores con ayuda de Copilot → optimizar el código final para grandes volúmenes de datos. Al concluir, cada participante habrá creado dos automatizaciones operativas aplicables directamente a su entorno de trabajo en contraloría.

---

## 3. Objetivos de Aprendizaje

Al finalizar este laboratorio, el participante será capaz de:

- [ ] Formular prompts efectivos en lenguaje natural para que Copilot genere código VBA funcional orientado a tareas contables recurrentes.
- [ ] Insertar, ejecutar y verificar macros VBA en el Editor de Visual Basic de Excel (*Alt + F11*).
- [ ] Identificar y corregir errores en código VBA generado por Copilot utilizando el mismo asistente como herramienta de depuración.
- [ ] Optimizar macros VBA para mejorar su rendimiento con grandes volúmenes de datos financieros, aplicando instrucciones de optimización dirigidas a Copilot.
- [ ] Crear al menos dos automatizaciones completas para tareas de contraloría: formateo automático de reportes contables y consolidación de hojas de pólizas.

---

## 4. Prerrequisitos

### Conocimientos previos

| Área | Nivel requerido |
|---|---|
| Manejo de Excel (hojas, celdas, fórmulas básicas) | Básico |
| Conceptos contables (pólizas, catálogo de cuentas, balance) | Básico-Intermedio |
| Navegación en la interfaz de Microsoft 365 | Básico |
| Conocimientos de programación VBA | **No requerido** |

### Acceso y licencias

| Recurso | Estado requerido |
|---|---|
| Cuenta corporativa Microsoft 365 activa | ✅ Confirmada |
| Licencia Microsoft 365 Copilot asignada | ✅ Habilitada en el tenant |
| Acceso al Editor VBA en Excel (`Alt + F11`) | ✅ Sin restricciones de grupo de políticas (GPO) |
| Archivos de práctica del Módulo 1 | ✅ Descargados y disponibles localmente |
| Conexión a Internet estable (mínimo 10 Mbps) | ✅ Activa |

> ⚠️ **Importante:** Si al abrir Excel no aparece el panel de Copilot en la cinta de opciones, contacta al administrador de TI antes de iniciar el laboratorio. Sin acceso a Copilot, los pasos de generación de código no pueden ejecutarse.

---

## 5. Entorno del Laboratorio

### Hardware recomendado

| Componente | Especificación mínima | Recomendado |
|---|---|---|
| Procesador | Intel Core i5 8ª gen / AMD Ryzen 5 | Intel Core i7 / AMD Ryzen 7 |
| RAM | 8 GB | 16 GB |
| Almacenamiento disponible | 5 GB libres | SSD con 10 GB libres |
| Resolución de pantalla | 1920×1080 | Monitor dual 1920×1080 |
| Conexión a Internet | 10 Mbps | 25 Mbps o superior |

### Software requerido

| Aplicación | Versión mínima | Verificación |
|---|---|---|
| Microsoft Excel | Microsoft 365 versión 2308 o superior | `Archivo → Cuenta → Acerca de Excel` |
| Microsoft 365 Copilot | Versión empresarial 2024 integrada | Panel lateral visible en Excel |
| Editor VBA (integrado en Excel) | Incluido en Excel 365 | `Alt + F11` abre el IDE |
| Navegador web | Edge o Chrome (versión más reciente) | Para consultas de soporte |

### Configuración inicial del entorno

Antes de comenzar los pasos del laboratorio, ejecuta las siguientes verificaciones:

**Verificación 1 — Confirmar versión de Excel:**
```
Archivo → Cuenta → Acerca de Excel
Verificar que la versión sea 2308 o superior y que el canal sea "Canal mensual"
```

**Verificación 2 — Confirmar que el panel de Copilot está disponible:**
```
En Excel: Inicio → buscar el botón "Copilot" en la cinta de opciones (extremo derecho)
Si no aparece: Archivo → Opciones → Complementos → verificar que Copilot no esté deshabilitado
```

**Verificación 3 — Habilitar la pestaña Programador en Excel (para acceso rápido al Editor VBA):**
```
Archivo → Opciones → Personalizar cinta de opciones → 
Marcar la casilla "Programador" en la columna derecha → Aceptar
```

**Verificación 4 — Abrir los archivos de práctica:**
```
Abrir los tres archivos proporcionados por el instructor:
- M1_Catalogo_Cuentas.xlsx
- M1_Polizas_Enero.xlsx
- M1_Balance_General.xlsx

Guardar una copia de trabajo de cada archivo con el sufijo _TRABAJO antes de comenzar:
Ejemplo: M1_Polizas_Enero_TRABAJO.xlsx
```

> 🔒 **Nota de seguridad:** Trabaja **siempre** sobre las copias `_TRABAJO`. Los archivos originales son tu respaldo en caso de error durante los ejercicios.

---

## 6. Procedimiento Paso a Paso

El laboratorio está organizado en **cuatro secciones**:

| Sección | Actividad | Tiempo estimado |
|---|---|---|
| A | Exploración del entorno y primer prompt VBA | 10 min |
| B | Automatización 1: Formateo automático de reporte contable | 15 min |
| C | Automatización 2: Consolidación de hojas de pólizas | 15 min |
| D | Depuración y optimización de código con Copilot | 20 min |

---

### Sección A — Exploración del Entorno y Primer Prompt VBA

**Objetivo de la sección:** Familiarizarse con el panel de Copilot en Excel y comprender cómo formular el primer prompt para generar código VBA básico.

---

#### Paso A.1 — Abrir el archivo de trabajo y activar el panel de Copilot

**Objetivo:** Confirmar que el entorno está correctamente configurado y que Copilot responde en el contexto del libro de Excel.

**Instrucciones:**

1. Abre el archivo `M1_Catalogo_Cuentas_TRABAJO.xlsx` (copia de trabajo que creaste en la configuración inicial).
2. Observa la estructura del archivo: debe contener una hoja llamada **`Catalogo`** con las columnas: `Cuenta`, `Descripcion`, `Tipo`, `Nivel`, `Saldo_Actual`.
3. Selecciona cualquier celda dentro del rango de datos.
4. Presiona `Ctrl + T` para convertir el rango en una **Tabla de Excel**. En el cuadro de diálogo, confirma que "La tabla tiene encabezados" esté marcado. Haz clic en **Aceptar**.
5. Haz clic derecho sobre el nombre de la tabla (aparece en la pestaña **Diseño de tabla** como `Tabla1`) y renómbrala como **`CatalogoCuentas`**.
6. En la cinta de opciones, haz clic en **Inicio** → busca el botón **Copilot** (ícono de estrella/chispa en el extremo derecho de la cinta) y haz clic para abrir el panel lateral.
7. Verifica que el panel de Copilot se abra en el lado derecho de la pantalla y muestre un campo de texto para ingresar prompts.

**Resultado esperado:** El panel de Copilot está abierto, la tabla `CatalogoCuentas` está definida y el archivo está listo para recibir macros.

**Verificación:** En la pestaña **Diseño de tabla**, confirma que el nombre de la tabla sea exactamente `CatalogoCuentas` (sin espacios, sin tildes). Este nombre será referenciado en los prompts posteriores.

---

#### Paso A.2 — Escribir el primer prompt de exploración

**Objetivo:** Comprender cómo Copilot interpreta instrucciones en lenguaje natural y genera código VBA básico.

**Instrucciones:**

1. En el campo de texto del panel de Copilot, escribe el siguiente prompt exactamente como se muestra:

```
Crea una macro VBA que cuente cuántas filas tiene la tabla CatalogoCuentas 
en la hoja Catalogo y muestre el resultado en un mensaje emergente.
```

2. Presiona **Enter** o haz clic en el botón de enviar (ícono de flecha).
3. Espera la respuesta de Copilot (generalmente entre 5 y 15 segundos).
4. Lee el código generado completo antes de copiarlo. Identifica visualmente:
   - La línea `Sub` que define el nombre de la macro.
   - La referencia a la tabla `CatalogoCuentas`.
   - La instrucción `MsgBox` que muestra el resultado.
5. Haz clic en el botón **Copiar código** (si está disponible en el panel) o selecciona manualmente el bloque de código y cópialo con `Ctrl + C`.
6. Abre el Editor de Visual Basic presionando `Alt + F11`.
7. En el Editor VBA, ve a **Insertar → Módulo** para crear un nuevo módulo en blanco.
8. Pega el código copiado con `Ctrl + V`.
9. Presiona `F5` o haz clic en el botón **Ejecutar** (triángulo verde) para ejecutar la macro.

**Resultado esperado:** Aparece un cuadro de mensaje (MsgBox) indicando el número de filas de la tabla `CatalogoCuentas`. El número debe coincidir con el total de cuentas en el catálogo del archivo de práctica.

**Verificación:** Cierra el cuadro de mensaje. En la hoja `Catalogo`, cuenta manualmente las filas de datos (sin incluir el encabezado) y confirma que el número mostrado por la macro sea idéntico.

---

### Sección B — Automatización 1: Formateo Automático de Reporte Contable

**Objetivo de la sección:** Crear una macro completa que aplique formato profesional a un reporte del Balance General, automatizando una tarea que normalmente toma 15-20 minutos de trabajo manual.

---

#### Paso B.1 — Preparar el archivo de Balance General

**Objetivo:** Estructurar correctamente el archivo de práctica para que Copilot pueda generar código VBA preciso con referencias correctas.

**Instrucciones:**

1. Abre el archivo `M1_Balance_General_TRABAJO.xlsx`.
2. Verifica que contenga una hoja llamada **`Balance`** con la siguiente estructura de columnas: `Cuenta`, `Descripcion`, `Saldo_Anterior`, `Movimientos`, `Saldo_Actual`, `Variacion_Pct`.
3. Si los datos no están en formato de Tabla de Excel, selecciona cualquier celda del rango y presiona `Ctrl + T`. Nombra la tabla como **`BalanceGeneral`**.
4. Observa el estado actual del formato: sin colores, sin bordes, sin distinción visual entre tipos de cuenta (Activo, Pasivo, Capital).
5. Regresa al panel de Copilot (si lo cerraste, vuelve a abrirlo desde la cinta de opciones → **Copilot**).

**Resultado esperado:** La hoja `Balance` contiene la tabla `BalanceGeneral` sin formato especial, lista para recibir la macro de formateo.

**Verificación:** En la pestaña **Diseño de tabla**, confirma que el nombre sea `BalanceGeneral`.

---

#### Paso B.2 — Generar la macro de formateo con Copilot

**Objetivo:** Utilizar un prompt enriquecido con contexto para obtener código VBA de formateo profesional.

**Instrucciones:**

1. En el panel de Copilot, escribe el siguiente prompt. Tómate el tiempo para leerlo completo antes de enviarlo:

```
Crea una macro VBA llamada FormatearBalanceGeneral que aplique el siguiente 
formato a la tabla BalanceGeneral en la hoja Balance:

1. Encabezados de columna: fondo azul oscuro (#003366), texto blanco, negrita, 
   fuente Calibri 11.
2. Filas donde la columna Descripcion contenga la palabra "TOTAL": 
   fondo gris claro (#D9D9D9), texto negrita.
3. Filas donde la columna Variacion_Pct sea menor que -0.05 (variación negativa 
   mayor al 5%): texto rojo y negrita en la columna Variacion_Pct.
4. Filas donde la columna Variacion_Pct sea mayor que 0.05 (variación positiva 
   mayor al 5%): texto verde oscuro en la columna Variacion_Pct.
5. Aplicar bordes delgados a todas las celdas de la tabla.
6. Ajustar automáticamente el ancho de todas las columnas.
7. Al finalizar, mostrar un mensaje que diga "Formato aplicado correctamente".
```

2. Presiona **Enter** y espera la respuesta de Copilot.
3. Revisa el código generado. Verifica que contenga:
   - Referencias a la tabla `BalanceGeneral` o a la hoja `Balance`.
   - Uso de colores mediante `RGB()` o códigos hexadecimales.
   - Un bucle `For` que recorra las filas.
   - La instrucción `MsgBox` al final.
4. Copia el código generado.
5. En el Editor VBA (`Alt + F11`), en el mismo módulo creado anteriormente (o en uno nuevo), pega el código.

**Resultado esperado:** El Editor VBA muestra el código de la macro `FormatearBalanceGeneral` completo, sin errores de sintaxis visibles (no aparecen líneas subrayadas en rojo en el editor).

**Verificación:** En el Editor VBA, ve a **Depurar → Compilar VBAProject**. Si no aparece ningún cuadro de error, el código compila correctamente.

---

#### Paso B.3 — Ejecutar y verificar la macro de formateo

**Objetivo:** Confirmar que la macro aplica el formato esperado y funciona correctamente sobre los datos del Balance General.

**Instrucciones:**

1. Posiciona el cursor dentro del código de `FormatearBalanceGeneral` en el Editor VBA.
2. Presiona `F5` para ejecutar la macro.
3. Regresa a la hoja `Balance` de Excel (minimiza o cierra el Editor VBA con `Alt + F11`).
4. Verifica visualmente que se hayan aplicado los siguientes elementos de formato:
   - [ ] Encabezados con fondo azul oscuro y texto blanco.
   - [ ] Filas de totales con fondo gris.
   - [ ] Valores negativos en rojo en la columna `Variacion_Pct`.
   - [ ] Valores positivos en verde en la columna `Variacion_Pct`.
   - [ ] Bordes visibles en todas las celdas.
   - [ ] Columnas con ancho ajustado automáticamente.
5. Confirma que apareció el mensaje "Formato aplicado correctamente".

**Resultado esperado:** El Balance General tiene formato profesional aplicado automáticamente. La tarea que tomaría 15-20 minutos manualmente se completó en segundos.

**Verificación:** Identifica al menos una fila con variación negativa y confirma que el texto en `Variacion_Pct` aparece en rojo. Identifica al menos una fila de total y confirma el fondo gris.

> 💡 **Punto de reflexión:** Anota en qué columna está `Variacion_Pct` en tu tabla. Si el código generado por Copilot usó un índice de columna numérico (ej. `.Cells(i, 6)`) en lugar del nombre de columna, verifica que ese número corresponda a la posición real de la columna en tu tabla. Si no coincide, continúa al Paso B.4 para corregirlo con Copilot.

---

#### Paso B.4 — Corregir un error de referencia de columna con Copilot (opcional/condicional)

**Objetivo:** Practicar la depuración asistida por Copilot cuando el código referencia columnas por posición en lugar de por nombre.

> Este paso se ejecuta **solo si** el formato no se aplicó correctamente en el Paso B.3 debido a que las columnas no están en el orden esperado.

**Instrucciones:**

1. En el panel de Copilot, pega el código problemático y escribe el siguiente prompt:

```
El siguiente código VBA usa índices numéricos de columna (.Cells(i, 6)) pero 
en mi tabla BalanceGeneral la columna Variacion_Pct está en la posición 6 
contando desde la primera columna de la tabla. Modifica el código para que 
use el nombre de columna directamente desde la tabla ListObject en lugar de 
índices numéricos, para que sea más robusto ante cambios de estructura.

[Pega aquí el código actual]
```

2. Copia el código corregido que genera Copilot.
3. En el Editor VBA, selecciona y elimina el código anterior de `FormatearBalanceGeneral`.
4. Pega el código corregido.
5. Ejecuta nuevamente con `F5` y verifica que el formato se aplique correctamente.

**Resultado esperado:** El código corregido usa referencias por nombre de columna, haciéndolo más robusto. El formato se aplica correctamente en todas las filas.

**Verificación:** Ejecuta la macro dos veces seguidas. El resultado debe ser idéntico en ambas ejecuciones (el formato se aplica limpiamente sin duplicar estilos).

---

### Sección C — Automatización 2: Consolidación de Hojas de Pólizas

**Objetivo de la sección:** Crear una macro que consolide automáticamente múltiples hojas de pólizas contables en una sola hoja de resumen, una tarea de alto valor en el cierre mensual de contraloría.

---

#### Paso C.1 — Explorar la estructura del archivo de pólizas

**Objetivo:** Comprender la estructura de datos antes de formular el prompt para que Copilot genere código preciso.

**Instrucciones:**

1. Abre el archivo `M1_Polizas_Enero_TRABAJO.xlsx`.
2. Observa que el archivo contiene **múltiples hojas**, cada una representando un tipo de póliza:
   - `Polizas_Ingresos`
   - `Polizas_Egresos`
   - `Polizas_Diario`
   - `Resumen` (hoja vacía donde se consolidará la información)
3. Verifica que cada hoja de pólizas tenga los mismos encabezados de columna: `Fecha`, `Numero_Poliza`, `Cuenta`, `Descripcion`, `Debe`, `Haber`, `Tipo_Poliza`.
4. Anota el nombre exacto de cada hoja (respetando mayúsculas y guiones bajos).
5. Abre el panel de Copilot desde la cinta de opciones.

**Resultado esperado:** Conoces la estructura exacta del archivo: tres hojas de origen con la misma estructura de columnas y una hoja de destino vacía llamada `Resumen`.

**Verificación:** Confirma que la hoja `Resumen` existe y está vacía. Si no existe, créala manualmente: clic derecho en una pestaña de hoja → **Insertar → Hoja de cálculo** → renombrar como `Resumen`.

---

#### Paso C.2 — Generar la macro de consolidación con Copilot

**Objetivo:** Crear un prompt que genere una macro de consolidación robusta, aplicando los principios de prompt efectivo aprendidos en la lección.

**Instrucciones:**

1. En el panel de Copilot, escribe el siguiente prompt:

```
Crea una macro VBA llamada ConsolidarPolizas que realice lo siguiente:

1. Recorra las hojas: Polizas_Ingresos, Polizas_Egresos y Polizas_Diario.
2. Copie todos los datos de cada hoja (sin incluir la fila de encabezados, 
   excepto en la primera hoja) a la hoja llamada Resumen.
3. Antes de copiar, limpie completamente el contenido de la hoja Resumen.
4. En la hoja Resumen, agrega una columna adicional llamada Fuente al final 
   de cada fila, que indique el nombre de la hoja de origen de cada registro.
5. Al finalizar, muestra un mensaje con el total de registros consolidados 
   y el desglose por hoja de origen (cuántos vinieron de cada una).
6. La macro debe manejar errores en caso de que alguna hoja no exista.
```

2. Envía el prompt y espera la respuesta de Copilot.
3. Lee el código completo. Verifica que contenga:
   - Un array o lista con los nombres de las tres hojas de origen.
   - Un bucle que itere sobre las hojas.
   - Lógica para copiar encabezados solo una vez.
   - Escritura del nombre de la hoja en la columna `Fuente`.
   - Manejo de errores (`On Error`).
   - Un `MsgBox` con el resumen.
4. Copia el código y pégalo en un nuevo módulo del Editor VBA (`Alt + F11` → **Insertar → Módulo**).

**Resultado esperado:** El Editor VBA contiene el código de `ConsolidarPolizas` completo. El código incluye al menos un bucle principal y referencias a las tres hojas de origen.

**Verificación:** Ejecuta **Depurar → Compilar VBAProject**. No deben aparecer errores de compilación.

---

#### Paso C.3 — Ejecutar la macro de consolidación y verificar resultados

**Objetivo:** Confirmar que la consolidación funciona correctamente y que los datos en la hoja `Resumen` son completos y precisos.

**Instrucciones:**

1. Posiciona el cursor dentro del código de `ConsolidarPolizas` en el Editor VBA.
2. Presiona `F5` para ejecutar la macro.
3. Si aparece algún error de ejecución (cuadro rojo con número de error), anota el número de error y el mensaje. Continúa al Paso C.4 para depurarlo con Copilot.
4. Si la macro se ejecuta sin errores, regresa a la hoja `Resumen` en Excel.
5. Verifica los siguientes puntos:
   - [ ] La hoja `Resumen` contiene una fila de encabezados con las columnas originales más la columna `Fuente`.
   - [ ] Los registros de las tres hojas están consolidados en orden secuencial.
   - [ ] La columna `Fuente` muestra correctamente el nombre de la hoja de origen para cada fila.
   - [ ] No hay filas de encabezado duplicadas en medio de los datos.
6. Cuenta manualmente los registros en cada hoja de origen y suma el total. Compara con el número mostrado en el mensaje de la macro.

**Resultado esperado:** La hoja `Resumen` contiene todos los registros de las tres hojas de pólizas consolidados, con la columna `Fuente` correctamente poblada. El total de registros en el mensaje coincide con la suma manual.

**Verificación:** En la hoja `Resumen`, usa la función `=CONTAR.SI(G:G,"Polizas_Ingresos")` (ajusta la letra de columna según donde esté `Fuente`) para contar los registros de cada fuente y compara con el desglose del mensaje de la macro.

---

#### Paso C.4 — Depurar un error de ejecución con Copilot (condicional)

**Objetivo:** Practicar el uso de Copilot como herramienta de depuración cuando una macro genera un error en tiempo de ejecución.

> Este paso se ejecuta **solo si** la macro generó un error en el Paso C.3. Si no hubo errores, lee el procedimiento como referencia y continúa a la Sección D.

**Instrucciones:**

1. Anota el número de error (ej. "Error 9: Subíndice fuera del intervalo" o "Error 1004: Error definido por la aplicación").
2. En el panel de Copilot, escribe el siguiente prompt adaptado al error que obtuviste:

```
Mi macro ConsolidarPolizas genera el siguiente error al ejecutarse:
[Error X: descripción del error]
El error ocurre en la línea: [copia aquí la línea donde se detuvo la ejecución]

El archivo tiene hojas llamadas exactamente: Polizas_Ingresos, Polizas_Egresos, 
Polizas_Diario y Resumen. ¿Cuál es la causa del error y cómo lo corrijo?
```

3. Lee la explicación de Copilot. Identifica la causa raíz del error.
4. Aplica la corrección sugerida en el código dentro del Editor VBA.
5. Ejecuta nuevamente con `F5`.
6. Repite este proceso hasta que la macro se ejecute sin errores.

**Resultado esperado:** La macro se ejecuta exitosamente después de aplicar la corrección sugerida por Copilot.

**Verificación:** Ejecuta la macro una segunda vez para confirmar que el resultado es consistente y que la hoja `Resumen` se limpia y repopula correctamente en cada ejecución.

---

### Sección D — Depuración y Optimización de Código con Copilot

**Objetivo de la sección:** Aplicar técnicas de optimización de rendimiento a las macros creadas, y practicar la explicación inversa de código VBA utilizando Copilot como herramienta de comprensión.

---

#### Paso D.1 — Optimizar la macro de consolidación para grandes volúmenes

**Objetivo:** Aplicar optimizaciones estándar de rendimiento en VBA utilizando Copilot, reduciendo el tiempo de ejecución para archivos con decenas de miles de registros.

**Instrucciones:**

1. En el panel de Copilot, escribe el siguiente prompt:

```
Optimiza la macro ConsolidarPolizas para mejorar su rendimiento cuando 
trabaje con archivos que contienen 50,000 registros o más. 
Específicamente:
1. Desactiva la actualización de pantalla (ScreenUpdating) durante la ejecución.
2. Desactiva el cálculo automático (Calculation) durante la ejecución.
3. Desactiva los eventos de Excel (EnableEvents) durante la ejecución.
4. Reactiva todas estas opciones al finalizar, incluso si ocurre un error.
5. Agrega también un contador de tiempo que muestre en el mensaje final 
   cuántos segundos tardó la ejecución.

[Pega aquí el código actual de ConsolidarPolizas]
```

2. Copia el código optimizado generado por Copilot.
3. En el Editor VBA, **reemplaza** el código anterior de `ConsolidarPolizas` con el código optimizado.
4. Verifica que el código optimizado contenga al inicio del procedimiento:
   ```vba
   Application.ScreenUpdating = False
   Application.Calculation = xlCalculationManual
   Application.EnableEvents = False
   ```
5. Verifica que al final (y también en el bloque de manejo de errores) contenga:
   ```vba
   Application.ScreenUpdating = True
   Application.Calculation = xlCalculationAutomatic
   Application.EnableEvents = True
   ```
6. Ejecuta la macro optimizada con `F5` y observa la diferencia en velocidad de ejecución.

**Resultado esperado:** La macro optimizada se ejecuta notablemente más rápido (la pantalla no parpadea durante la ejecución). El mensaje final incluye el tiempo de ejecución en segundos.

**Verificación:** Confirma que después de la ejecución, Excel responde normalmente (el cálculo automático y los eventos están reactivados). Para verificar: cambia el valor de una celda con una fórmula y confirma que el resultado se actualiza automáticamente.

---

#### Paso D.2 — Optimizar la macro de formateo para rendimiento

**Objetivo:** Aplicar el mismo patrón de optimización a la macro `FormatearBalanceGeneral` y comparar el resultado.

**Instrucciones:**

1. En el panel de Copilot, escribe:

```
Aplica las mismas optimizaciones de rendimiento (ScreenUpdating, Calculation, 
EnableEvents) a la macro FormatearBalanceGeneral. Además, modifica el bucle 
principal para que en lugar de aplicar formato celda por celda, use 
Union() para agrupar los rangos del mismo tipo de formato y aplicarlos 
todos de una vez al final del bucle. Esto reduce las llamadas al motor 
de renderizado de Excel.

[Pega aquí el código actual de FormatearBalanceGeneral]
```

2. Revisa el código generado. Identifica el uso de `Union()` para agrupar rangos.
3. Reemplaza el código anterior en el Editor VBA.
4. Ejecuta la macro optimizada sobre el archivo `M1_Balance_General_TRABAJO.xlsx`.

**Resultado esperado:** La macro de formateo aplica todos los colores y estilos de manera más eficiente. En archivos grandes, la diferencia de velocidad es perceptible.

**Verificación:** El formato del Balance General es idéntico al obtenido en el Paso B.3, pero la ejecución fue más rápida y sin parpadeo de pantalla.

---

#### Paso D.3 — Usar Copilot para explicar código VBA heredado

**Objetivo:** Practicar la capacidad de "explicación inversa" de Copilot para comprender macros existentes, una habilidad clave en auditoría de procesos.

**Instrucciones:**

1. El instructor ha incluido en el archivo `M1_Polizas_Enero_TRABAJO.xlsx` una macro oculta llamada `MacroHeredada` en un módulo llamado `Modulo_Legado`. Ábrela en el Editor VBA.

> Si el archivo no incluye esta macro, el instructor la proporcionará por separado, o puedes usar el siguiente código de ejemplo:

```vba
Sub MacroHeredada()
    Dim ws As Worksheet
    Dim lr As Long
    Dim i As Long
    Dim suma As Double
    Set ws = ThisWorkbook.Sheets(1)
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If ws.Cells(i, 5).Value > 0 And ws.Cells(i, 6).Value = 0 Then
            ws.Cells(i, 7).Value = "REVISAR"
            ws.Cells(i, 7).Interior.Color = RGB(255, 255, 0)
            suma = suma + ws.Cells(i, 5).Value
        End If
    Next i
    ws.Cells(lr + 2, 5).Value = suma
    ws.Cells(lr + 2, 5).Font.Bold = True
End Sub
```

2. Copia el código de `MacroHeredada`.
3. En el panel de Copilot, escribe el siguiente prompt:

```
Explica en español, sección por sección, qué hace el siguiente código VBA. 
Describe el propósito de cada bloque de instrucciones y qué resultado 
produce sobre los datos de Excel. Usa lenguaje accesible para un contador 
sin conocimientos de programación.

[Pega aquí el código de MacroHeredada]
```

4. Lee la explicación generada por Copilot.
5. Verifica que la explicación mencione:
   - Qué columnas evalúa la macro (columna 5 = `Debe`, columna 6 = `Haber` según la estructura del archivo).
   - Qué condición determina que una fila se marque como "REVISAR".
   - Qué hace con la suma al final.

**Resultado esperado:** Copilot genera una explicación clara en español de cada sección del código, permitiendo al participante comprender completamente qué hace la macro sin necesidad de conocer VBA.

**Verificación:** Sin ver el código, explica oralmente (o por escrito en una celda de Excel) qué hace `MacroHeredada` usando solo la explicación de Copilot. Si puedes describirla correctamente, el objetivo se cumplió.

---

#### Paso D.4 — Solicitar mejoras a la macro heredada

**Objetivo:** Aplicar el ciclo completo de optimización: tomar código existente, mejorarlo con Copilot y verificar que el resultado es superior.

**Instrucciones:**

1. En el panel de Copilot, escribe:

```
Mejora la macro MacroHeredada con los siguientes cambios:
1. Reemplaza los índices numéricos de columna (.Cells(i,5), .Cells(i,6)) 
   por referencias a nombres de columna de la tabla, para que sea más 
   fácil de mantener.
2. Agrega manejo de errores para que si la hoja no existe o está vacía, 
   muestre un mensaje amigable en lugar de fallar silenciosamente.
3. Agrega un comentario explicativo en español al inicio de cada sección 
   del código.
4. Renombra la macro como ValidarPolizasSinContrapartida.

[Pega aquí el código de MacroHeredada]
```

2. Copia el código mejorado.
3. Agrégalo como un nuevo procedimiento en el módulo del Editor VBA.
4. Ejecuta `ValidarPolizasSinContrapartida` sobre el archivo de pólizas.
5. Verifica que las filas donde `Debe > 0` y `Haber = 0` queden marcadas en amarillo con la etiqueta "REVISAR".

**Resultado esperado:** La macro mejorada es más legible, robusta y autodocumentada. Produce el mismo resultado funcional que la macro original pero con mejor estructura de código.

**Verificación:** Abre el código de `ValidarPolizasSinContrapartida` en el Editor VBA y confirma que los comentarios en español están presentes y son descriptivos.

---

## 7. Validación y Pruebas Finales

Al concluir todas las secciones, realiza las siguientes verificaciones de cierre para confirmar que los objetivos del laboratorio se cumplieron:

| # | Verificación | Criterio de éxito | ✅ |
|---|---|---|---|
| 1 | La tabla `CatalogoCuentas` existe y está correctamente nombrada | Nombre visible en pestaña **Diseño de tabla** | ☐ |
| 2 | La macro `FormatearBalanceGeneral` ejecuta sin errores | No aparece cuadro de error rojo; formato aplicado visualmente | ☐ |
| 3 | El Balance General muestra filas de totales en gris y variaciones en rojo/verde | Verificación visual en hoja `Balance` | ☐ |
| 4 | La macro `ConsolidarPolizas` consolida las tres hojas en `Resumen` | Total de filas en `Resumen` = suma de filas en las tres hojas de origen | ☐ |
| 5 | La columna `Fuente` en `Resumen` está correctamente poblada | `=CONTAR.SI()` devuelve el número correcto por hoja | ☐ |
| 6 | Las macros optimizadas incluyen `ScreenUpdating = False/True` | Visible en el código en el Editor VBA | ☐ |
| 7 | `ValidarPolizasSinContrapartida` marca correctamente las filas sin contrapartida | Filas con `Debe > 0` y `Haber = 0` aparecen en amarillo con "REVISAR" | ☐ |
| 8 | Copilot generó una explicación comprensible de `MacroHeredada` | El participante puede describir la función de la macro sin ver el código | ☐ |

**Criterio de aprobación del laboratorio:** Completar al menos 6 de las 8 verificaciones con éxito.

---

## 8. Resolución de Problemas

### Problema 1 — El panel de Copilot no aparece en la cinta de opciones de Excel

**Síntoma:** Al abrir Excel y buscar el botón "Copilot" en la cinta de opciones, el botón no está visible o aparece atenuado (gris, no interactivo).

**Causa probable:** Una o más de las siguientes condiciones:
- La licencia de Microsoft 365 Copilot no está asignada al usuario en el portal de administración de Microsoft 365.
- Excel no está actualizado al **Canal mensual** de Microsoft 365 Apps (el Canal semestral puede tener funcionalidades de Copilot con retraso de hasta 6 meses).
- El tenant corporativo tiene Copilot habilitado, pero el usuario específico no tiene la licencia asignada individualmente.
- Una política de grupo (GPO) corporativa está deshabilitando la funcionalidad de IA en las aplicaciones de Office.

**Solución paso a paso:**

1. Verifica la versión de Excel: `Archivo → Cuenta → Acerca de Excel`. La versión debe ser **2308 o superior** y el canal debe decir "Canal mensual" o "Canal mensual (dirigido)".
2. Si el canal es "Canal semestral empresarial", contacta al administrador de TI para solicitar el cambio de canal o la actualización manual.
3. Verifica la licencia: accede a `https://portal.office.com` con tu cuenta corporativa → haz clic en tu perfil → **Ver cuenta** → **Suscripciones**. Confirma que "Microsoft 365 Copilot" aparece como licencia activa.
4. Si la licencia está activa pero el botón no aparece, cierra Excel completamente y vuelve a abrirlo. Si persiste, ejecuta `Inicio → Ejecutar → winword /safe` para abrir en modo seguro y verificar si hay un complemento bloqueando la funcionalidad.
5. Como última opción, contacta al administrador de TI corporativo para verificar que las URLs de Copilot no estén bloqueadas en el firewall: `*.microsoft.com`, `*.office.com`, `*.copilot.microsoft.com`.

---

### Problema 2 — La macro se ejecuta pero produce resultados incorrectos (columnas desfasadas o datos en celdas equivocadas)

**Síntoma:** La macro se ejecuta sin error, pero los datos se copian en columnas incorrectas, los colores se aplican a las filas equivocadas, o el conteo de registros no coincide con el esperado. Por ejemplo, la columna `Fuente` aparece en la columna A en lugar de al final, o el formato rojo se aplica a filas que no tienen variación negativa.

**Causa probable:** El código generado por Copilot usó **índices numéricos de columna** (ej. `.Cells(i, 3)`) basados en una suposición sobre el orden de las columnas, pero el orden real en el archivo de práctica es diferente. Esto ocurre cuando el prompt no especificó explícitamente los nombres de las columnas o cuando la tabla no estaba correctamente definida como Tabla de Excel al momento de escribir el prompt.

**Solución paso a paso:**

1. Abre el Editor VBA (`Alt + F11`) y localiza las líneas que usan índices numéricos de columna (busca patrones como `.Cells(i, 4)` o `.Range("D" & i)`).
2. Regresa a la hoja de Excel y verifica el orden real de las columnas contando desde la izquierda (columna 1 = A, columna 2 = B, etc.).
3. En el panel de Copilot, escribe el siguiente prompt de corrección:

```
El código generado usa índices numéricos de columna que no coinciden con 
mi archivo. En mi tabla [nombre de la tabla], las columnas están en este orden:
Columna 1: [nombre], Columna 2: [nombre], Columna 3: [nombre]...
Por favor, reescribe el código usando los nombres de columna directamente 
(accediendo a ellas por nombre desde el ListObject) en lugar de índices numéricos.
```

4. Reemplaza el código problemático con el código corregido que genera Copilot.
5. Ejecuta nuevamente y verifica que los resultados sean correctos.
6. **Prevención futura:** Siempre incluye el orden explícito de las columnas en el prompt inicial, o mejor aún, convierte los datos en Tabla de Excel con nombres descriptivos antes de escribir el prompt, para que Copilot pueda inferir la estructura correctamente.

---

## 9. Limpieza del Entorno

Al finalizar el laboratorio, realiza los siguientes pasos para dejar el entorno en orden:

1. **Guardar el trabajo:**
   - En cada archivo `_TRABAJO`, guarda los cambios: `Ctrl + S`.
   - Si Excel pregunta sobre el formato (`.xlsx` vs `.xlsm` para macros), elige **Guardar como libro de Excel habilitado para macros (`.xlsm`)** para conservar las macros creadas.

2. **Exportar el código VBA como respaldo:**
   - En el Editor VBA (`Alt + F11`), haz clic derecho sobre cada módulo en el panel izquierdo.
   - Selecciona **Exportar archivo** y guarda el archivo `.bas` en tu carpeta de trabajo.
   - Esto crea un respaldo del código independiente del archivo Excel.

3. **Cerrar el Editor VBA:**
   - Cierra el Editor VBA con `Alt + F11` o haciendo clic en la X del Editor.

4. **Verificar que los archivos originales no fueron modificados:**
   - Abre brevemente `M1_Polizas_Enero.xlsx` (sin el sufijo `_TRABAJO`) y confirma que no contiene las macros creadas durante el laboratorio.
   - Si por error se modificó el archivo original, notifica al instructor para obtener una copia limpia.

5. **Cerrar los archivos de trabajo:**
   - Cierra los tres archivos `_TRABAJO` en Excel.

6. **Opcional — Limpiar el historial del panel de Copilot:**
   - En el panel de Copilot, busca la opción de limpiar el historial de conversación si la política corporativa de privacidad lo requiere.

---

## 10. Resumen y Recursos Adicionales

### Resumen del Laboratorio

En este laboratorio completaste el ciclo completo de desarrollo de macros VBA asistido por IA:

| Habilidad desarrollada | Sección | Herramienta clave |
|---|---|---|
| Formulación de prompts efectivos en lenguaje natural | A, B, C | Panel de Copilot en Excel |
| Inserción y ejecución de código VBA | A, B, C | Editor VBA (`Alt + F11`) |
| Creación de automatización de formateo contable | B | Copilot + VBA |
| Creación de automatización de consolidación de hojas | C | Copilot + VBA |
| Depuración de errores asistida por IA | C.4, D | Copilot como debugger |
| Optimización de rendimiento para grandes volúmenes | D.1, D.2 | Copilot + VBA |
| Explicación inversa de código heredado | D.3 | Copilot como documentador |
| Mejora de código existente | D.4 | Copilot como refactor |

### Conceptos Clave para Recordar

- **El contexto mejora el código:** Tablas de Excel con nombres descriptivos + encabezados claros = prompts más precisos = código más confiable.
- **Anatomía del prompt efectivo:** Acción + Objeto + Condición/Resultado.
- **Revisar siempre antes de ejecutar:** El código de Copilot es un punto de partida, no un producto final garantizado.
- **Optimización estándar VBA:** `ScreenUpdating`, `Calculation` y `EnableEvents` son las tres líneas que más impacto tienen en el rendimiento.
- **Copilot como herramienta de auditoría:** La explicación inversa de código permite comprender macros heredadas sin conocimientos de programación.

### Recursos Adicionales

| Recurso | URL |
|---|---|
| Documentación oficial de Microsoft Copilot en Excel | https://learn.microsoft.com/es-es/copilot/microsoft-365/microsoft-365-copilot-overview |
| Referencia del lenguaje VBA para Office | https://learn.microsoft.com/es-es/office/vba/api/overview/excel |
| Guía de introducción a macros en Excel | https://support.microsoft.com/es-es/office/crear-o-eliminar-una-macro-974ef220-f716-4e01-b015-3ea70e64937b |
| Mejores prácticas de rendimiento en VBA | https://learn.microsoft.com/es-es/office/vba/excel/concepts/excel-performance/excel-improving-calcuation-performance |
| Microsoft 365 Copilot — Blog oficial | https://www.microsoft.com/es-es/microsoft-365/blog/2023/03/16/introducing-microsoft-365-copilot-a-whole-new-way-to-work/ |

### Preparación para la Siguiente Práctica

En el **Laboratorio 01-00-02** profundizarás en la automatización de flujos de trabajo completos utilizando Copilot sin necesidad de interactuar con el Editor VBA. Aprenderás a orquestar procesos desde la importación de datos hasta la generación de reportes utilizando exclusivamente instrucciones en lenguaje natural en el panel de Copilot, haciendo la automatización accesible para todo el equipo de contraloría independientemente de su perfil técnico.

---
*Laboratorio desarrollado para el curso: Microsoft 365 Copilot para Automatización Contable y Financiera — Módulo 1*
*Versión: 1.0 | Nivel: Aplicar (Bloom) | Duración: 60 minutos*
