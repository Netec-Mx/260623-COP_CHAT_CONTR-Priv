---LAB_START---
LAB_ID: 02-00-01
---MARKDOWN---
# Diseño de flujos de trabajo para cruzar cuentas contables, identificar discrepancias automáticamente y generar reportes de excepción utilizando IA.

---

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 90 minutos |
| **Complejidad** | Alta (Hard) |
| **Nivel Bloom** | Aplicar (Apply) |
| **Módulo** | Módulo 2 — Flujos de Trabajo Inteligentes con Copilot |
| **Versión de revisión** | 1.0 |
| **Instructor responsable** | Asignado por organización |

---

## 2. Descripción General

En este laboratorio aplicarás los conceptos de la Lección 2.1 para diseñar e implementar un flujo de trabajo completo de cruce de cuentas contables asistido por Microsoft 365 Copilot en Excel. Trabajarás con tres conjuntos de datos simulados que representan el libro mayor, el auxiliar de cuentas por pagar (CxP) y el auxiliar de cuentas por cobrar (CxC) de una empresa manufacturera ficticia, los cuales contienen discrepancias intencionalmente insertadas. Utilizarás el panel de Copilot para generar fórmulas avanzadas, ejecutar análisis comparativos, identificar excepciones y producir un reporte de excepción estructurado con narrativa ejecutiva. Al concluir, habrás construido un artefacto reutilizable que puede adaptarse a los procesos reales de conciliación de tu área de contraloría.

---

## 3. Objetivos de Aprendizaje

Al finalizar este laboratorio serás capaz de:

- [ ] **Diseñar** un flujo de trabajo estructurado en Excel con asistencia de Copilot que automatice el cruce de cuentas entre el libro mayor y los auxiliares contables (CxP y CxC).
- [ ] **Configurar y ejecutar** análisis asistidos por Copilot para identificar discrepancias, saldos descuadrados y registros sin contraparte en los datos contables.
- [ ] **Aplicar** fórmulas avanzadas de Excel (BUSCARV / XLOOKUP, COINCIDIR, SUMAR.SI.CONJUNTO, SI.ERROR) sugeridas por Copilot para soportar el proceso de validación.
- [ ] **Generar** un reporte de excepción dinámico y estructurado que clasifique y priorice las discrepancias encontradas, incluyendo una narrativa ejecutiva producida por Copilot.

---

## 4. Prerrequisitos

### 4.1 Conocimientos previos

| Área | Nivel requerido |
|---|---|
| Uso básico de Copilot en Excel (panel lateral, instrucciones en lenguaje natural) | Básico — cubierto en Práctica 1, Módulo 1 |
| Conceptos contables: libro mayor, auxiliares, cuentas de balance, conciliación | Básico — se asume familiaridad operativa |
| Tablas de Excel (Insertar → Tabla, rangos con nombre) | Básico |
| Fórmulas de búsqueda: BUSCARV o XLOOKUP | Básico |
| Tablas dinámicas y formato condicional | Básico |

### 4.2 Acceso y archivos necesarios

| Elemento | Estado requerido |
|---|---|
| Licencia Microsoft 365 Copilot activa y verificada | ✅ Confirmado antes del laboratorio |
| Microsoft Excel 365 versión 2308 o superior, canal mensual | ✅ Actualizado |
| Copilot habilitado en el panel lateral de Excel | ✅ Visible en pestaña Inicio |
| Archivo `M2_Libro_Mayor.xlsx` (proporcionado por instructor) | ✅ Descargado en escritorio |
| Archivo `M2_Auxiliar_CxP.xlsx` (proporcionado por instructor) | ✅ Descargado en escritorio |
| Archivo `M2_Auxiliar_CxC.xlsx` (proporcionado por instructor) | ✅ Descargado en escritorio |
| Archivo `M2_Plantilla_Reporte_Excepcion.xlsx` (proporcionado por instructor) | ✅ Descargado en escritorio |
| Conexión a Internet activa (mínimo 10 Mbps) | ✅ Verificada |

> ⚠️ **Aviso importante:** Todos los archivos de práctica contienen datos financieros **simulados y ficticios**. No cargues datos reales de tu organización en el entorno de Copilot sin autorización del área de seguridad de la información.

---

## 5. Entorno del Laboratorio

### 5.1 Hardware recomendado

| Componente | Especificación mínima |
|---|---|
| Procesador | Intel Core i5 8ª gen o AMD Ryzen 5 equivalente |
| RAM | 8 GB (se recomienda 16 GB para archivos con >10,000 filas) |
| Almacenamiento disponible | 1 GB libre para archivos de trabajo generados |
| Pantalla | 1920×1080 mínimo; monitor dual recomendado |
| Conectividad | Internet estable ≥ 10 Mbps |

### 5.2 Software requerido

| Aplicación | Versión |
|---|---|
| Microsoft Excel | Microsoft 365, versión 2308 o superior |
| Microsoft 365 Copilot | Versión empresarial integrada (2024) |
| Navegador web | Microsoft Edge o Google Chrome (versión actual) |

### 5.3 Configuración inicial del entorno

Antes de iniciar los pasos del laboratorio, realiza las siguientes verificaciones:

**Paso A — Verificar que Copilot está habilitado en Excel:**

1. Abre Microsoft Excel.
2. En la pestaña **Inicio**, confirma que el botón **Copilot** (ícono de constelación/estrella) está visible en la barra de herramientas.
3. Haz clic en **Copilot** para abrir el panel lateral. Debe aparecer el campo de chat con el mensaje de bienvenida de Copilot.
4. Si el botón no aparece, contacta a tu administrador de TI antes de continuar.

**Paso B — Organizar los archivos de práctica:**

1. Crea una carpeta en tu escritorio llamada `Lab02_CruceContable`.
2. Copia los cuatro archivos proporcionados por el instructor a esta carpeta:
   - `M2_Libro_Mayor.xlsx`
   - `M2_Auxiliar_CxP.xlsx`
   - `M2_Auxiliar_CxC.xlsx`
   - `M2_Plantilla_Reporte_Excepcion.xlsx`

**Paso C — Guardar el archivo de trabajo principal:**

1. Abre `M2_Libro_Mayor.xlsx`.
2. Guarda inmediatamente una copia con el nombre `Lab02_[TuNombre]_CruceContable.xlsx` en la carpeta `Lab02_CruceContable`.
3. Todos los pasos del laboratorio se realizarán sobre esta copia de trabajo.

---

## 6. Procedimiento Paso a Paso

El laboratorio está organizado en **cinco fases** que siguen el flujo de trabajo inteligente descrito en la Lección 2.1:

| Fase | Nombre | Tiempo estimado |
|---|---|---|
| Fase 1 | Carga y normalización de datos | 15 min |
| Fase 2 | Diseño del cruce de cuentas con Copilot | 20 min |
| Fase 3 | Identificación de discrepancias | 20 min |
| Fase 4 | Generación del reporte de excepción | 20 min |
| Fase 5 | Narrativa ejecutiva con Copilot | 15 min |

---

### Fase 1: Carga y Normalización de Datos

**Objetivo de la fase:** Consolidar las tres fuentes de datos en un único libro de trabajo y estructurarlas como Tablas de Excel con nombres descriptivos, garantizando uniformidad de formatos para que Copilot pueda operar con precisión.

---

#### Paso 1.1 — Importar el Auxiliar de CxP al libro de trabajo principal

**Objetivo:** Incorporar los datos del auxiliar de cuentas por pagar al libro de trabajo principal.

**Instrucciones:**

1. Con `Lab02_[TuNombre]_CruceContable.xlsx` abierto, ve a la pestaña **Hoja2** (o inserta una hoja nueva si no existe) y renómbrala haciendo doble clic sobre la pestaña: escribe `Auxiliar_CxP` y presiona **Enter**.
2. Abre el archivo `M2_Auxiliar_CxP.xlsx` desde la carpeta `Lab02_CruceContable`.
3. Selecciona todos los datos del archivo auxiliar (presiona **Ctrl + Shift + End** para seleccionar hasta la última celda con datos, luego ajusta desde A1).
4. Copia los datos (**Ctrl + C**).
5. Regresa a `Lab02_[TuNombre]_CruceContable.xlsx`, haz clic en la celda **A1** de la hoja `Auxiliar_CxP` y pega (**Ctrl + V**).
6. Repite los pasos 1–5 para el archivo `M2_Auxiliar_CxC.xlsx`, creando una hoja llamada `Auxiliar_CxC`.

**Resultado esperado:** El libro de trabajo debe contener al menos tres hojas: `Libro_Mayor`, `Auxiliar_CxP` y `Auxiliar_CxC`, cada una con sus datos originales.

**Verificación:** Confirma que las tres hojas tienen datos y que la fila 1 de cada hoja contiene encabezados de columna (no datos).

---

#### Paso 1.2 — Convertir los rangos de datos en Tablas de Excel con nombres descriptivos

**Objetivo:** Estructurar cada conjunto de datos como una Tabla oficial de Excel para que Copilot pueda referenciarlas con precisión.

**Instrucciones:**

1. Ve a la hoja `Libro_Mayor`. Haz clic en cualquier celda dentro del rango de datos.
2. Presiona **Ctrl + T** (o ve a **Insertar → Tabla**). Confirma que la casilla **"La tabla tiene encabezados"** está marcada. Haz clic en **Aceptar**.
3. Con la tabla seleccionada, ve a la pestaña **Diseño de tabla** (aparece en la cinta al seleccionar la tabla). En el campo **Nombre de tabla** (extremo izquierdo de la cinta), borra el nombre automático y escribe: `TablaLibroMayor`. Presiona **Enter**.
4. Repite los pasos 1–3 para la hoja `Auxiliar_CxP`, asignando el nombre `TablaAuxCxP`.
5. Repite los pasos 1–3 para la hoja `Auxiliar_CxC`, asignando el nombre `TablaAuxCxC`.

**Resultado esperado:** Tres tablas con nombres descriptivos, visibles en el **Administrador de nombres** (pestaña **Fórmulas → Administrador de nombres**).

**Verificación:**
- Ve a **Fórmulas → Administrador de nombres**.
- Confirma que aparecen `TablaLibroMayor`, `TablaAuxCxP` y `TablaAuxCxC` en la lista.
- Cierra el Administrador de nombres.

---

#### Paso 1.3 — Normalizar formatos de fecha e importe con asistencia de Copilot

**Objetivo:** Garantizar que las columnas de fecha y monto tienen formato uniforme en las tres tablas, condición indispensable para el cruce correcto.

**Instrucciones:**

1. Ve a la hoja `Libro_Mayor` y selecciona la columna de fechas (identifica el encabezado, por ejemplo `Fecha_Poliza`).
2. Abre el panel de Copilot haciendo clic en **Inicio → Copilot**.
3. En el campo de chat de Copilot, escribe la siguiente instrucción:

   > *"En la tabla TablaLibroMayor, verifica si la columna Fecha_Poliza tiene formato de fecha uniforme DD/MM/AAAA. Si hay celdas con texto o formato inconsistente, sugiere cómo corregirlas."*

4. Revisa la respuesta de Copilot. Si sugiere una fórmula de conversión (por ejemplo usando `FECHANUMERO` o `TEXTO`), aplícala en una columna auxiliar temporal llamada `Fecha_Normalizada`.
5. Repite el proceso para las columnas de importe: escribe en Copilot:

   > *"En TablaLibroMayor, confirma que la columna Monto contiene solo valores numéricos sin símbolos de moneda ni separadores de texto. Si hay inconsistencias, muéstrame cómo limpiarlas."*

6. Aplica las correcciones sugeridas si las hay. Repite la verificación para `TablaAuxCxP` y `TablaAuxCxC`.

> 💡 **Nota:** Los archivos de práctica están diseñados para tener formatos mayormente limpios, pero pueden incluir algunas celdas con formato texto intencionalmente para simular problemas reales de exportación. Copilot debe detectarlas.

**Resultado esperado:** Las tres tablas tienen columnas de fecha en formato de fecha de Excel (no texto) y columnas de monto con valores numéricos puros.

**Verificación:** Selecciona una celda de fecha en cada tabla. En la barra inferior de Excel (esquina derecha), confirma que el **Promedio** y la **Suma** se calculan correctamente, lo que indica que los valores son numéricos/fecha y no texto.

---

### Fase 2: Diseño del Cruce de Cuentas con Copilot

**Objetivo de la fase:** Crear la hoja de cruce central donde se compararán el libro mayor contra los auxiliares, utilizando fórmulas avanzadas generadas por Copilot.

---

#### Paso 2.1 — Crear la hoja de trabajo de cruce

**Objetivo:** Preparar el espacio de trabajo donde se ejecutará el cruce de cuentas.

**Instrucciones:**

1. Inserta una nueva hoja en el libro. Haz clic derecho sobre una pestaña de hoja → **Insertar → Hoja de cálculo**. Renómbrala como `Cruce_LM_CxP`.
2. En la celda **A1** de esta nueva hoja, escribe el encabezado del área de trabajo:
   ```
   CRUCE: LIBRO MAYOR vs. AUXILIAR CxP — MAYO 2024
   ```
3. En la fila **3**, crea los siguientes encabezados de columna (una por celda, comenzando en A3):

   | Columna | Encabezado |
   |---|---|
   | A | No_Poliza_LM |
   | B | Cuenta_LM |
   | C | Fecha_LM |
   | D | Monto_LM |
   | E | Referencia_LM |
   | F | No_Documento_CxP |
   | G | Monto_CxP |
   | H | Diferencia |
   | I | Estado_Cruce |
   | J | Prioridad |

4. Guarda el archivo (**Ctrl + S**).

**Resultado esperado:** Una hoja vacía con estructura de encabezados lista para recibir los datos del cruce.

**Verificación:** Confirma que los 10 encabezados están en la fila 3 y que la hoja se llama exactamente `Cruce_LM_CxP`.

---

#### Paso 2.2 — Poblar las columnas del libro mayor en la hoja de cruce

**Objetivo:** Traer los registros del libro mayor a la hoja de cruce como base del análisis comparativo.

**Instrucciones:**

1. Haz clic en la celda **A4** de la hoja `Cruce_LM_CxP`.
2. Abre el panel de Copilot y escribe:

   > *"Necesito traer todos los registros de TablaLibroMayor que correspondan a cuentas de cuentas por pagar (identifica las cuentas CxP por el prefijo de cuenta o por el campo que las clasifique) a la hoja Cruce_LM_CxP, comenzando en la celda A4. Las columnas que necesito son: número de póliza, cuenta contable, fecha, monto y referencia. Sugiere la fórmula o método más eficiente para hacer esto."*

3. Copilot sugerirá un enfoque. Las opciones más comunes serán:
   - **Opción A (recomendada):** Usar `FILTRAR` para traer dinámicamente los registros filtrados por tipo de cuenta.
   - **Opción B:** Copiar y pegar valores filtrados manualmente.

4. Si Copilot sugiere la función `FILTRAR`, aplica la fórmula en A4. El resultado será un rango dinámico con todos los registros de CxP del libro mayor. Ejemplo de estructura que Copilot puede generar:

   ```excel
   =FILTRAR(
     TablaLibroMayor[[No_Poliza]:[Referencia]],
     TablaLibroMayor[Tipo_Cuenta]="CxP",
     "Sin registros"
   )
   ```

   > *Nota: El nombre exacto de las columnas dependerá de los encabezados en tu archivo de práctica. Ajusta según lo que Copilot identifique.*

5. Confirma que los datos aparecen correctamente en la hoja de cruce a partir de la fila 4.

**Resultado esperado:** Las columnas A–E de la hoja `Cruce_LM_CxP` muestran los registros de CxP provenientes del libro mayor.

**Verificación:** Cuenta el número de filas con datos en la columna A. Anota este número (por ejemplo: "El libro mayor tiene **X** registros de CxP"). Lo necesitarás para validar el cruce en la Fase 3.

---

#### Paso 2.3 — Generar la fórmula de búsqueda para el Auxiliar CxP con Copilot

**Objetivo:** Usar Copilot para generar la fórmula XLOOKUP/BUSCARX que busque cada registro del libro mayor en el auxiliar de CxP y traiga el monto correspondiente.

**Instrucciones:**

1. Haz clic en la celda **F4** de la hoja `Cruce_LM_CxP`.
2. En el panel de Copilot, escribe:

   > *"En la hoja Cruce_LM_CxP, necesito una fórmula en la columna F (No_Documento_CxP) que busque el valor de la columna E (Referencia_LM) dentro de la columna de referencia de TablaAuxCxP y devuelva el número de documento correspondiente. Si no encuentra coincidencia, debe mostrar el texto 'SIN CONTRAPARTE'. Usa BUSCARX o XLOOKUP."*

3. Copilot generará una fórmula similar a:

   ```excel
   =SI.ERROR(
     BUSCARX(
       E4,
       TablaAuxCxP[Referencia_Doc],
       TablaAuxCxP[No_Documento],
       "SIN CONTRAPARTE",
       0
     ),
     "SIN CONTRAPARTE"
   )
   ```

4. Aplica la fórmula en F4 y extiéndela hacia abajo para cubrir todos los registros (doble clic en el controlador de relleno o arrastra hacia abajo).

5. Ahora haz clic en la celda **G4** y pide a Copilot:

   > *"En la celda G4 de Cruce_LM_CxP, crea una fórmula que busque el mismo valor de la columna E (Referencia_LM) en TablaAuxCxP y devuelva el monto del auxiliar. Si no hay coincidencia, devuelve 0."*

6. Aplica la fórmula generada en G4 y extiéndela hacia abajo.

**Resultado esperado:** Las columnas F y G muestran el número de documento y el monto del auxiliar CxP para cada registro del libro mayor. Las celdas sin coincidencia muestran "SIN CONTRAPARTE" en F y 0 en G.

**Verificación:** Revisa visualmente 5 filas al azar. Para cada una, verifica manualmente en la hoja `Auxiliar_CxP` que la referencia existe o no, y confirma que la fórmula devuelve el resultado correcto.

---

#### Paso 2.4 — Calcular la diferencia y clasificar el estado del cruce

**Objetivo:** Calcular la columna de diferencia y asignar automáticamente el estado de cada registro (Conciliado, Diferencia, Sin Contraparte).

**Instrucciones:**

1. Haz clic en la celda **H4** y escribe la fórmula de diferencia:

   ```excel
   =D4-G4
   ```

   Extiende esta fórmula hacia abajo para todas las filas con datos.

2. Haz clic en la celda **I4**. En el panel de Copilot, escribe:

   > *"En la celda I4 de la hoja Cruce_LM_CxP, crea una fórmula que evalúe las siguientes condiciones y asigne un estado: Si F4 es 'SIN CONTRAPARTE', el estado es 'SIN CONTRAPARTE'. Si H4 es 0 (diferencia igual a cero), el estado es 'CONCILIADO'. Si H4 es diferente de 0, el estado es 'DIFERENCIA'. Usa la función SI anidada."*

3. Copilot generará una fórmula como:

   ```excel
   =SI(F4="SIN CONTRAPARTE","SIN CONTRAPARTE",
     SI(H4=0,"CONCILIADO","DIFERENCIA"))
   ```

4. Aplica la fórmula en I4 y extiéndela hacia abajo.

5. Haz clic en la celda **J4**. Pide a Copilot:

   > *"En J4, crea una fórmula de prioridad basada en el valor absoluto de H4: si ABS(H4) es mayor a 10000, prioridad 'ALTA'; si está entre 1000 y 10000, prioridad 'MEDIA'; si es menor a 1000 y mayor a 0, prioridad 'BAJA'; si es 0 y el estado es CONCILIADO, prioridad 'N/A'."*

6. Aplica la fórmula generada en J4 y extiéndela hacia abajo.

**Resultado esperado:** Las columnas H, I y J muestran la diferencia monetaria, el estado del cruce y la prioridad para cada registro.

**Verificación:** Confirma que:
- Al menos algunos registros muestran "CONCILIADO" (diferencia = 0).
- Al menos algunos registros muestran "DIFERENCIA" (las discrepancias intencionalmente insertadas).
- Al menos algunos registros muestran "SIN CONTRAPARTE".

---

### Fase 3: Identificación de Discrepancias con Copilot

**Objetivo de la fase:** Usar el panel analítico de Copilot para formular preguntas sobre los datos del cruce, identificar patrones de discrepancia y aplicar formato condicional para visualización inmediata.

---

#### Paso 3.1 — Aplicar formato condicional para visualización de excepciones

**Objetivo:** Hacer visualmente distinguibles los tres estados del cruce mediante colores, facilitando la revisión inmediata.

**Instrucciones:**

1. Selecciona el rango de la columna **I** que contiene los estados (por ejemplo, I4:I200, ajusta según el número de filas con datos).
2. Ve a **Inicio → Formato condicional → Nueva regla**.
3. Selecciona **"Aplicar formato únicamente a las celdas que contengan"**.
4. Configura la primera regla:
   - Condición: **El texto de la celda contiene** → `CONCILIADO`
   - Formato: Relleno verde claro (RGB: 198, 239, 206), fuente verde oscuro.
   - Haz clic en **Aceptar**.
5. Repite para crear una segunda regla:
   - Condición: **El texto de la celda contiene** → `DIFERENCIA`
   - Formato: Relleno amarillo (RGB: 255, 235, 156), fuente naranja oscuro.
6. Repite para una tercera regla:
   - Condición: **El texto de la celda contiene** → `SIN CONTRAPARTE`
   - Formato: Relleno rojo claro (RGB: 255, 199, 206), fuente rojo oscuro.
7. Aplica el mismo esquema de colores a la columna **J** (Prioridad): ALTA = rojo, MEDIA = amarillo, BAJA = verde, N/A = gris.

**Resultado esperado:** La columna de estado muestra un semáforo visual inmediato: verde para conciliados, amarillo para diferencias y rojo para sin contraparte.

**Verificación:** Desplázate por la hoja y confirma que los colores son consistentes con los valores de texto en cada celda.

---

#### Paso 3.2 — Formular preguntas analíticas a Copilot sobre las discrepancias

**Objetivo:** Usar el lenguaje natural de Copilot para obtener análisis cuantitativos de las discrepancias identificadas.

**Instrucciones:**

1. Asegúrate de estar en la hoja `Cruce_LM_CxP` con el panel de Copilot abierto.
2. Escribe la siguiente pregunta en Copilot:

   > *"¿Cuántos registros tienen estado 'DIFERENCIA' y cuál es el monto total de las diferencias? ¿Cuántos tienen estado 'SIN CONTRAPARTE'?"*

3. Anota los resultados que Copilot proporciona en un bloc de notas o en una celda de la hoja de trabajo (por ejemplo, en la celda L1 escribe un resumen).

4. Escribe la siguiente pregunta:

   > *"¿Qué cuentas contables (columna B) tienen las diferencias más grandes? Muéstrame las 5 cuentas con mayor diferencia acumulada."*

5. Si Copilot genera una tabla de resultados, haz clic en **"Insertar en hoja"** o copia los resultados a una celda disponible.

6. Escribe una tercera pregunta:

   > *"¿Existen registros con diferencias mayores a $10,000? Lista las referencias y los montos."*

7. Anota o inserta los resultados.

**Resultado esperado:** Copilot proporciona respuestas cuantitativas con conteos y montos, identificando las cuentas y registros más críticos. Deberías encontrar al menos 3–5 registros con diferencias superiores a $10,000 (discrepancias intencionalmente insertadas en los datos de práctica).

**Verificación:** Valida manualmente una de las respuestas de Copilot. Por ejemplo, si Copilot dice que hay 8 registros con diferencias > $10,000, aplica un autofiltro en la columna H para valores > 10000 y confirma el conteo.

---

#### Paso 3.3 — Crear tabla resumen de discrepancias con SUMAR.SI.CONJUNTO

**Objetivo:** Construir una tabla de resumen que agregue las discrepancias por cuenta contable y estado, usando fórmulas sugeridas por Copilot.

**Instrucciones:**

1. Inserta una nueva hoja y nómbrala `Resumen_Discrepancias`.
2. Crea los siguientes encabezados en la fila 1:

   | A | B | C | D | E |
   |---|---|---|---|---|
   | Cuenta_Contable | Total_Conciliado | Total_Diferencias | Total_Sin_Contraparte | Monto_Diferencia_Neta |

3. En el panel de Copilot, escribe:

   > *"En la hoja Resumen_Discrepancias, necesito fórmulas SUMAR.SI.CONJUNTO para calcular: en la columna B, la suma de Monto_LM de la hoja Cruce_LM_CxP donde el Estado_Cruce sea 'CONCILIADO' para cada cuenta de la columna A. En la columna C, la misma suma pero donde Estado_Cruce sea 'DIFERENCIA'. En la columna D, donde Estado_Cruce sea 'SIN CONTRAPARTE'. En la columna E, la suma de la columna Diferencia donde Estado_Cruce sea 'DIFERENCIA'. Genera las fórmulas para la primera fila de datos (fila 2) y yo las extenderé."*

4. Copilot generará fórmulas similares a:

   ```excel
   ' Columna B (Total Conciliado para la cuenta en A2):
   =SUMAR.SI.CONJUNTO(
     Cruce_LM_CxP!$D:$D,
     Cruce_LM_CxP!$B:$B, $A2,
     Cruce_LM_CxP!$I:$I, "CONCILIADO"
   )
   
   ' Columna E (Monto diferencia neta para la cuenta en A2):
   =SUMAR.SI.CONJUNTO(
     Cruce_LM_CxP!$H:$H,
     Cruce_LM_CxP!$B:$B, $A2,
     Cruce_LM_CxP!$I:$I, "DIFERENCIA"
   )
   ```

5. Primero, llena la columna A con la lista única de cuentas contables. Pide a Copilot:

   > *"¿Cómo puedo obtener una lista de valores únicos de la columna Cuenta_LM de la hoja Cruce_LM_CxP? Sugiere la fórmula o método más eficiente."*

6. Aplica el método sugerido (probablemente `ÚNICO` o una tabla dinámica) para poblar la columna A con las cuentas únicas.
7. Aplica las fórmulas de las columnas B–E y extiéndelas para todas las cuentas.

**Resultado esperado:** Una tabla resumen que muestra, por cada cuenta contable, los montos conciliados, en diferencia y sin contraparte, con el monto neto de discrepancia.

**Verificación:** Suma la columna E (Monto_Diferencia_Neta). El total debe coincidir con el monto total de diferencias que Copilot reportó en el Paso 3.2.

---

### Fase 4: Generación del Reporte de Excepción

**Objetivo de la fase:** Construir el reporte de excepción formal en la plantilla proporcionada, integrando los resultados del cruce y las visualizaciones necesarias para la revisión ejecutiva.

---

#### Paso 4.1 — Abrir y estructurar la plantilla de reporte de excepción

**Objetivo:** Preparar la plantilla oficial del reporte de excepción con los datos del análisis realizado.

**Instrucciones:**

1. Abre el archivo `M2_Plantilla_Reporte_Excepcion.xlsx`.
2. Guarda una copia como `Lab02_[TuNombre]_Reporte_Excepcion.xlsx` en la carpeta `Lab02_CruceContable`.
3. Revisa la estructura de la plantilla. Debe contener al menos las siguientes secciones:
   - **Portada:** Nombre del reporte, período, responsable, fecha de emisión.
   - **Sección 1:** Resumen ejecutivo (espacio para narrativa).
   - **Sección 2:** Tabla de discrepancias por cuenta.
   - **Sección 3:** Detalle de registros sin contraparte.
   - **Sección 4:** Gráfica o visualización de excepciones.
4. Completa los campos de la portada:
   - Nombre del reporte: `Reporte de Excepción — Cruce Libro Mayor vs. Auxiliar CxP`
   - Período: `Mayo 2024`
   - Responsable: Tu nombre
   - Fecha de emisión: Fecha actual

**Resultado esperado:** La plantilla tiene la portada completa y está lista para recibir los datos del análisis.

**Verificación:** Confirma que el archivo se guardó correctamente con el nombre personalizado.

---

#### Paso 4.2 — Vincular los datos del resumen al reporte de excepción

**Objetivo:** Traer la tabla de resumen de discrepancias al reporte de excepción de forma dinámica.

**Instrucciones:**

1. En la hoja de **Sección 2** de la plantilla de reporte, haz clic en la celda donde debe comenzar la tabla de discrepancias.
2. En el panel de Copilot (asegúrate de que ambos archivos estén abiertos), escribe:

   > *"Necesito copiar la tabla de la hoja Resumen_Discrepancias del archivo Lab02_[TuNombre]_CruceContable.xlsx a esta sección del reporte. ¿Cuál es la mejor forma de hacerlo para que los datos se actualicen automáticamente si el análisis cambia?"*

3. Copilot sugerirá opciones como:
   - **Opción A:** Pegado especial con vínculo (Pegado Especial → Pegar vínculo).
   - **Opción B:** Fórmulas de referencia externa directa.
4. Aplica el método sugerido. Si usas referencias externas, la fórmula tendrá la forma:

   ```excel
   ='[Lab02_[TuNombre]_CruceContable.xlsx]Resumen_Discrepancias'!$A$1
   ```

5. Alternativamente (si los vínculos externos presentan problemas), copia y pega los valores de la tabla resumen como **Pegado especial → Solo valores** en la Sección 2 del reporte.

6. Aplica formato de tabla al rango pegado: selecciona el rango → **Ctrl + T** → asigna nombre `TablaReporteExcepcion`.

**Resultado esperado:** La Sección 2 del reporte muestra la tabla de discrepancias por cuenta contable con formato profesional.

**Verificación:** Confirma que la tabla tiene los mismos valores que la hoja `Resumen_Discrepancias` del archivo de cruce.

---

#### Paso 4.3 — Crear una visualización de excepciones con Copilot

**Objetivo:** Generar una gráfica que muestre la distribución de los estados de cruce para incluir en el reporte ejecutivo.

**Instrucciones:**

1. Regresa al archivo `Lab02_[TuNombre]_CruceContable.xlsx`, hoja `Resumen_Discrepancias`.
2. En el panel de Copilot, escribe:

   > *"Crea una gráfica de barras apiladas o un gráfico de pastel que muestre la distribución de registros por estado de cruce (CONCILIADO, DIFERENCIA, SIN CONTRAPARTE) con los datos de esta hoja. El gráfico debe tener título 'Distribución del Estado de Cruce — Mayo 2024' y usar los colores verde, amarillo y rojo para cada estado respectivamente."*

3. Copilot generará el gráfico o te guiará para crearlo. Sigue las instrucciones.
4. Si Copilot no puede crear el gráfico directamente, crea manualmente una tabla de conteo auxiliar:
   - Usa `CONTAR.SI` en la hoja `Cruce_LM_CxP` para contar registros por estado:

   ```excel
   =CONTAR.SI(Cruce_LM_CxP!$I:$I,"CONCILIADO")
   =CONTAR.SI(Cruce_LM_CxP!$I:$I,"DIFERENCIA")
   =CONTAR.SI(Cruce_LM_CxP!$I:$I,"SIN CONTRAPARTE")
   ```

   - Selecciona esta tabla de conteo → **Insertar → Gráfico de sectores (pastel)**.
   - Aplica los colores del semáforo manualmente.

5. Copia el gráfico (**Ctrl + C**) y pégalo en la Sección 4 del reporte de excepción (**Pegado especial → Imagen o Como objeto de Excel**).

**Resultado esperado:** El reporte de excepción incluye una visualización clara de la distribución de estados de cruce.

**Verificación:** El gráfico muestra los tres estados con sus colores correspondientes y los porcentajes/cantidades son consistentes con los datos del cruce.

---

### Fase 5: Narrativa Ejecutiva con Copilot

**Objetivo de la fase:** Usar Copilot para generar el resumen ejecutivo narrativo del reporte de excepción, estructurado para la dirección financiera corporativa.

---

#### Paso 5.1 — Preparar el contexto para Copilot antes de generar la narrativa

**Objetivo:** Recopilar los datos clave del análisis para proporcionar a Copilot el contexto necesario para generar una narrativa precisa y profesional.

**Instrucciones:**

1. En una celda libre de la hoja `Resumen_Discrepancias` (por ejemplo, celda G1), recopila los siguientes datos del análisis (puedes escribirlos manualmente basándote en los resultados de los pasos anteriores):

   ```
   Total de registros analizados: [número]
   Registros CONCILIADOS: [número] ([porcentaje]%)
   Registros con DIFERENCIA: [número] ([porcentaje]%)
   Registros SIN CONTRAPARTE: [número] ([porcentaje]%)
   Monto total de diferencias: $[monto]
   Cuenta con mayor discrepancia: [cuenta] — $[monto]
   Número de diferencias > $10,000: [número]
   ```

2. Revisa estos datos para asegurarte de que son correctos antes de pasarlos a Copilot.

**Resultado esperado:** Un bloque de datos consolidado listo para ser utilizado como contexto en la generación de la narrativa.

**Verificación:** Confirma que los números son consistentes con los resultados de las fases anteriores.

---

#### Paso 5.2 — Generar la narrativa ejecutiva con Copilot

**Objetivo:** Usar Copilot para redactar el resumen ejecutivo del reporte de excepción en tono profesional y orientado a la dirección financiera.

**Instrucciones:**

1. Abre el archivo `Lab02_[TuNombre]_Reporte_Excepcion.xlsx` y navega a la **Sección 1 (Resumen Ejecutivo)**.
2. Abre el panel de Copilot y escribe la siguiente instrucción (sustituye los corchetes con los datos reales que recopilaste en el paso anterior):

   > *"Redacta un resumen ejecutivo de máximo 200 palabras para un reporte de excepción contable dirigido a la Dirección Financiera. El reporte corresponde al cruce del Libro Mayor contra el Auxiliar de Cuentas por Pagar del período Mayo 2024. Los hallazgos son los siguientes:*
   >
   > *- Total de registros analizados: [número]*
   > *- Registros conciliados: [número] ([porcentaje]%)*
   > *- Registros con diferencia monetaria: [número] ([porcentaje]%) por un monto total de $[monto]*
   > *- Registros sin contraparte en el auxiliar: [número]*
   > *- La cuenta con mayor discrepancia es [cuenta] con $[monto] de diferencia*
   > *- Hay [número] diferencias individuales superiores a $10,000 que requieren atención prioritaria*
   >
   > *El tono debe ser formal, conciso y orientado a la acción. Incluye una recomendación de seguimiento."*

3. Revisa el texto generado por Copilot. Evalúa:
   - ¿Es preciso en los datos mencionados?
   - ¿El tono es apropiado para la dirección financiera?
   - ¿Incluye una recomendación clara?

4. Si necesitas ajustes, escribe en Copilot:

   > *"Ajusta el párrafo anterior para hacerlo más conciso. Enfatiza la urgencia de las diferencias mayores a $10,000 y agrega una fecha límite sugerida de resolución de 5 días hábiles."*

5. Una vez satisfecho con el texto, cópialo y pégalo en la celda de resumen ejecutivo de la plantilla del reporte.

**Resultado esperado:** La Sección 1 del reporte contiene un resumen ejecutivo profesional de 150–200 palabras con datos precisos, tono formal y recomendación de acción.

**Verificación:** Lee el resumen en voz alta. Confirma que cualquier persona de la dirección financiera, sin ver los datos detallados, puede comprender la situación y las acciones requeridas.

---

#### Paso 5.3 — Finalizar y guardar el reporte de excepción

**Objetivo:** Dar los toques finales al reporte y guardarlo en formato listo para distribución.

**Instrucciones:**

1. Revisa el reporte completo navegando por todas las secciones:
   - **Portada:** Datos completos y correctos.
   - **Sección 1:** Narrativa ejecutiva insertada.
   - **Sección 2:** Tabla de discrepancias por cuenta con formato profesional.
   - **Sección 3:** Lista de registros sin contraparte (si la plantilla la incluye, copia los registros con estado "SIN CONTRAPARTE" de la hoja `Cruce_LM_CxP`).
   - **Sección 4:** Gráfica de distribución de estados.

2. Ajusta el ancho de columnas y la presentación visual según sea necesario.

3. Guarda el reporte en formato Excel (**Ctrl + S**).

4. Opcionalmente, guarda también una copia en PDF: **Archivo → Exportar → Crear documento PDF/XPS**.

5. Guarda también el archivo de cruce `Lab02_[TuNombre]_CruceContable.xlsx` con todos los cambios.

**Resultado esperado:** Dos archivos finales guardados: el libro de trabajo de cruce completo y el reporte de excepción en formato Excel (y opcionalmente PDF).

**Verificación:** Cierra ambos archivos y vuélvelos a abrir para confirmar que todos los datos y fórmulas se guardaron correctamente.

---

## 7. Validación y Pruebas

Una vez completadas las cinco fases, ejecuta las siguientes pruebas de validación para confirmar que el flujo de trabajo funciona correctamente:

### 7.1 Prueba de integridad del cruce

| Prueba | Procedimiento | Resultado esperado |
|---|---|---|
| **Conteo de registros** | Cuenta las filas con datos en la hoja `Cruce_LM_CxP`. Compara con el número de registros de CxP en `TablaLibroMayor`. | Deben coincidir (ningún registro del libro mayor debe haberse perdido en el cruce). |
| **Suma de montos LM** | Suma la columna D (Monto_LM) en `Cruce_LM_CxP`. Compara con la suma de los mismos registros en `TablaLibroMayor`. | Las sumas deben ser idénticas. |
| **Clasificación exhaustiva** | Confirma que la columna I no tiene celdas vacías ni valores distintos a CONCILIADO, DIFERENCIA o SIN CONTRAPARTE. | 100% de celdas clasificadas. |

### 7.2 Prueba de precisión de discrepancias

1. Selecciona manualmente 3 registros con estado "DIFERENCIA" de la hoja `Cruce_LM_CxP`.
2. Para cada uno, verifica manualmente en la hoja `Auxiliar_CxP` el monto correspondiente.
3. Confirma que la diferencia calculada en la columna H es correcta (Monto_LM − Monto_CxP).

### 7.3 Prueba del reporte de excepción

1. Abre `Lab02_[TuNombre]_Reporte_Excepcion.xlsx`.
2. Confirma que los totales de la Sección 2 son consistentes con los datos de la hoja `Resumen_Discrepancias`.
3. Confirma que la narrativa de la Sección 1 menciona cifras que corresponden exactamente a los resultados del análisis.

### 7.4 Lista de verificación final

- [ ] La hoja `Cruce_LM_CxP` tiene datos en todas las columnas A–J para todos los registros.
- [ ] El formato condicional muestra correctamente los tres colores del semáforo.
- [ ] La hoja `Resumen_Discrepancias` tiene fórmulas SUMAR.SI.CONJUNTO funcionando.
- [ ] El reporte de excepción tiene las cuatro secciones completas.
- [ ] La narrativa ejecutiva fue generada con Copilot y revisada por el participante.
- [ ] Ambos archivos están guardados en la carpeta `Lab02_CruceContable`.

---

## 8. Solución de Problemas

### Problema 1: Copilot no reconoce los nombres de las tablas al formular instrucciones

**Síntoma:** Al escribir instrucciones en Copilot mencionando `TablaLibroMayor`, `TablaAuxCxP` o `TablaAuxCxC`, Copilot responde que no puede encontrar esas tablas o genera fórmulas con referencias incorrectas (por ejemplo, usa `Hoja1!A:A` en lugar del nombre de tabla).

**Causa:** El nombre de la tabla no fue asignado correctamente, o Copilot está operando sobre un libro que no tiene las tablas activas en la vista actual. Copilot en Excel tiene mejor contexto cuando la hoja activa contiene o referencia las tablas mencionadas.

**Solución:**
1. Verifica los nombres de las tablas: ve a **Fórmulas → Administrador de nombres** y confirma que los nombres `TablaLibroMayor`, `TablaAuxCxP` y `TablaAuxCxC` aparecen en la lista con el tipo "Tabla".
2. Si los nombres no existen o son diferentes, haz clic en la tabla correspondiente, ve a **Diseño de tabla** y corrige el nombre en el campo "Nombre de tabla".
3. Antes de escribir la instrucción en Copilot, navega a la hoja que contiene la tabla principal que mencionas en la instrucción. Copilot tiene mejor contexto cuando la tabla está visible.
4. Si el problema persiste, reformula la instrucción usando referencias de hoja explícitas: en lugar de "en TablaLibroMayor", escribe "en la hoja Libro_Mayor, en el rango de datos que comienza en A1".

---

### Problema 2: Las fórmulas BUSCARX / SUMAR.SI.CONJUNTO devuelven resultados incorrectos o errores #N/A en registros que deberían coincidir

**Síntoma:** Varios registros muestran "SIN CONTRAPARTE" o diferencias inesperadas, pero al revisar manualmente en el auxiliar, la referencia sí existe. O las fórmulas SUMAR.SI.CONJUNTO devuelven 0 cuando debería haber valores.

**Causa:** Incompatibilidad de formato entre los campos de búsqueda. El valor en la columna de referencia del libro mayor puede ser texto mientras que en el auxiliar es número (o viceversa), o puede haber espacios adicionales al inicio/final de las cadenas de texto que impiden la coincidencia exacta.

**Solución:**
1. Selecciona una celda de referencia en el libro mayor y una celda equivalente en el auxiliar. Verifica en la barra de fórmulas si los valores se alinean a la izquierda (texto) o a la derecha (número). Deben ser del mismo tipo.
2. Para limpiar espacios, pide a Copilot:
   > *"¿Cómo puedo eliminar espacios adicionales al inicio y al final de los valores en la columna Referencia de TablaLibroMayor y TablaAuxCxP? Sugiere una fórmula o proceso."*
   Copilot sugerirá usar `ESPACIOS()` (TRIM en inglés) en columnas auxiliares.
3. Si el problema es de tipo de dato (texto vs. número), usa `VALOR()` para convertir texto a número o `TEXTO()` para convertir número a texto, según corresponda, en una columna normalizada antes de ejecutar el cruce.
4. Después de normalizar, actualiza las fórmulas de cruce para que referencien las columnas normalizadas en lugar de las originales.

---

## 9. Limpieza del Entorno

Al finalizar el laboratorio, realiza los siguientes pasos de limpieza para mantener ordenado tu entorno de trabajo:

1. **Guardar todos los archivos:** Confirma que `Lab02_[TuNombre]_CruceContable.xlsx` y `Lab02_[TuNombre]_Reporte_Excepcion.xlsx` están guardados con los últimos cambios.

2. **Cerrar archivos de origen:** Cierra los archivos `M2_Libro_Mayor.xlsx`, `M2_Auxiliar_CxP.xlsx` y `M2_Auxiliar_CxC.xlsx` sin guardar cambios (estos son los originales proporcionados por el instructor y no deben modificarse).

3. **Cerrar el panel de Copilot:** Haz clic en la **X** del panel lateral de Copilot para cerrarlo y liberar recursos.

4. **Verificar la carpeta de trabajo:** Abre la carpeta `Lab02_CruceContable` en el Explorador de archivos y confirma que contiene:
   - `M2_Libro_Mayor.xlsx` (original, sin modificar)
   - `M2_Auxiliar_CxP.xlsx` (original, sin modificar)
   - `M2_Auxiliar_CxC.xlsx` (original, sin modificar)
   - `M2_Plantilla_Reporte_Excepcion.xlsx` (original, sin modificar)
   - `Lab02_[TuNombre]_CruceContable.xlsx` (tu archivo de trabajo)
   - `Lab02_[TuNombre]_Reporte_Excepcion.xlsx` (tu reporte final)

5. **No eliminar los archivos de trabajo:** Estos serán utilizados como referencia en laboratorios posteriores del Módulo 2.

> ⚠️ **Recordatorio de confidencialidad:** Si durante el laboratorio introdujiste accidentalmente algún dato real de tu organización, notifica al instructor y elimina esa información de los archivos antes de cerrar la sesión.

---

## 10. Resumen

### Puntos Clave del Laboratorio

En este laboratorio aplicaste de manera integral el flujo de trabajo inteligente de cruce de cuentas descrito en la Lección 2.1. Los aprendizajes consolidados son:

| Competencia | Herramienta/Técnica aplicada |
|---|---|
| Normalización de datos para cruce | Tablas de Excel con nombres descriptivos, fórmulas de limpieza asistidas por Copilot |
| Cruce automatizado de cuentas | BUSCARX / XLOOKUP con SI.ERROR, instrucciones en lenguaje natural a Copilot |
| Cálculo y clasificación de diferencias | Fórmulas SI anidadas, columna de prioridad por rango de monto |
| Análisis cuantitativo de discrepancias | Preguntas analíticas al panel de Copilot, SUMAR.SI.CONJUNTO, CONTAR.SI |
| Visualización de excepciones | Formato condicional semáforo, gráfica de distribución de estados |
| Reporte ejecutivo estructurado | Plantilla de reporte, narrativa generada con Copilot y revisada por el profesional |

### Reflexión sobre el Valor Profesional

El flujo que construiste en este laboratorio replica lo que en muchas organizaciones toma dos o más días hábiles de trabajo manual. La combinación de estructuras de datos bien definidas (Tablas de Excel), fórmulas avanzadas generadas por Copilot y análisis en lenguaje natural comprime ese tiempo a menos de dos horas, con mayor trazabilidad y reproducibilidad. Más importante aún: el resultado no es solo más rápido, sino más auditable — cada fórmula y criterio de cruce queda documentado y puede ser revisado por auditores internos o externos.

El juicio profesional del contador sigue siendo indispensable: Copilot sugiere y ejecuta, pero eres tú quien define los criterios de negocio, valida los resultados y toma decisiones sobre las discrepancias encontradas.

### Próximos Pasos

- **Laboratorio 02-00-02:** Extenderás este flujo para incluir el cruce con el Auxiliar de CxC y aplicarás técnicas de detección de anomalías más avanzadas (patrones inusuales, duplicados, registros fuera de período).
- **Módulo 3:** Aprenderás a integrar directamente exportaciones de Oracle en este tipo de flujos, eliminando pasos manuales de preparación de datos.

### Recursos de Referencia

| Recurso | Enlace |
|---|---|
| Documentación oficial de Copilot en Excel | [support.microsoft.com — Copilot en Excel](https://support.microsoft.com/es-es/office/copilot-en-excel-d7110502-0334-4b4f-a175-a73abdfc118a) |
| Función BUSCARX (XLOOKUP) — Microsoft Support | [support.microsoft.com — BUSCARX](https://support.microsoft.com/es-es/office/funci%C3%B3n-buscarx-b7fd680e-6d10-43e6-84f9-88eae8bf5929) |
| Función SUMAR.SI.CONJUNTO — Microsoft Support | [support.microsoft.com — SUMAR.SI.CONJUNTO](https://support.microsoft.com/es-es/office/funci%C3%B3n-sumar-si-conjunto-86472c60-a4f1-4f87-8f87-6d73af58c2e5) |
| Tablas de Excel para análisis de datos — Microsoft Learn | [learn.microsoft.com — Tablas Excel](https://learn.microsoft.com/es-es/office/troubleshoot/excel/use-excel-table-feature) |
| Mejores prácticas de conciliación bancaria — IMCP | [imcp.org.mx — Normas de Información Financiera](https://imcp.org.mx/normas-de-informacion-financiera/) |

---

*Fin del Laboratorio 02-00-01*

---
LAB_END---
