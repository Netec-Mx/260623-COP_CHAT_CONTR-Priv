---LAB_START---
LAB_ID: 04-00-01
---MARKDOWN---
# Uso de Copilot para la revisión de evidencias, detección de anomalías en registros contables y generación de resúmenes ejecutivos para la dirección financiera.

---

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración estimada** | 60 minutos |
| **Complejidad** | Media |
| **Nivel Bloom** | Aplicar (*Apply*) |
| **Módulo** | Módulo 4 — Revisión de Evidencias y Generación de Reportes Ejecutivos |
| **Laboratorio** | Lab 04-00-01 (Práctica 4 — Laboratorio de culminación del curso) |
| **Instructor responsable** | Asignado por la organización capacitadora |

---

## 2. Descripción General

Este laboratorio representa la práctica de culminación del curso. Los participantes integrarán todas las competencias adquiridas en los módulos anteriores para ejecutar un flujo completo de contraloría corporativa asistido por inteligencia artificial. Trabajando con un conjunto de registros contables que contiene anomalías intencionalmente insertadas, utilizarán Copilot en Excel para detectar transacciones sospechosas mediante análisis estadístico y reglas de negocio, y posteriormente emplearán Copilot en Word para transformar esos hallazgos técnicos en un resumen ejecutivo profesional dirigido a la dirección financiera. Al concluir, los participantes contarán con un entregable ejecutivo completo que podrán usar como referencia para sus flujos reales de auditoría y contraloría.

---

## 3. Objetivos de Aprendizaje

Al finalizar este laboratorio, el participante será capaz de:

- [ ] Estructurar y preparar datos de registros contables en Excel para que Copilot pueda analizarlos de manera efectiva, aplicando los criterios de calidad de datos aprendidos en el curso.
- [ ] Formular *prompts* específicos y bien estructurados en Copilot para detectar anomalías contables: transacciones fuera de horario, montos sobre límites de autorización, posibles duplicados, fraccionamiento de pagos y usuarios no autorizados.
- [ ] Aplicar formato condicional y crear visualizaciones (tablas dinámicas y gráficos) asistidas por Copilot para consolidar y priorizar los hallazgos de auditoría.
- [ ] Generar un resumen ejecutivo estructurado y profesional en Word utilizando Copilot, comunicando hallazgos, nivel de riesgo y recomendaciones a la dirección financiera de forma clara y sin tecnicismos innecesarios.
- [ ] Revisar y validar el entregable final como si fuera un reporte real para la dirección financiera, evaluando claridad, precisión y utilidad ejecutiva del documento generado con IA.

---

## 4. Prerrequisitos

### 4.1 Conocimientos Previos

| Área | Nivel Requerido |
|---|---|
| Microsoft Excel 365 con Copilot | Haber completado las Prácticas 1, 2 y 3, o experiencia equivalente |
| Conceptos contables básicos | Familiaridad con pólizas, comprobantes, libro mayor, cuentas por pagar |
| Auditoría interna | Conocimiento básico de evidencias, niveles de autorización y segregación de funciones |
| Copilot en Word | Haber usado Copilot en Word al menos en la Práctica 3, o equivalente |

### 4.2 Acceso y Licencias

| Requisito | Verificación |
|---|---|
| Licencia Microsoft 365 Copilot activa | Abrir Excel → Cinta de opciones → confirmar botón **Copilot** visible |
| Microsoft Word 365 con Copilot | Abrir Word → confirmar botón **Copilot** en la cinta *Inicio* |
| Archivos de práctica del Módulo 4 | Recibidos del instructor antes del inicio del laboratorio |
| Conexión a Internet estable (≥ 10 Mbps) | Verificar con el instructor de sala |

### 4.3 Archivos de Práctica Requeridos

Solicitar al instructor los siguientes archivos antes de iniciar:

```
M4_Registros_Contables_Auditoria.xlsx   ← Datos principales con anomalías insertadas
M4_Catalogo_Usuarios_Autorizados.xlsx   ← Lista de usuarios con permisos de registro
M4_Limites_Autorizacion.xlsx            ← Tabla de límites de aprobación por nivel
M4_Plantilla_Resumen_Ejecutivo.docx     ← Plantilla base para el informe final
```

> ⚠️ **Aviso de confidencialidad:** Todos los archivos contienen datos financieros **simulados**. No sustituya estos archivos por datos reales de su organización sin autorización del área de seguridad de la información.

---

## 5. Entorno del Laboratorio

### 5.1 Hardware Mínimo Recomendado

| Componente | Especificación Mínima |
|---|---|
| Procesador | Intel Core i5 8ª gen. / AMD Ryzen 5 o superior |
| RAM | 8 GB (se recomienda 16 GB para archivos grandes) |
| Almacenamiento libre | 500 MB disponibles para archivos de trabajo del laboratorio |
| Pantalla | Resolución 1920×1080 o superior |
| Conexión | Internet estable ≥ 10 Mbps |

### 5.2 Software Requerido

| Software | Versión | Estado Requerido |
|---|---|---|
| Microsoft Excel 365 | Versión 2308 o superior (canal mensual) | Abierto con cuenta corporativa |
| Microsoft Word 365 | Versión 2308 o superior | Disponible con la misma cuenta |
| Microsoft 365 Copilot | Versión empresarial 2024 integrada | Licencia activa y verificada |
| Navegador web | Edge o Chrome (versión más reciente) | Para acceso de respaldo a Copilot web |

### 5.3 Configuración Inicial del Entorno

Antes de iniciar los pasos del laboratorio, complete las siguientes verificaciones:

**1. Verificar que Copilot esté activo en Excel:**
```
Abrir Excel → Crear libro en blanco → Cinta de opciones → 
confirmar que el botón "Copilot" aparece en la sección "Inicio"
```

**2. Verificar canal de actualización de Microsoft 365:**
```
Excel → Archivo → Cuenta → Información de actualización de Office →
confirmar "Canal actual (mensual)" — NO "Canal semestral"
```

**3. Organizar archivos de práctica:**
```
Crear carpeta local: C:\Capacitacion_M365Copilot\Modulo4\
Copiar los 4 archivos de práctica a esa carpeta
```

**4. Verificar Copilot en Word:**
```
Abrir Word → Documento en blanco → Cinta "Inicio" →
confirmar botón "Copilot" visible en la sección correspondiente
```

---

## 6. Procedimiento Paso a Paso

> 📋 **Estructura del laboratorio:** El laboratorio se divide en **cuatro fases** que simulan el flujo real de un proceso de auditoría asistida por IA:
> - **Fase 1** (Pasos 1–2): Preparación y exploración inicial de datos
> - **Fase 2** (Pasos 3–5): Detección de anomalías con Copilot en Excel
> - **Fase 3** (Paso 6): Consolidación visual de hallazgos
> - **Fase 4** (Pasos 7–8): Generación del resumen ejecutivo en Word

---

### Paso 1 — Apertura y Estructuración de los Datos Contables

**Objetivo:** Cargar el archivo principal de registros contables en Excel y verificar que los datos están correctamente estructurados como tabla, condición indispensable para que Copilot pueda analizarlos de manera efectiva.

#### Instrucciones

1. Abra **Microsoft Excel 365** con su cuenta corporativa.

2. Navegue a `C:\Capacitacion_M365Copilot\Modulo4\` y abra el archivo:
   ```
   M4_Registros_Contables_Auditoria.xlsx
   ```

3. Observe la hoja **"Registros_Q4"**. Familiarícese con la estructura de columnas. El archivo debe contener las siguientes columnas principales:

   | Columna | Descripción |
   |---|---|
   | `ID_Transaccion` | Identificador único de cada registro |
   | `Fecha_Registro` | Fecha y hora del registro contable |
   | `Cuenta_Contable` | Número de cuenta del catálogo |
   | `Descripcion_Cuenta` | Nombre descriptivo de la cuenta |
   | `Proveedor_Beneficiario` | Nombre del proveedor o beneficiario |
   | `RFC_Proveedor` | RFC del proveedor |
   | `Monto` | Importe de la transacción (MXN) |
   | `Usuario_Registro` | Usuario que capturó el registro |
   | `Tipo_Poliza` | Tipo de póliza (Egreso, Ingreso, Diario) |
   | `Numero_Factura` | Folio de la factura soporte |
   | `Centro_Costo` | Centro de costo al que se carga |
   | `Nivel_Autorizacion` | Nivel jerárquico requerido para aprobar |

4. Verifique que los datos **no tengan filas en blanco** dentro del rango. Si encuentra alguna, elimínela.

5. Haga clic en cualquier celda dentro del rango de datos y presione `Ctrl + T` para convertir el rango en tabla formal de Excel.

6. En el cuadro de diálogo **"Crear tabla"**, confirme que la opción **"La tabla tiene encabezados"** esté marcada. Haga clic en **Aceptar**.

7. En la cinta **"Diseño de tabla"** (aparece cuando la tabla está seleccionada), en el campo **"Nombre de tabla"**, escriba:
   ```
   TablaAuditoria_Q4
   ```
   Presione `Enter` para confirmar.

8. Guarde el archivo con `Ctrl + S`.

#### Resultado Esperado

La hoja de Excel muestra los registros contables formateados como tabla con bandas de color alternas, el nombre `TablaAuditoria_Q4` visible en la cinta de diseño, y el botón **Copilot** activo en la cinta de opciones de Excel (no atenuado).

#### Verificación

- [ ] El nombre `TablaAuditoria_Q4` aparece en el cuadro de nombre de tabla (cinta Diseño de tabla).
- [ ] No hay celdas combinadas en los encabezados.
- [ ] El botón **Copilot** en la cinta de opciones está disponible (no en gris).
- [ ] La columna `Monto` tiene formato de número (no texto).

---

### Paso 2 — Análisis Exploratorio Inicial con Copilot

**Objetivo:** Utilizar Copilot para obtener una visión estadística descriptiva de los datos, establecer líneas base (promedios, rangos, distribución) que servirán como referencia para identificar anomalías en los pasos siguientes.

#### Instrucciones

1. Con el cursor dentro de la tabla `TablaAuditoria_Q4`, haga clic en el botón **Copilot** en la cinta de opciones. Se abrirá el panel de Copilot en el lado derecho de la pantalla.

2. En el campo de texto del panel de Copilot, ingrese el siguiente *prompt* y presione `Enter`:

   ```text
   Proporciona un análisis estadístico descriptivo de la tabla TablaAuditoria_Q4.
   Incluye: número total de registros, suma total de montos, monto promedio,
   monto máximo, monto mínimo, y el número de proveedores únicos.
   Muestra los resultados en una tabla resumen clara.
   ```

3. Revise el resumen estadístico que genera Copilot. **Anote** los valores clave en el cuaderno de trabajo o en una hoja auxiliar del mismo libro (puede insertar una hoja nueva llamada **"Análisis_Base"**):
   - Monto promedio de transacciones: `$__________`
   - Monto máximo registrado: `$__________`
   - Número total de registros: `__________`
   - Número de proveedores únicos: `__________`

4. Ahora ingrese un segundo *prompt* para explorar la distribución temporal de los registros:

   ```text
   ¿En qué días de la semana se concentra el mayor número de transacciones
   en la tabla TablaAuditoria_Q4? Muestra un conteo por día de la semana
   ordenado de mayor a menor.
   ```

5. Observe si Copilot identifica transacciones en **sábado o domingo**. Esto es relevante porque los registros en días no laborables son una señal de alerta en auditoría.

6. Ingrese un tercer *prompt* para explorar la distribución por tipo de póliza:

   ```text
   Muestra cuántas transacciones hay por cada valor de la columna Tipo_Poliza
   y cuál es el monto total por cada tipo. Ordena de mayor monto total a menor.
   ```

7. Haga clic en **"Insertar en hoja"** o **"Agregar a hoja"** (según la versión de Copilot) para que el resumen quede registrado en la hoja **"Análisis_Base"**.

#### Resultado Esperado

El panel de Copilot muestra tablas de resumen estadístico con los valores descriptivos solicitados. La hoja "Análisis_Base" contiene al menos una tabla con estadísticas base. Se identifican preliminarmente si existen registros en días no laborables.

#### Verificación

- [ ] Los valores estadísticos base están documentados (monto promedio, máximo, mínimo).
- [ ] Se obtuvo el conteo de transacciones por día de la semana.
- [ ] Se identificó si hay o no transacciones en sábado/domingo.
- [ ] La hoja "Análisis_Base" existe en el libro con al menos una tabla de resumen.

---

### Paso 3 — Detección de Anomalías: Transacciones Fuera de Horario y Usuarios No Autorizados

**Objetivo:** Aplicar *prompts* específicos en Copilot para identificar dos tipos críticos de anomalías: registros contables fuera del horario laboral y transacciones capturadas por usuarios que no están en el catálogo de usuarios autorizados.

#### Instrucciones

**Parte A — Transacciones fuera de horario laboral**

1. Regrese a la hoja **"Registros_Q4"** con la tabla `TablaAuditoria_Q4`.

2. En el panel de Copilot, ingrese el siguiente *prompt*:

   ```text
   En la tabla TablaAuditoria_Q4, agrega una columna llamada "Alerta_Horario"
   que contenga el valor "FUERA_HORARIO" si la transacción fue registrada
   un sábado, un domingo, o en un horario fuera de las 07:00 a las 20:00 horas.
   Para el resto de los registros, coloca el valor "NORMAL".
   ```

3. Revise la fórmula que Copilot propone. Copilot generará una fórmula usando funciones como `WEEKDAY()`, `HOUR()` o `TEXT()`. Haga clic en **"Insertar columna"** para aplicarla.

4. Una vez insertada la columna, ingrese el siguiente *prompt* para visualizar el resultado:

   ```text
   ¿Cuántos registros tienen el valor "FUERA_HORARIO" en la columna Alerta_Horario?
   Muéstrame también el monto total de esas transacciones y los 5 proveedores
   con mayor monto en transacciones fuera de horario.
   ```

5. **Anote** el número de transacciones fuera de horario detectadas: `__________`

**Parte B — Usuarios no autorizados**

6. **Sin cerrar** el archivo principal, abra en una ventana separada el archivo:
   ```
   M4_Catalogo_Usuarios_Autorizados.xlsx
   ```
   Familiarícese con su estructura: debe contener al menos las columnas `Usuario` y `Estado` (Activo/Inactivo/Bloqueado).

7. Regrese al archivo principal `M4_Registros_Contables_Auditoria.xlsx` y en el panel de Copilot ingrese:

   ```text
   Agrega una columna llamada "Alerta_Usuario" en la tabla TablaAuditoria_Q4.
   La lista de usuarios autorizados activos es la siguiente:
   [INSTRUCCIÓN: En este punto, el participante debe pegar manualmente la lista
   de usuarios activos del catálogo, separados por comas, por ejemplo:
   JGARCIA, MLOPEZ, RHERRERA, AMORALES, CPEREZ, LTORRES]
   
   Si el valor de la columna Usuario_Registro NO está en esa lista,
   coloca "NO_AUTORIZADO". Si sí está en la lista, coloca "AUTORIZADO".
   ```

   > 💡 **Nota:** Copilot puede generar una fórmula `IFERROR(MATCH(...))` o `COUNTIF()` para comparar la columna de usuarios contra la lista proporcionada. Si la lista es extensa, puede ser más eficiente usar `VLOOKUP` hacia la tabla del catálogo si ambos archivos están abiertos.

8. Haga clic en **"Insertar columna"** para aplicar la fórmula propuesta por Copilot.

9. Ingrese el siguiente *prompt* de seguimiento:

   ```text
   ¿Cuántos registros tienen el valor "NO_AUTORIZADO" en la columna Alerta_Usuario?
   Lista los usuarios no autorizados únicos encontrados y el monto total
   de las transacciones que registraron.
   ```

10. **Anote** los hallazgos: número de registros con usuario no autorizado y lista de usuarios detectados.

#### Resultado Esperado

La tabla `TablaAuditoria_Q4` ahora tiene dos columnas nuevas: `Alerta_Horario` y `Alerta_Usuario`. El panel de Copilot muestra resúmenes cuantitativos de ambos tipos de anomalías. Se han identificado transacciones específicas que requieren revisión.

#### Verificación

- [ ] La columna `Alerta_Horario` existe y contiene valores "FUERA_HORARIO" o "NORMAL".
- [ ] La columna `Alerta_Usuario` existe y contiene valores "NO_AUTORIZADO" o "AUTORIZADO".
- [ ] Se documentó el conteo de registros anómalos en cada categoría.
- [ ] Se identificaron los montos totales involucrados en cada tipo de anomalía.

---

### Paso 4 — Detección de Anomalías: Montos Sobre Límites de Autorización y Fraccionamiento de Pagos

**Objetivo:** Identificar transacciones que superan los límites de autorización establecidos por política corporativa, y detectar patrones de fraccionamiento de pagos (múltiples transacciones pequeñas al mismo proveedor para evadir controles).

#### Instrucciones

**Parte A — Montos sobre límites de autorización**

1. Abra en ventana separada el archivo:
   ```
   M4_Limites_Autorizacion.xlsx
   ```
   Revise su estructura. Debe contener columnas como `Nivel_Autorizacion`, `Monto_Maximo_Permitido` y posiblemente `Tipo_Poliza`.

2. En el panel de Copilot del archivo principal, ingrese:

   ```text
   Agrega una columna llamada "Alerta_Limite" en la tabla TablaAuditoria_Q4.
   Usa las siguientes reglas de límites de autorización por nivel:
   - Nivel 1: máximo $10,000
   - Nivel 2: máximo $50,000
   - Nivel 3: máximo $150,000
   - Nivel 4: sin límite
   
   Si el monto de la transacción supera el límite permitido para su nivel
   de autorización, coloca "EXCEDE_LIMITE". Si está dentro del límite,
   coloca "DENTRO_LIMITE".
   
   [Ajuste los valores según los que aparezcan en el archivo M4_Limites_Autorizacion.xlsx
   proporcionado por el instructor]
   ```

3. Aplique la columna propuesta por Copilot haciendo clic en **"Insertar columna"**.

4. Consulte a Copilot:

   ```text
   De los registros con "EXCEDE_LIMITE" en la columna Alerta_Limite,
   muéstrame los 10 con mayor monto, incluyendo proveedor, cuenta contable,
   usuario de registro y fecha. Ordena de mayor a menor monto.
   ```

**Parte B — Detección de fraccionamiento de pagos**

5. El fraccionamiento de pagos (*splitting*) es una práctica donde se dividen pagos grandes en varios pagos menores para evadir los controles de autorización. Ingrese el siguiente *prompt* en Copilot:

   ```text
   Analiza la tabla TablaAuditoria_Q4 para detectar posibles fraccionamientos
   de pago. Un fraccionamiento se define como: el mismo proveedor (RFC_Proveedor)
   con 3 o más transacciones en el mismo día calendario, donde ninguna supera
   individualmente $10,000 pero la suma del día sí supera $10,000.
   
   Agrega una columna llamada "Alerta_Fraccionamiento" con el valor
   "POSIBLE_FRACCIONAMIENTO" si cumple esa condición, o "NORMAL" si no.
   ```

6. Copilot generará una fórmula compleja usando `COUNTIFS` y `SUMIFS`. Revise la lógica propuesta antes de insertar. Haga clic en **"Insertar columna"**.

7. Consulte el resumen:

   ```text
   ¿Cuántos proveedores únicos aparecen con "POSIBLE_FRACCIONAMIENTO"?
   Muéstrame la lista con el RFC del proveedor, número de transacciones
   fraccionadas detectadas y el monto total acumulado por proveedor.
   ```

8. **Anote** los proveedores con posible fraccionamiento detectado.

#### Resultado Esperado

La tabla cuenta ahora con las columnas `Alerta_Limite` y `Alerta_Fraccionamiento`. Copilot ha identificado los registros que exceden límites y los proveedores con patrones de fraccionamiento, proporcionando listas ordenadas por riesgo.

#### Verificación

- [ ] La columna `Alerta_Limite` existe con valores "EXCEDE_LIMITE" o "DENTRO_LIMITE".
- [ ] La columna `Alerta_Fraccionamiento` existe con valores "POSIBLE_FRACCIONAMIENTO" o "NORMAL".
- [ ] Se obtuvo la lista de los 10 registros de mayor monto que exceden límites.
- [ ] Se identificaron proveedores con patrones de fraccionamiento.

---

### Paso 5 — Detección de Duplicados y Aplicación de Formato Condicional

**Objetivo:** Identificar posibles registros duplicados en la base de datos contable y aplicar formato condicional visual para facilitar la revisión por parte del equipo auditor.

#### Instrucciones

**Parte A — Detección de posibles duplicados**

1. En el panel de Copilot, ingrese:

   ```text
   En la tabla TablaAuditoria_Q4, identifica posibles registros duplicados.
   Un duplicado se define como dos o más registros con el mismo valor en
   las columnas: RFC_Proveedor, Numero_Factura y Monto (los tres campos
   deben coincidir simultáneamente).
   
   Agrega una columna llamada "Posible_Duplicado" con el valor "SÍ"
   si el registro aparece más de una vez con esa combinación, o "NO" si es único.
   ```

2. Revise la fórmula `COUNTIFS` generada por Copilot y aplíquela haciendo clic en **"Insertar columna"**.

3. Consulte el detalle:

   ```text
   Lista todos los registros donde Posible_Duplicado es "SÍ".
   Incluye ID_Transaccion, Fecha_Registro, Proveedor_Beneficiario,
   Numero_Factura y Monto. Ordena por Numero_Factura para ver
   los duplicados juntos.
   ```

**Parte B — Formato condicional visual para todas las alertas**

4. Ingrese el siguiente *prompt* para aplicar formato condicional que resalte visualmente todas las anomalías detectadas:

   ```text
   Aplica formato condicional a la tabla TablaAuditoria_Q4 con las
   siguientes reglas de color:
   - Filas donde Posible_Duplicado = "SÍ": fondo rojo claro
   - Filas donde Alerta_Limite = "EXCEDE_LIMITE": fondo naranja claro
   - Filas donde Alerta_Usuario = "NO_AUTORIZADO": fondo amarillo
   - Filas donde Alerta_Horario = "FUERA_HORARIO": fondo azul claro
   - Filas donde Alerta_Fraccionamiento = "POSIBLE_FRACCIONAMIENTO": fondo morado claro
   ```

   > 💡 **Nota:** Copilot puede generar las instrucciones de formato condicional o aplicarlas directamente. Si propone los pasos manuales, sígalos en el orden indicado. El formato condicional se aplica desde: **Inicio → Formato condicional → Nueva regla → Usar una fórmula**.

5. Una vez aplicado el formato, desplácese por la tabla para verificar visualmente que las filas con anomalías están resaltadas con los colores correspondientes.

6. Guarde el archivo con `Ctrl + S`.

#### Resultado Esperado

La tabla `TablaAuditoria_Q4` tiene cinco columnas de alerta (`Alerta_Horario`, `Alerta_Usuario`, `Alerta_Limite`, `Alerta_Fraccionamiento`, `Posible_Duplicado`) y formato condicional de colores aplicado. Las anomalías son visualmente distinguibles a primera vista.

#### Verificación

- [ ] La columna `Posible_Duplicado` existe con valores "SÍ" o "NO".
- [ ] El formato condicional está aplicado con al menos 3 colores diferenciados.
- [ ] Las filas con anomalías aparecen visualmente resaltadas al desplazarse por la tabla.
- [ ] El archivo está guardado con los cambios.

---

### Paso 6 — Consolidación de Hallazgos: Tabla Dinámica y Dashboard de Auditoría

**Objetivo:** Crear una tabla dinámica y un gráfico resumen que consoliden todos los hallazgos de anomalías, generando un panel de control (*dashboard*) de auditoría que sirva como base para el informe ejecutivo.

#### Instrucciones

1. En el panel de Copilot, ingrese el siguiente *prompt* para crear la tabla dinámica de hallazgos:

   ```text
   Crea una tabla dinámica a partir de TablaAuditoria_Q4 en una hoja nueva
   llamada "Dashboard_Auditoria". La tabla dinámica debe mostrar:
   - Filas: los cinco tipos de alerta (Alerta_Horario, Alerta_Usuario,
     Alerta_Limite, Alerta_Fraccionamiento, Posible_Duplicado)
   - Valores: conteo de registros con anomalía y suma de montos involucrados
   
   Ordena por suma de montos de mayor a menor.
   ```

2. Si Copilot genera los pasos para crear la tabla dinámica manualmente, sígalos:
   - Haga clic en cualquier celda de `TablaAuditoria_Q4`
   - Vaya a **Insertar → Tabla dinámica**
   - Seleccione **Nueva hoja de cálculo** y nombre la hoja `Dashboard_Auditoria`
   - Configure los campos según las instrucciones de Copilot

3. Una vez creada la tabla dinámica, regrese al panel de Copilot y solicite:

   ```text
   En la hoja Dashboard_Auditoria, crea un gráfico de barras agrupadas que muestre
   el número de transacciones anómalas por tipo de alerta. Usa colores distintos
   para cada barra. Agrega título: "Resumen de Anomalías Detectadas - Q4 2024"
   y etiquetas de datos visibles en cada barra.
   ```

4. Revise y ajuste el gráfico si es necesario (títulos de ejes, leyenda, tamaño).

5. En la misma hoja `Dashboard_Auditoria`, solicite a Copilot un resumen ejecutivo breve de los datos:

   ```text
   Genera un texto breve de 3 a 5 oraciones que resuma los hallazgos
   más importantes del análisis de anomalías de la tabla TablaAuditoria_Q4.
   Menciona los tipos de anomalías encontradas, el número total de registros
   afectados y el monto total en riesgo. Usa un tono profesional y objetivo.
   ```

6. Copie el texto generado por Copilot y péguelo en una celda de texto (cuadro de texto insertado) en la hoja `Dashboard_Auditoria`, debajo del gráfico.

7. Guarde el archivo completo con `Ctrl + S`.

#### Resultado Esperado

El libro de Excel contiene una hoja `Dashboard_Auditoria` con una tabla dinámica de hallazgos, un gráfico de barras con el resumen visual de anomalías por tipo, y un párrafo de resumen ejecutivo preliminar. Este material servirá como insumo directo para el informe en Word.

#### Verificación

- [ ] La hoja `Dashboard_Auditoria` existe en el libro.
- [ ] La tabla dinámica muestra conteos y montos por tipo de alerta.
- [ ] El gráfico de barras tiene título, etiquetas de datos y colores diferenciados.
- [ ] El texto de resumen preliminar está documentado en la hoja.
- [ ] El archivo está guardado.

---

### Paso 7 — Generación del Resumen Ejecutivo en Word con Copilot

**Objetivo:** Utilizar Copilot en Word para transformar los hallazgos técnicos del análisis de Excel en un resumen ejecutivo profesional, estructurado y orientado a la dirección financiera corporativa.

#### Instrucciones

1. Abra **Microsoft Word 365** y abra el archivo de plantilla:
   ```
   M4_Plantilla_Resumen_Ejecutivo.docx
   ```

2. Revise la estructura de la plantilla. Debe contener secciones predefinidas como:
   - Resumen Ejecutivo
   - Alcance y Metodología
   - Hallazgos Principales
   - Nivel de Riesgo
   - Recomendaciones
   - Anexos

3. Haga clic en el botón **Copilot** en la cinta de opciones de Word (sección *Inicio*). Se abrirá el panel de Copilot en Word.

4. Ingrese el siguiente *prompt* para generar el contenido de la sección **Resumen Ejecutivo**:

   ```text
   Redacta el contenido para la sección "Resumen Ejecutivo" de un informe
   de auditoría interna dirigido a la Dirección Financiera Corporativa.
   
   El análisis revisó [NÚMERO TOTAL DE REGISTROS] transacciones contables
   del cuarto trimestre de 2024. Se detectaron los siguientes hallazgos:
   - [X] transacciones registradas fuera de horario laboral por un monto total de $[MONTO]
   - [X] registros capturados por usuarios no autorizados por $[MONTO]
   - [X] transacciones que exceden límites de autorización por $[MONTO]
   - [X] casos de posible fraccionamiento de pagos en [X] proveedores
   - [X] posibles duplicados por $[MONTO]
   
   [Sustituya los valores entre corchetes con los datos reales anotados
   durante los pasos anteriores del laboratorio]
   
   El resumen debe ser de máximo 150 palabras, en tono ejecutivo, sin tecnicismos,
   orientado a la toma de decisiones. Debe iniciar con una frase de impacto
   sobre el nivel de riesgo identificado.
   ```

5. Revise el texto generado por Copilot. Si requiere ajustes, use el botón **"Regenerar"** o ingrese un *prompt* de refinamiento como:

   ```text
   Ajusta el tono para que sea más directo y urgente. Agrega una frase
   que mencione la necesidad de acción inmediata en los hallazgos de mayor riesgo.
   ```

6. Haga clic en **"Insertar"** para colocar el texto en la sección correspondiente de la plantilla.

7. Ahora genere el contenido para la sección **Hallazgos Principales**:

   ```text
   Redacta la sección "Hallazgos Principales" del mismo informe de auditoría.
   Organiza los hallazgos en una tabla con las columnas:
   Hallazgo | Número de Registros | Monto Total Involucrado | Nivel de Riesgo
   
   Clasifica el nivel de riesgo como: ALTO, MEDIO o BAJO, usando estos criterios:
   - ALTO: involucra usuarios no autorizados o supera límites de autorización
   - MEDIO: fraccionamiento de pagos o registros fuera de horario
   - BAJO: posibles duplicados sin indicios de intención
   
   Usa los datos del análisis realizado en Excel.
   [Ingrese los datos específicos de sus hallazgos]
   ```

8. Inserte la tabla generada en la sección correspondiente.

9. Genere el contenido para la sección **Recomendaciones**:

   ```text
   Redacta 5 recomendaciones concretas y accionables para la Dirección Financiera,
   derivadas de los hallazgos del análisis. Cada recomendación debe:
   - Estar numerada
   - Tener un título breve en negrita
   - Incluir 2-3 oraciones de descripción
   - Mencionar el área responsable de implementarla (Contraloría, TI, RRHH, etc.)
   - Indicar un plazo sugerido de implementación (inmediato, 30 días, 90 días)
   
   Tono: ejecutivo, directo, orientado a resultados y mejora del control interno.
   ```

10. Inserte las recomendaciones en la sección correspondiente de la plantilla.

#### Resultado Esperado

El documento Word tiene las secciones "Resumen Ejecutivo", "Hallazgos Principales" y "Recomendaciones" completadas con contenido profesional generado por Copilot, basado en los datos reales del análisis realizado en Excel. El documento tiene coherencia narrativa y tono ejecutivo apropiado.

#### Verificación

- [ ] La sección "Resumen Ejecutivo" está redactada con máximo 150 palabras en tono ejecutivo.
- [ ] La tabla de "Hallazgos Principales" incluye nivel de riesgo para cada hallazgo.
- [ ] Las 5 recomendaciones están numeradas con área responsable y plazo.
- [ ] El documento Word está guardado con los cambios.

---

### Paso 8 — Revisión Final y Refinamiento del Entregable Ejecutivo

**Objetivo:** Revisar el documento ejecutivo completo como si fuera un reporte real para la dirección financiera, evaluar su calidad con apoyo de Copilot, y realizar los ajustes finales para garantizar claridad, precisión y utilidad directiva.

#### Instrucciones

1. Con el documento de resumen ejecutivo abierto en Word, seleccione **todo el texto** con `Ctrl + A`.

2. En el panel de Copilot de Word, ingrese el siguiente *prompt* de revisión de calidad:

   ```text
   Revisa el documento completo y proporciona retroalimentación sobre:
   1. Claridad y comprensibilidad para un lector no técnico de nivel directivo
   2. Consistencia en el tono y el nivel de formalidad
   3. Si las recomendaciones son suficientemente concretas y accionables
   4. Si hay tecnicismos o jerga contable que deba simplificarse
   5. Sugerencias para mejorar el impacto ejecutivo del documento
   
   Proporciona la retroalimentación en forma de lista numerada con observaciones
   específicas y sugerencias de mejora.
   ```

3. Revise la retroalimentación de Copilot. Implemente al menos **dos** de las mejoras sugeridas.

4. Solicite a Copilot que genere un **título ejecutivo** para el documento:

   ```text
   Propón 3 opciones de título para este informe ejecutivo de auditoría.
   El título debe ser formal, descriptivo y apropiado para presentarse
   ante la Dirección General y la Dirección Financiera Corporativa.
   Incluye el período auditado (Q4 2024) en el título.
   ```

5. Seleccione el título más apropiado e insértelo en el encabezado del documento.

6. Agregue la fecha del informe, el nombre del área emisora (Contraloría Corporativa / Auditoría Interna) y la clasificación de confidencialidad (**CONFIDENCIAL — USO INTERNO**) en el encabezado o pie de página del documento.

7. Guarde el documento con un nombre descriptivo:
   ```
   Informe_Auditoria_Anomalias_Q4_2024_[SusIniciales].docx
   ```

8. **Revisión final de pares (actividad grupal):** Si el laboratorio se realiza en grupo, intercambie su documento con el compañero de al lado y evalúe su informe usando la siguiente rúbrica de 1 minuto:

   | Criterio | ¿Cumple? (Sí/No) |
   |---|---|
   | El resumen ejecutivo es comprensible sin leer el resto del documento | |
   | Los hallazgos tienen nivel de riesgo claramente asignado | |
   | Las recomendaciones mencionan área responsable y plazo | |
   | El tono es ejecutivo (no técnico ni coloquial) | |
   | El documento tiene fecha, área emisora y clasificación de confidencialidad | |

9. Comparta verbalmente con el grupo **un hallazgo** que le haya parecido más relevante del análisis y **una recomendación** que considere prioritaria para la dirección financiera.

#### Resultado Esperado

El documento ejecutivo está completo, revisado y refinado. Tiene título formal, fecha, clasificación de confidencialidad, y todas las secciones completadas con contenido coherente y profesional. El participante puede defender los hallazgos y recomendaciones ante la dirección financiera.

#### Verificación

- [ ] Se implementaron al menos dos mejoras sugeridas por Copilot en la revisión de calidad.
- [ ] El documento tiene título formal que incluye el período Q4 2024.
- [ ] El encabezado o pie de página incluye fecha, área emisora y clasificación CONFIDENCIAL.
- [ ] El archivo está guardado con el nombre descriptivo indicado.
- [ ] Se completó la rúbrica de revisión de pares (si aplica en modalidad grupal).

---

## 7. Validación y Pruebas del Laboratorio

Al finalizar todos los pasos, realice la siguiente verificación integral para confirmar que el laboratorio fue completado satisfactoriamente:

### Lista de Verificación Final

| Entregable | Criterio de Éxito | ✓ |
|---|---|---|
| `M4_Registros_Contables_Auditoria.xlsx` | Tabla `TablaAuditoria_Q4` con 5 columnas de alerta insertadas | |
| `M4_Registros_Contables_Auditoria.xlsx` | Formato condicional de colores aplicado a filas con anomalías | |
| `M4_Registros_Contables_Auditoria.xlsx` | Hoja `Análisis_Base` con estadísticas descriptivas documentadas | |
| `M4_Registros_Contables_Auditoria.xlsx` | Hoja `Dashboard_Auditoria` con tabla dinámica y gráfico de barras | |
| `Informe_Auditoria_Anomalias_Q4_2024_[Iniciales].docx` | Sección Resumen Ejecutivo ≤ 150 palabras, tono ejecutivo | |
| `Informe_Auditoria_Anomalias_Q4_2024_[Iniciales].docx` | Tabla de Hallazgos con nivel de riesgo (ALTO/MEDIO/BAJO) | |
| `Informe_Auditoria_Anomalias_Q4_2024_[Iniciales].docx` | 5 recomendaciones con área responsable y plazo | |
| `Informe_Auditoria_Anomalias_Q4_2024_[Iniciales].docx` | Título formal, fecha y clasificación CONFIDENCIAL | |

### Prueba de Consistencia de Datos

Realice esta prueba rápida para verificar que las fórmulas de alerta funcionan correctamente:

1. En la hoja `Registros_Q4`, filtre la columna `Alerta_Horario` por el valor "FUERA_HORARIO".
2. Verifique manualmente que al menos 2 de los registros mostrados correspondan efectivamente a fechas de fin de semana o fuera del rango 07:00–20:00.
3. Quite el filtro y repita para `Alerta_Usuario = "NO_AUTORIZADO"`: verifique que al menos un usuario listado **no** aparezca en el archivo `M4_Catalogo_Usuarios_Autorizados.xlsx`.

Si ambas verificaciones son correctas, las fórmulas de detección funcionan adecuadamente.

---

## 8. Solución de Problemas

### Problema 1: El botón Copilot aparece en gris o no está disponible en Excel

**Síntomas:**
- El botón "Copilot" en la cinta de opciones de Excel aparece atenuado (gris) y no responde al hacer clic.
- Al pasar el cursor sobre el botón, aparece un mensaje como "Copilot no está disponible" o "Se requiere una suscripción".
- El panel de Copilot no se abre al hacer clic.

**Causa probable:**
La cuenta de Microsoft 365 con la que se inició sesión en Excel no tiene la licencia de Copilot asignada, o el equipo no está en el canal de actualización mensual de Microsoft 365 Apps (puede estar en el canal semestral que tiene funcionalidades con retraso), o el archivo de Excel no está guardado en OneDrive/SharePoint (Copilot en Excel requiere que el archivo esté en la nube).

**Solución:**

```
Paso 1 — Verificar la cuenta activa:
   Excel → Archivo → Cuenta → confirmar que la cuenta mostrada es
   la cuenta corporativa con licencia Copilot asignada.
   Si no es la cuenta correcta: cerrar sesión e iniciar con la cuenta corporativa.

Paso 2 — Verificar el canal de actualización:
   Excel → Archivo → Cuenta → Opciones de actualización →
   confirmar "Canal actual (mensual)".
   Si está en canal semestral, contactar al administrador de TI para cambiar el canal.

Paso 3 — Guardar el archivo en la nube:
   Excel → Archivo → Guardar como → OneDrive - [Nombre Organización]
   Guardar el archivo en OneDrive corporativo (no en disco local).
   Cerrar y reabrir el archivo desde OneDrive.

Paso 4 — Si el problema persiste:
   Contactar al administrador de TI para verificar que la licencia
   Microsoft 365 Copilot esté asignada al usuario en el portal de administración
   de Microsoft 365 (admin.microsoft.com → Usuarios → Licencias).
```

---

### Problema 2: Copilot genera fórmulas incorrectas o con errores (#VALUE!, #REF!) al insertar columnas de alerta

**Síntomas:**
- Al insertar una columna de alerta propuesta por Copilot, las celdas muestran errores como `#VALUE!`, `#REF!`, `#NAME?` o `#N/A`.
- Los valores en la columna no corresponden a lo esperado (por ejemplo, todas las celdas muestran "NORMAL" aunque debería haber alertas).
- La fórmula propuesta por Copilot usa nombres de columnas que no coinciden exactamente con los encabezados de la tabla.

**Causa probable:**
Los nombres de columna en la tabla de Excel tienen espacios, caracteres especiales, acentos o mayúsculas/minúsculas que no coinciden exactamente con los que Copilot interpretó del *prompt*. También puede ocurrir que el tipo de dato de la columna `Fecha_Registro` no sea reconocido como fecha-hora por Excel (está almacenado como texto), lo que impide que las funciones `WEEKDAY()` o `HOUR()` funcionen correctamente.

**Solución:**

```
Paso 1 — Verificar nombres de columna exactos:
   Haga clic en el encabezado de la columna problemática.
   En la barra de fórmulas, confirme el nombre exacto (incluyendo
   espacios, acentos y mayúsculas).
   En el siguiente prompt a Copilot, use el nombre exacto entre comillas:
   "...la columna [Fecha_Registro] de la tabla TablaAuditoria_Q4..."

Paso 2 — Verificar formato de fecha-hora:
   Seleccione una celda de la columna Fecha_Registro.
   Cinta Inicio → Número → verificar que diga "Fecha" o "Fecha y hora",
   NO "Texto" o "General".
   Si dice "Texto": seleccionar toda la columna → Datos → Texto en columnas
   → Delimitado → Siguiente → Siguiente → seleccionar "Fecha: DMY" → Finalizar.

Paso 3 — Solicitar a Copilot que revise la fórmula:
   Si la fórmula tiene errores, ingrese en Copilot:
   "La fórmula que generaste para la columna Alerta_Horario está
   produciendo errores #VALUE!. El formato de la columna Fecha_Registro
   es [fecha y hora en formato DD/MM/YYYY HH:MM]. Por favor revisa
   y corrige la fórmula."

Paso 4 — Alternativa manual para WEEKDAY:
   Si Copilot no puede resolver el error, use manualmente:
   =IF(OR(WEEKDAY([Fecha_Registro],2)>5, HOUR([Fecha_Registro])<7,
   HOUR([Fecha_Registro])>=20), "FUERA_HORARIO", "NORMAL")
   Ajustando [Fecha_Registro] al nombre exacto de la columna.
```

---

## 9. Limpieza del Entorno

Una vez completado el laboratorio, realice los siguientes pasos de limpieza para mantener el entorno ordenado y proteger la confidencialidad de los archivos de práctica:

1. **Guardar todos los archivos** con `Ctrl + S` en Excel y Word antes de cerrar.

2. **Cerrar los archivos auxiliares** que fueron abiertos para consulta:
   ```
   Cerrar: M4_Catalogo_Usuarios_Autorizados.xlsx
   Cerrar: M4_Limites_Autorizacion.xlsx
   (Sin guardar cambios si no se modificaron)
   ```

3. **Verificar que los entregables finales estén guardados** en la carpeta de trabajo:
   ```
   C:\Capacitacion_M365Copilot\Modulo4\
   ├── M4_Registros_Contables_Auditoria.xlsx  (con análisis completado)
   ├── Informe_Auditoria_Anomalias_Q4_2024_[SusIniciales].docx
   └── (archivos originales sin modificar)
   ```

4. **No subir archivos** a servicios de almacenamiento en la nube personales (Google Drive, Dropbox, etc.). Los archivos deben permanecer en OneDrive corporativo o en la carpeta local designada por el instructor.

5. **Cerrar el panel de Copilot** en Excel y Word haciendo clic en la **X** del panel lateral.

6. Si el laboratorio se realizó en un equipo compartido o de sala de cómputo, **cerrar sesión** de Microsoft 365:
   ```
   Excel → Archivo → Cuenta → Cerrar sesión
   Word → Archivo → Cuenta → Cerrar sesión
   ```

7. **Eliminar archivos temporales** si el instructor lo indica (solo si el equipo es compartido y los archivos fueron proporcionados como copia temporal).

---

## 10. Resumen y Recursos Adicionales

### Resumen del Laboratorio

En este laboratorio de culminación del curso, los participantes ejecutaron un flujo completo de contraloría corporativa asistida por inteligencia artificial, integrando las competencias desarrolladas en los módulos anteriores:

| Fase | Actividad Realizada | Herramienta |
|---|---|---|
| Preparación de datos | Estructuración de registros como tabla Excel con nombre formal | Excel 365 |
| Análisis exploratorio | Estadísticas descriptivas y distribución temporal con Copilot | Copilot en Excel |
| Detección de anomalías | 5 tipos de alerta: horario, usuarios, límites, fraccionamiento, duplicados | Copilot en Excel |
| Visualización | Formato condicional, tabla dinámica y gráfico de barras | Excel + Copilot |
| Informe ejecutivo | Redacción de resumen ejecutivo, hallazgos y recomendaciones | Copilot en Word |
| Revisión de calidad | Retroalimentación de Copilot y refinamiento del entregable | Copilot en Word |

### Conceptos Clave Aplicados

- **Tablas de Excel como requisito** para el análisis efectivo de Copilot: encabezados claros, tipos de dato consistentes, sin filas en blanco.
- **Prompts específicos y contextualizados**: la calidad del análisis de Copilot depende directamente de la precisión y el contexto de las instrucciones.
- **Técnicas estadísticas aplicadas**: comparaciones condicionales, `COUNTIFS`, `SUMIFS`, análisis de percentiles y desviaciones del comportamiento esperado.
- **Flujo de auditoría completo**: de los datos crudos al entregable ejecutivo, sin necesidad de programación manual.
- **Rol del profesional contable**: Copilot amplifica la capacidad analítica, pero el juicio profesional, la interpretación y la responsabilidad de los hallazgos son siempre humanos.

### Recursos Adicionales

| Recurso | Descripción | Enlace |
|---|---|---|
| Documentación oficial de Copilot en Excel | Guía de Microsoft sobre análisis de datos con Copilot | [support.microsoft.com](https://support.microsoft.com/es-es/topic/copilot-en-excel-a72fd749-49a1-4b6c-bef9-dc191e3c8ea7) |
| Microsoft 365 Copilot — Visión general | Documentación técnica y casos de uso empresariales | [learn.microsoft.com](https://learn.microsoft.com/es-es/copilot/microsoft-365/microsoft-365-copilot-overview) |
| Ley de Benford en auditoría forense | Aplicación en detección de fraude financiero | [Journal of Accountancy](https://www.journalofaccountancy.com/issues/2017/apr/benfords-law-and-fraud-detection.html) |
| Marco de Auditoría Interna — IIA | Estándares globales de auditoría interna basada en datos | [theiia.org](https://www.theiia.org/en/standards/what-are-the-standards/global-internal-audit-standards/) |
| ISACA — IA en Auditoría y Control Interno | Guía de uso de inteligencia artificial en auditoría | [isaca.org](https://www.isaca.org/resources/isaca-journal/issues/2023/volume-2/using-ai-in-internal-audit) |

### Próximos Pasos Sugeridos

Una vez completado este laboratorio y el curso, se recomienda:

1. **Adaptar los *prompts* de este laboratorio** a los datos reales de su organización (previa autorización del área de seguridad de la información).
2. **Documentar un catálogo interno de *prompts* de auditoría** efectivos para reutilizar en revisiones periódicas.
3. **Establecer un flujo de trabajo recurrente** que integre la exportación de Oracle, el análisis con Copilot en Excel y la generación de reportes en Word como proceso estándar de contraloría.
4. **Compartir los entregables generados** en este laboratorio con el equipo de contraloría como demostración del potencial de Copilot para acelerar los procesos de auditoría interna.

---

> 🏆 **¡Felicitaciones!** Ha completado la Práctica 4 y con ella el curso completo de **Microsoft 365 Copilot para Automatización de Tareas Contables y Financieras**. Ha demostrado la capacidad de integrar inteligencia artificial en un flujo real de contraloría corporativa, desde la gestión de datos hasta la comunicación ejecutiva de hallazgos.

---
LAB_END---
