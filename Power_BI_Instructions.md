# Guía de Integración de Power BI: Encuestas de Satisfacción y NPS (UDLA)

Esta guía detalla paso a paso cómo cargar el modelo de datos consolidado (`NPS_Consolidado.xlsx`) en Power BI Desktop, cómo estructurar las tablas, diseñar el panel y escribir las medidas DAX exactas para replicar el análisis de forma profesional.

---

## 1. Preparación y Carga de Datos

El script de consolidación ha generado el archivo estructurado listo para BI en la raíz de tu proyecto:
`NPS_Consolidado.xlsx`

Este archivo cuenta con tres hojas clave:
1. **`Detalle_Respuestas` (Tabla de Hechos)**: Contiene las 1072 respuestas individuales limpias de estudiantes, con columnas de Facultad, Programa, Periodo, Coordinador, Calificaciones numéricas (Q1 a Q10) en escala 1-7, y la clasificación NPS.
2. **`Resumen_Programas` (Tabla de Resumen)**: Resumen agrupado por Periodo/Facultad/Programa para visualizaciones directas de alto nivel.
3. **`Comentarios` (Tabla de Texto)**: Contiene los 694 comentarios con su respectiva calificación NPS, clasificación de sentimiento y metadatos.

### Paso a paso para importar en Power BI Desktop:
1. Abre **Power BI Desktop**.
2. Haz clic en **Obtener datos** (Get Data) -> **Excel**.
3. Selecciona el archivo [NPS_Consolidado.xlsx](file:///c:/Users/sroman/Desktop/ENCUESTAS%20NPS/NPS_Consolidado.xlsx).
4. En el panel de navegación, marca las tres tablas: `Detalle_Respuestas`, `Resumen_Programas` y `Comentarios`.
5. Haz clic en **Transformar datos** (Power Query) para verificar tipos de datos (asegúrate de que `respondent_id` sea texto y las notas `Q1` a `Q10` sean números enteros).
6. Haz clic en **Cerrar y aplicar**.

---

## 2. Modelado de Datos (Esquema Estrella)

Para habilitar un filtrado interactivo perfecto, ve a la pestaña **Modelado** (Model View) en Power BI y establece las siguientes relaciones (Esquema Estrella):

- **Relación de Periodo y Programa**:
  - Vincula `Detalle_Respuestas[programa]` con `Resumen_Programas[programa]` (Relación de Varios a Varios o de Uno a Varios si creas una dimensión única de Programas).
  - Vincula `Detalle_Respuestas[respondent_id]` con `Comentarios[respondent_id]` (si aplica) o asocia ambas tablas por medio de la tabla de dimensiones común **Programas**.
  > [!TIP]
  > Para un filtrado limpio, mantén las relaciones de filtro en **ambas direcciones** (Both directions) o crea tablas de dimensiones separadas para **Programas** y **Coordinadores** a partir de las columnas de hechos.

---

## 3. Medidas DAX Clave (Copiar y Pegar)

Haz clic derecho en la tabla `Detalle_Respuestas` en Power BI, selecciona **Nueva medida** y pega las siguientes fórmulas DAX fundamentales:

### A. Total de Encuestas Respondidas
```dax
Total Respuestas = COUNT(Detalle_Respuestas[respondent_id])
```

### B. Cantidad de Estudiantes Promotores (Notas 6 y 7)
```dax
Promotores = CALCULATE(
    COUNT(Detalle_Respuestas[respondent_id]),
    Detalle_Respuestas[tipo_nps] = "Promotor"
)
```

### C. Cantidad de Estudiantes Detractores (Notas 1 a 4)
```dax
Detractores = CALCULATE(
    COUNT(Detalle_Respuestas[respondent_id]),
    Detalle_Respuestas[tipo_nps] = "Detractor"
)
```

### D. Score NPS Neto (Fórmula Oficial)
```dax
NPS = 
VAR TotalValidos = CALCULATE(
    COUNT(Detalle_Respuestas[respondent_id]),
    NOT(ISBLANK(Detalle_Respuestas[Q1_Recomendacion_NPS]))
)
RETURN
IF(
    TotalValidos > 0,
    DIVIDE([Promotores] - [Detractores], TotalValidos, 0),
    0
)
```
*(Asegúrate de formatear esta medida como **Porcentaje (%)** en la pestaña de herramientas de medición).*

### E. Promedio de Satisfacción de Contenidos (Q2)
```dax
Satisfaccion Contenidos = AVERAGE(Detalle_Respuestas[Q2_Contenidos])
```

### F. Promedio de Servicio de Coordinación Académica
```dax
Servicio Coordinador = AVERAGE(Detalle_Respuestas[Q5_Servicio_Coordinador])
```

### G. % de Cobertura de la Encuesta
Para habilitar el cálculo de cobertura en Power BI:
1. Crea una tabla simple en Excel o ingresa datos directamente en Power BI llamada **`Matricula`** con las columnas: `Programa`, `Periodo`, y `Alumnos_Totales`.
2. Vincula esta tabla a tu modelo de datos.
3. Escribe la medida DAX de cobertura:
```dax
% Cobertura = 
VAR AlumnosMatriculados = SUM(Matricula[Alumnos_Totales])
RETURN
IF(
    AlumnosMatriculados > 0,
    DIVIDE([Total Respuestas], AlumnosMatriculados, 0),
    0
)
```

---

## 4. Diseño Visual Basado en Colores Corporativos de UDLA

Para que tu reporte de Power BI tenga una apariencia institucional impecable y un diseño premium en sintonía con la marca UDLA, configura los siguientes colores en el **Tema de Power BI**:

### Paleta de Colores HEX:
*   **Color Primario (Acentos/KPIs):** Naranja UDLA -> `#F05A24`
*   **Color Secundario (Fondos/Títulos):** Azul Marino -> `#1B365D`
*   **Color Neutro Oscuro (Texto):** Gris Oscuro / Slate -> `#334155`
*   **Color Neutro Claro (Fondos de Tarjetas):** Gris Claro -> `#F8FAFC`
*   **Sentimiento Positivo:** Verde Promotor -> `#10B981`
*   **Sentimiento Negativo:** Rojo Detractor -> `#EF4444`

### Visualizaciones Recomendadas en tu Lienzo:
1. **Tarjetas de KPI (Visual de Tarjeta Nueva)**:
   - Coloca en grande el **NPS neto** (ej. `+72%`) y el **Total de Respuestas** (`1.072`).
   - Usa bordes suaves de tarjeta (Radio de esquina de 10px) y un fondo gris muy claro (`#F8FAFC`).
2. **Segmentadores (Slicers)**:
   - Configura selectores en cascada para **Facultad**, **Programa**, **Periodo** y **Coordinador** en el panel izquierdo.
3. **Gráfico de Evolución de Periodos**:
   - Inserta un **Gráfico de líneas y columnas agrupadas**.
   - Eje X: `periodo`. Eje Y de columnas: `Total Respuestas`. Eje Y de líneas: `NPS` y `Satisfaccion Contenidos` (con dos escalas independientes).
4. **Tabla de Calificaciones por Pregunta**:
   - Usa un **Gráfico de barras horizontales** para mostrar el promedio de las preguntas `Q1` a `Q10`.
   - Utiliza **formato condicional** para rellenar las barras automáticamente: barras en color Verde (`#10B981`) para notas >= 6.0, Amarillo (`#F59E0B`) para notas entre 5.0 y 5.9, y Naranja UDLA (`#F05A24`) para notas inferiores a 5.0.

---

## 5. Publicación del Dashboard

1. En Power BI Desktop, haz clic en el botón **Publicar** (Publish) en la cinta superior.
2. Inicia sesión con tu cuenta institucional de UDLA.
3. Selecciona tu **Área de Trabajo** (Workspace).
4. Una vez publicado, ve a [app.powerbi.com](https://app.powerbi.com).
5. Abre el reporte y haz clic en **Archivo** -> **Insertar informe** -> **Publicar en la Web (Público)** o **Insertar en portal seguro**, para obtener el enlace público interactivo que podrás compartir con coordinadores, directores y decanos de facultad.

---

## 6. Automatización de Actualizaciones
Cuando el usuario agregue nuevos archivos Excel y dé doble clic a `Actualizar_Dashboard.bat`, el archivo `NPS_Consolidado.xlsx` se actualizará de inmediato.

*   **En Power BI Desktop**: Solo debes hacer clic en el botón **Actualizar** (Refresh) en la barra de herramientas superior para que los gráficos lean instantáneamente los nuevos datos.
*   **En Power BI Servicio (Web)**: Configura la actualización programada o utiliza la puerta de enlace (Gateway) vinculada a la carpeta de tu computadora para que la web se actualice automáticamente cada vez que corras el actualizador local.
