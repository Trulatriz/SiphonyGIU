# Ayuda de PressTech – Siphony GUI

Esta guía recoge todo lo necesario para usar la aplicación sin errores: cómo crear un paper, gestionar tipos de espuma y usar cada módulo, incluyendo los formatos de archivo esperados.

## Visión general
- Todo se hace dentro de la aplicación (sin editar JSON a mano).
- Estructura minimalista por carpeta: solo Input y Output.
- Flujo de trabajo en dos niveles:
  1) Nivel paper (global): Smart Combine y Analysis Results.
  2) Nivel espuma (específico): PDR, OC (Picnometry), DSC y SEM.

## Estructura de carpetas por paper
Al crear un paper desde la app se generará automáticamente:
- Paper root
  - DoE.xlsx y Density.xlsx (plantillas o ficheros base)
  - Por cada polímero/tipo de espuma (p.ej., HDPE, LDPE, PP, PET, PS):
    - PDR/
      - Input/
      - Output/
    - Picnometry/ (para OC)
      - Input/
      - Output/
    - DSC/
      - Input/
      - Output/
    - SEM/
      - Input/
      - Output/

Notas:
- Coloca siempre tus datos de entrada en la subcarpeta Input del módulo correspondiente.
- Los resultados se guardan en Output.

## Gestión de papers y espumas
- Nuevo paper: Menú File > Select Paper… > botón “New Paper” en el selector. Define nombre y tipos de espuma a usar.
- Cambiar paper: File > Select Paper…
- Cambiar tipo de espuma activo: File > Select Foam Type…
- Gestionar espumas del paper: desde el selector de paper/espumas (puedes añadir o borrar espumas que no estén usadas en ningún paper).

## Reglas generales de etiquetas (“Label”)
- Usa la misma etiqueta Label para la misma muestra en todos los ficheros (DoE, PDR, Density, OC, SEM, DSC).
- Evita espacios finales y caracteres raros. La app normaliza mínimamente, pero la coincidencia exacta ayuda a evitar huecos.

## Formatos de archivo requeridos
Asegúrate de respetar nombres de columnas, hojas y formatos para evitar errores.

### 1) DoE.xlsx (nivel paper)
- Un único libro Excel con los DoE de todas las espumas.
- Hoja por polímero/tipo de espuma, con nombre igual al polímero (p. ej., HDPE, LDPE, PP, PET, PS).
- Columnas mínimas por hoja:
  - Label
  - m(g)
  - Water (g)
  - T (ºC)
  - P CO2 (bar)
  - t (min)

### 2) Density.xlsx (nivel paper)
- Un único libro Excel con densidades por polímero.
- Hoja por polímero (nombre de hoja = nombre del polímero).
- Para Combine (resultados globales): se usan, si existen, las columnas siguientes:
  - Label
  - Av Exp ρ foam (g/cm3)
  - Desvest Exp ρ foam (g/cm3)
  - %DER Exp ρ foam (g/cm3)
  - ρr
  - X
  - Porosity (%)
- Para OC (picnometría): como mínimo debe existir por hoja:
  - Label
  - Density (g/cm3)

Sugerencia: si necesitas ambos usos, incluye todas las columnas anteriores en la hoja del polímero.

### 3) PDR – archivos CSV (nivel espuma)
- Coloca los CSV en Paper/Polymer/PDR/Input.
- Formato CSV requerido (cabeceras exactas):
  - Time
  - T1 (ºC)
  - T2 (ºC)
  - P (bar)
- Notas:
  - El separador decimal en P puede ser coma o punto (la app lo gestiona).
  - La app genera/actualiza un Excel “registros_promedios.xlsx” en Output con hoja “Registros” y cabeceras:
    - Filename | Pi (MPa) | Pf (MPa) | PDR (MPa/s) | Chart

### 4) OC – Picnometry (nivel espuma)
- Coloca los ficheros originales de picnometría (.xls o .xlsx) en Paper/Polymer/Picnometry/Input.
- El módulo permite:
  - Selección múltiple tipo Ctrl+Click, con “Select All” y “Select None”.
  - Tabla de revisión: muestra Label, masa, densidad, Vpyc calculado, análisis de comentarios, etc.
  - Corrección manual del volumen de bola si la lectura automática del comentario no es correcta.
- Salida recomendada para compatibilidad con Combine:
  - Archivo Excel en Paper/Polymer/Picnometry/Output
  - Nombre recomendado: Polymer_OC.xlsx (por ejemplo, HDPE_OC.xlsx)
- Columnas que escribe el módulo (pueden variar según datos presentes):
  - Label
  - Density (g/cm3)
  - m (g)
  - Vext (cm3)
  - Vpyc unfixed (cm3)
  - Vpyc (cm3)
  - ρr
  - Vext - Vpyc (cm3)
  - 1-ρr
  - Vext(1-ρr) (cm3)
  - %OC
  - Comment Analysis
- Importante sobre “Comment Analysis”:
  - Si no editas manualmente, se guarda “Original | Calculated” (o “No comment”).
  - Solo se marca “Manual” si haces una edición manual del volumen/bolas en la tabla de revisión.
- Combine leerá la columna “%OC” y la renombrará a “OC (%)”.

### 5) DSC – ficheros de texto (nivel espuma)
- Coloca los .txt en Paper/Polymer/DSC/Input.
- Requisitos mínimos en cada .txt:
  - Debe contener secciones tipo “Sample:” y “Results:” (la app extrae Tg, Tm, Xc según el tipo de polímero y script).
- Salida típica: DSC_[Polymer].xlsx en Paper/Polymer/DSC/Output con, al menos, columnas:
  - Sample
  - Mass (mg)
  - 1st Heat Tg (°C)
  - 1st Heat Δcp (J/gK)
  - 2nd Heat Tg (°C)
  - 2nd Heat Δcp (J/gK)

### 6) SEM – imágenes y/o histogramas (nivel espuma)
- Editor de imágenes SEM en Paper/Polymer/SEM.
- Para Combine, si usas un Excel de histograma, mantén en el archivo de resumen las celdas con estas referencias (si aplica a tu plantilla):
  - L3: Av S3D (µm)
  - M3: Desvest S3D
  - AG3: Av Cell density Nv (cells·cm^3 foamed)
  - AH3: Desvest Cell density Nv

### 7) All_Results.xlsx (nivel paper)
- Salida de Combine. Úsala directamente en Analysis Results.
- Orden de columnas objetivo (resumen):
  - Polymer, Label, m(g), Water (g), T (ºC), P CO2 (bar), t (min)
  - Pi (MPa), Pf (MPa), PDR (MPa/s)
  - n, Av S3D (µm), Desvest S3D, DER S3D (%)
  - Av Cell density Nv (cells·cm^3 foamed), Desvest Cell density Nv
  - Av Exp ρ foam (g/cm3), Desvest Exp ρ foam (g/cm3), %DER Exp ρ foam (g/cm3), ρr, X, Porosity (%)
  - OC (%)
  - DSC Tm (°C), DSC Xc (%)
  - DSC Tg (°C)

## Uso de los módulos (paso a paso)

### A) Smart Combine (global)
1) Abre “⚡ SMART COMBINE”.
2) Selecciona el directorio base del paper (la app detecta subcarpetas por polímero).
3) Revisa/indica rutas de DoE.xlsx, Density.xlsx y, si procede, archivos de PDR/DSC/SEM/OC.
4) Ejecuta la combinación. Se generará/actualizará All_Results.xlsx en el paper.

### B) Analysis Results (global)
1) Abre “📊 ANALYSIS RESULTS”.
2) Selecciona All_Results.xlsx del paper.
3) El módulo realiza limpiezas, análisis y genera un Excel de análisis y gráficos en subcarpetas (con fecha).

### C) PDR (por espuma)
1) Abre “📊 Pressure Drop Rate”.
2) Asegúrate de tener CSV en PDR/Input.
3) Elige o crea “registros_promedios.xlsx” en PDR/Output (hoja “Registros”).
4) Procesa; el módulo añadirá filas nuevas para archivos no procesados aún.

### D) OC – Picnometry (por espuma)
1) Abre “🔓 Open-Cell Content”.
2) Selecciona archivos en Picnometry/Input (Ctrl+Click, Select All/None disponible).
3) Revisa resultados en la tabla:
   - Comprueba “Comment Analysis”. Si el comentario fue mal interpretado, edita manualmente el número/tamaño de bolas.
4) Guarda resultados en Picnometry/Output con nombre recomendado Polymer_OC.xlsx.

### E) DSC (por espuma)
1) Abre “🌡️ DSC Analysis”.
2) Coloca .txt en DSC/Input.
3) Procesa; se generará/actualizará DSC_[Polymer].xlsx en DSC/Output; si ya existe, se añaden solo muestras nuevas.

### F) SEM (por espuma)
1) Abre “🔬 SEM Image Editor”.
2) Sigue las instrucciones en pantalla para editar la foto

## Errores comunes y cómo evitarlos
- Nombres de hojas de Excel: deben coincidir con el nombre del polímero (HDPE, LDPE, PP, PET, PS).
- Nombres de columnas: respeta mayúsculas, paréntesis y unidades EXACTAS.
- Etiquetas Label: deben ser coherentes entre DoE, PDR, Density, OC, SEM, DSC.
- Separadores decimales: en CSV de PDR se admite coma o punto; en Excel usa punto decimal de forma consistente.
- Archivos de salida recomendados:
  - PDR/Output/registros_promedios.xlsx (hoja “Registros”).
  - Picnometry/Output/Polymer_OC.xlsx (uno por polímero) para que Combine lo detecte.
  - DSC/Output/DSC_[Polymer].xlsx.

## Dónde encontrar ejemplos/plantillas
- Carpeta templates/ contiene ejemplos mínimos y recomendaciones de estructura.

---
Si algo no encaja con tus datos reales, ajusta las plantillas para mantener los nombres de columnas y hojas indicados aquí. Así evitarás errores y la app combinará/análisisará todo correctamente.
