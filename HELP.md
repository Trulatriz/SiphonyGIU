# PressTech – Siphony GUI Help

Esta ayuda refleja el comportamiento real del proyecto tal como esta implementado ahora mismo. Usala como referencia para preparar carpetas, entender que espera cada modulo y saber que archivos produce cada flujo.

## 1. Que hace la aplicacion

La aplicacion organiza el trabajo por:

- `Paper`: contexto principal.
- `Foam type`: contexto secundario para modulos especificos de una espuma.

Desde la ventana principal tienes estos bloques:

- `SMART COMBINE`: fusiona `DoE`, `Density`, `PDR`, `DSC`, `SEM` y `OC` en un `All_Results_YYYYMMDD.xlsx`.
- `FOAM-SPECIFIC ANALYSIS`: acceso rapido a `PDR`, `OC`, `DSC`, combinador SEM y `Cell wall`.
- `SCATTER PLOTS`: graficas de publicacion a partir de `All_Results`.
- `HEATMAPS`: matrices de correlacion sobre hojas de `All_Results`.
- `PAPER IMAGES`: editores de imagen para `SEM`, `DSC` y `TGA`.

## 2. Archivos de configuracion

La aplicacion guarda estado local en:

- `settings.json`: ultimos archivos, tamano de ventana y ajustes generales.
- `foam_types_config.json`: papers, foam types, asociaciones paper-foam y rutas guardadas por modulo.

Normalmente no hace falta editarlos a mano.

## 3. Estructura recomendada de carpetas

Cuando creas un paper y foam types desde el gestor, el proyecto trabaja mejor con una estructura como esta:

```text
<Paper>/
|- DoE.xlsx
|- Density.xlsx
|- Results/
|- <FoamType>/
|  |- PDR/
|  |  |- Input/
|  |  \- Output/
|  |- DSC/
|  |  |- Input/
|  |  \- Output/
|  |- SEM/
|  |  |- Input/
|  |  \- Output/
|  |- Cell wall/
|  |  |- Input/
|  |  \- Output/
|  |- Open-cell content/
|  |  |- Input/
|  |  \- Output/
|  \- Combine/
|     \- Previous results/
```

Notas:

- `DoE.xlsx` y `Density.xlsx` son de nivel `Paper`.
- `PDR`, `DSC`, `SEM`, `OC` y `Cell wall` son por `Foam type`.
- `Results/` es el destino natural del `All_Results`.

## 4. Flujo recomendado

Orden practico:

1. Crear o seleccionar `Paper` y `Foam type`.
2. Preparar `DoE.xlsx` y `Density.xlsx`.
3. Generar resultados por modulo:
   - `PDR_Results_<Foam>.xlsx`
   - `DSC_Results_<Foam>.xlsx` o `DSC_Results_<Foam>_Tg.xlsx`
   - `SEM_Results_<Foam>.xlsx`
   - `OC_Results_<Foam>.xlsx`
4. Ejecutar `SMART COMBINE`.
5. Analizar el nuevo `All_Results_YYYYMMDD.xlsx` con `Scatter Plots` y `Heatmaps`.
6. Usar los editores de `SEM`, `DSC` y `TGA` para figuras finales de publicacion.

## 5. Requisitos reales de entrada

### 5.1 DoE.xlsx

`SMART COMBINE` trabaja con una hoja por foam type.

Columnas canonicas que forman parte del flujo real:

- `Polymer`
- `Additive`
- `Additive %`
- `Label`
- `m(g)`
- `Water (g)`
- `T (degC)`
- `P CO2 (bar)`
- `Psat (MPa)`
- `t (min)`

Importante:

- `Additive` y `Additive %` si se usan en el combine y si pasan al `All_Results`.
- El codigo tiene compatibilidad por cabecera, pero sigue existiendo logica posicional heredada. No conviene cambiar el orden libremente.
- Si `Psat (MPa)` no existe o esta vacio, se deriva como `P CO2 (bar) / 10`.

### 5.2 Density.xlsx

Se espera un libro con una hoja por foam type.

Columnas usadas o derivadas por el flujo:

- `Label`
- `rho foam (g/cm^3)`
- `rho foam (kg/m^3)`
- `Desvest rho foam (g/cm^3)`
- `Desvest rho foam (kg/m^3)`
- `%DER rho foam (g/cm^3)`
- `rho_r`
- `X`

`OC` ademas usa la hoja de la espuma actual para recuperar densidad y `rho_r`.

### 5.3 PDR

Entrada:

- carpeta con archivos `.csv`

Salida principal:

- archivos procesados `* procesado.xlsx` en `PDR/Output`
- resumen `PDR_Results_<Foam>.xlsx` con hoja `Registros`

Comportamiento relevante:

- puede crear el fichero de resultados si no existe
- trabaja de forma incremental
- intenta no duplicar muestras ya registradas
- en Windows puede usar Excel COM para copiar graficos al resumen; si falla, usa un fallback sin grafico

### 5.4 DSC

Entrada:

- carpeta con archivos `.txt`

Modos:

- `Semicrystalline Analysis`
- `Amorphous Analysis`

Salida:

- `DSC_Results_<Foam>.xlsx`
- `DSC_Results_<Foam>_Tg.xlsx`

Comportamiento relevante:

- extrae datos de la seccion `Results:`
- anade solo muestras nuevas si el Excel ya existe
- para semicristalinos guarda datos de primer calentamiento, enfriamiento y segundo calentamiento
- para amorfos guarda Tg y Dcp

### 5.5 Open-Cell Content (OC)

Entrada:

- carpeta con archivos de picnometria `.xlsx` o `.xls`
- `Density.xlsx` a nivel paper

Salida:

- `OC_Results_<Foam>.xlsx`

Comportamiento relevante:

- permite elegir tipo de espuma/instrumento: `Flexible` o `Rigid`
- calcula o corrige `Vpyc`
- abre una ventana de revision antes de guardar
- permite corregir manualmente el volumen de bolas detectado
- guarda formulas de Excel para `Vext`, `Vext - Vpyc`, `1-rho_r`, `Vext(1-rho_r)` y `%OC`

### 5.6 SEM

Hay dos piezas separadas:

- `SEM Image Editor`: editor de imagenes para calibrar, recortar y anadir elementos graficos.
- `Combine SEM Results`: combina histogramas `histogram_*.xlsx` en un `SEM_Results_<Foam>.xlsx`.

Entrada del combinador:

- carpeta `SEM/Output` con archivos `histogram_*.xlsx`

Salida:

- `SEM_Results_<Foam>.xlsx`

Advertencia:

- si alguna celda sale `NaN`, a veces hace falta abrir el `histogram_*.xlsx` original en Excel y guardarlo manualmente para recalcular su cache.

### 5.7 Cell wall

Entrada:

- carpeta con mascaras binarias `*_binary_mask.png`

Salida:

- histogramas, hojas de trabajo y metricas en `Cell wall/Output`

Funciones relevantes:

- escaneo de mascaras
- editor de inclusiones y exclusiones
- ajuste de histogramas
- exportacion de tablas combinadas y metricas como `walls_vs_struts`

## 6. Que genera Smart Combine

`SMART COMBINE` fusiona por `Label` normalizada y construye estas columnas canonicas:

- `Polymer`
- `Additive`
- `Additive %`
- `Label`
- `m(g)`
- `Water (g)`
- `T (degC)`
- `P CO2 (bar)`
- `Psat (MPa)`
- `t (min)`
- `Pi (MPa)`
- `Pf (MPa)`
- `PDR (MPa/s)`
- `n SEM images`
- `o (um)`
- `Desvest o (um)`
- `RSD o (%)`
- `Nv (cells*cm^3)`
- `Desvest Nv (cells*cm^3)`
- `RSD Nv (%)`
- `N₀ (nuclei/cm^3)`
- `Desvest N₀ (nuclei/cm^3)`
- `RSD N₀ (%)`
- `rho foam (g/cm^3)`
- `rho foam (kg/m^3)`
- `Desvest rho foam (g/cm^3)`
- `Desvest rho foam (kg/m^3)`
- `%DER rho foam (g/cm^3)`
- `rho_r`
- `Desvest rho_r`
- `RSD rho_r (%)`
- `X`
- `OC (%)`
- `DSC Tm (degC)`
- `DSC Xc (%)`
- `DSC Tg (degC)`

Comportamiento relevante:

- crea `All_Results_YYYYMMDD.xlsx`
- escribe una hoja global `All_Results` y hojas por foam type
- mueve resultados anteriores a `Previous results`
- copia tambien el nuevo resultado a `Previous results`
- calcula `Psat (MPa)` desde `P CO2 (bar)` si hace falta
- calcula `Desvest rho_r` y `RSD rho_r (%)` a partir de `rho foam`, `Desvest rho foam` y `rho_r`
- calcula `N₀ = Nv / rho_r`
- calcula `Desvest N₀` por propagacion de error usando la desviacion de `Nv` y la de `rho_r`
- calcula `RSD N₀ (%)` a partir de `N₀` y `Desvest N₀`

## 7. Scatter plots

Hay dos modulos:

- `Independent vs Dependent`
- `Dependent vs Dependent`

Entrada:

- cualquier `All_Results*.xlsx` con columnas validas

Funciones importantes:

- seleccion por hoja
- filtros y constancy rule
- agrupacion por color y/o forma
- barras de error cuando existe la desviacion correspondiente
- disponibles tambien para `N₀` y `rho_r` si el `All_Results` incluye sus columnas derivadas
- exportacion de figura
- copia al portapapeles en Windows
- exportacion de datos filtrados con JSON de reproduccion

Escalas de eje:

- `X scale`: `Linear` o `Log`
- `Y scale`: `Linear` o `Log`
- son independientes
- si intentas usar log con valores `<= 0` o limites manuales incompatibles, el modulo avisa y vuelve ese eje a lineal
- en escala log, los ticks base se muestran como `1`, `10`, `100` y `1000`; a partir de `10^4` pasan a notacion `10^n`
- `Group` y `Shape by` pueden usar la misma variable al mismo tiempo; por ejemplo, `Water (g)` puede controlar simultaneamente color/linea y marcador
- si `Group` y `Shape by` usan exactamente la misma variable, el grafico muestra una sola leyenda combinada en lugar de dos separadas
- `Psat (MPa)` y `Water (g)` se muestran sin `.0` cuando el valor es entero, tanto en el eje X como en leyendas/categorias

Los tamanos de exportacion estan fijados internamente para mantener consistencia de publicacion.

## 8. Heatmaps

Entrada:

- `All_Results*.xlsx`

Funciones:

- seleccion por hoja
- seleccion multiple de columnas independientes y dependientes
- pueden incluir columnas derivadas como `N₀ (nuclei/cm^3)` y `rho_r`
- metodos:
  - `Spearman`
  - `Pearson`
  - `Distance (dCor)`
- guardado a imagen
- copia al portapapeles en Windows

Solo usa columnas permitidas del modelo de datos del proyecto.

## 9. Editores de imagen

### 9.1 SEM Image Editor

Flujo:

1. abrir imagen
2. calibrar con linea horizontal
3. recortar region
4. configurar elementos
5. guardar imagen

Incluye:

- deshacer / rehacer
- barra de escala
- borde
- overlay opcional de cell size
- overlay opcional de densidad

### 9.2 DSC Image Editor

Entrada:

- un `.txt` de DSC

Salida:

- figuras individuales o exportacion masiva de 4 PNG
- copia al portapapeles

Caracteristicas:

- `600 dpi`
- layout fijo de exportacion
- varias fases del experimento
- copia de coordenadas `Tm/Xc` desde `1st Heating`

### 9.3 TGA Image Editor

Entrada:

- un `.txt` de TGA

Salida:

- figura overlay masa + DTG
- exportacion PNG
- copia al portapapeles

Caracteristicas:

- `600 dpi`
- layout fijo de exportacion
- control de leyenda, linea y etiqueta de `Td`
- inversion opcional del eje X

## 10. Gestion de papers y foams

`FoamTypeManager` guarda:

- lista global de papers
- lista global de foam types
- que foams pertenecen a cada paper
- ruta raiz de cada paper
- rutas guardadas por modulo
- rutas de `All_Results`

Consejos:

- cambia de `Paper` antes de trabajar con archivos
- cambia de `Foam type` antes de procesar modulos especificos
- usa `Save Paths` en los modulos si quieres fijar una ruta no estandar

## 11. Particularidades importantes

- El proyecto esta claramente orientado a Windows en varias funciones de portapapeles y automatizacion de Excel.
- Para copiar figuras al portapapeles necesitas `pywin32`.
- `PDR` y `DSC` son incrementales: no vuelven a anadir automaticamente muestras ya registradas.
- `SEM Image Editor` no genera por si solo `SEM_Results_<Foam>.xlsx`; eso lo hace el combinador de histogramas o la herramienta externa de analisis.
- `SMART COMBINE` funciona mejor si mantienes nombres de archivo y estructura de carpetas coherentes con la plantilla del proyecto.

## 12. Problemas frecuentes

- `No usable sheets` en heatmaps:
  la hoja no tiene suficientes columnas validas y numericas.

- advertencias en scatter:
  falta renderizar antes de guardar/copiar o no se cumple la constancy rule.

- log scale no disponible:
  hay valores `0`, negativos o limites manuales no compatibles.

- `NaN` en resultados SEM:
  abre y guarda manualmente el `histogram_*.xlsx` original en Excel.

- `Density file not loaded` en OC:
  revisa que `Density.xlsx` exista y que la hoja tenga el nombre exacto del foam type.

- muestras ausentes o mal alineadas en combine:
  revisa la normalizacion de `Label` en todos los ficheros de entrada.

## 13. Nombres de salida importantes

- `Results/All_Results_YYYYMMDD.xlsx`
- `PDR/Output/PDR_Results_<Foam>.xlsx`
- `DSC/Output/DSC_Results_<Foam>.xlsx`
- `DSC/Output/DSC_Results_<Foam>_Tg.xlsx`
- `SEM/Output/SEM_Results_<Foam>.xlsx`
- `Open-cell content/Output/OC_Results_<Foam>.xlsx`

Si mantienes esas rutas y nombres, el flujo completo funciona con muchas menos correcciones manuales.
