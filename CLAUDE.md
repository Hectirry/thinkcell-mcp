# Instrucciones para el Agente — think-cell MCP server

> Servidor MCP (Python, transporte stdio) que envuelve la automatización JSON
> `.ppttc` de think-cell: construye archivos `.ppttc` a partir de datos planos
> y los convierte a PowerPoint (`.pptx`) ejecutando `ppttc.exe`.

## Aprendizajes del Agente (Mejora Continua)

> **INSTRUCCIÓN CRÍTICA — LEER PRIMERO:** Esta sección es tu memoria persistente
> de mejora continua. **Con cada ciclo de ejecución** (al completar una tarea,
> resolver un error, descubrir un patrón, o ajustar un flujo) **y con cada
> actualización de cualquier Markdown** (CLAUDE.md, README, etc.) **debes
> agregar aquí un aprendizaje nuevo** si surgió algo no trivial. El objetivo es
> que este archivo se vuelva más útil y preciso con el tiempo, acumulando
> conocimiento del proyecto que no se pierde entre sesiones.
>
> **Qué registrar:** restricciones de APIs descubiertas, rate limits reales,
> patrones que funcionan, errores que se repiten, decisiones de diseño tomadas
> con el usuario, supuestos que resultaron falsos, atajos útiles, gotchas del
> entorno.
>
> **Qué NO registrar:** detalles efímeros de una sola tarea, información ya
> documentada en el README, cosas triviales derivables del código.
>
> **Formato de cada aprendizaje:**
> ```
> - **YYYY-MM-DD — [Tema corto]:** Descripción en 1-3 líneas. **Por qué
>   importa:** consecuencia práctica o cómo aplicarlo en el futuro.
> ```
>
> **Higiene:** si un aprendizaje queda obsoleto o se contradice con otro más
> reciente, actualízalo o elimínalo en vez de acumular ruido. Mantén la lista
> ordenada por fecha (más recientes arriba). Si superas ~25 entradas, consolida
> las más antiguas.

### Registro de aprendizajes

<!-- Agrega nuevas entradas arriba de esta línea. -->

- **2026-05-15 — Plantilla oficial sin-setup `thinkcell_auto.pptx`:** think-cell
  instala su propia plantilla de automatización con elementos ya nombrados
  (`SlideTitle`, `LeftChartTitle`, `RightChartTitle`, `LeftChart`,
  `RightChart`). Copiada al proyecto en `templates/thinkcell_auto.pptx` y usada
  por el tool `create_auto_deck`. **Por qué importa:** elimina el paso manual
  de abrir PowerPoint y nombrar gráficos; el usuario solo aporta datos.
- **2026-05-15 — Una entrada `.ppttc` = un slide:** cada objeto de nivel
  superior `{"template": ..., "data": [...]}` del array `.ppttc` genera UN
  slide. `build_presentation` mete todos los gráficos en una sola entrada (1
  slide); `create_auto_deck` emite una entrada por slide para decks reales de N
  slides. **Por qué importa:** para N slides hay que emitir N entradas, no
  apilar todo en `data`.
- **2026-05-15 — `validator.py` rechazaba celdas `percentage` válidas:**
  `CELL_TYPE_KEYS` omitía `"percentage"`, pero el esquema oficial
  (`ppttc-schema.json`) y `sample.ppttc` sí lo usan. Corregido: se añadió
  `percentage` a los tipos de celda y a la validación numérica. **Por qué
  importa:** `validate_ppttc` ya no marca como inválidos archivos correctos.
- **2026-05-15 — ppttc.exe sí existe:** think-cell 13.0.35.880 instala
  `C:\Program Files (x86)\think-cell\ppttc.exe` (más `ppttchdl.exe`,
  `tcdiag.exe`, etc.). La automatización `.ppttc` → `.pptx` por línea de
  comandos es real, no hace falta abrir el `.ppttc` a mano. **Por qué importa:**
  `convert_to_pptx` es funcional siempre que el gráfico del template esté
  nombrado.
- **2026-05-15 — Los nombres deben coincidir:** `ppttc.exe` corre con éxito
  (exit 0, genera el `.pptx`) aunque el `chart_name` del `.ppttc` no coincida
  con ningún elemento nombrado en el template; simplemente no rellena nada.
  **Por qué importa:** un "éxito" silencioso no garantiza datos. Verifica que
  el template tenga cada gráfico nombrado (mini-barra de think-cell → campo
  "Name").

---

## Qué es este proyecto

Servidor MCP que expone **7 herramientas** sobre stdio:

| Herramienta | Propósito |
| --- | --- |
| `create_chart` | Construir un `.ppttc` para un solo gráfico. |
| `build_presentation` | Combinar varios gráficos/slides en un `.ppttc`. |
| `create_auto_deck` | Construir un deck multi-slide SIN preparar template. |
| `convert_to_pptx` | Ejecutar `ppttc.exe` para producir el `.pptx`. |
| `validate_ppttc` | Validar estructuralmente un documento `.ppttc`. |
| `list_chart_types` | Describir cada tipo de gráfico soportado. |
| `diagnose_thinkcell` | Diagnosticar por qué la automatización no funciona. |

8 tipos de gráfico: `waterfall`, `bar`, `stacked_bar`, `line`, `scatter`,
`mekko`, `area`, `combo`.

## Arquitectura del código

El proyecto separa responsabilidades en módulos deterministas y testeables —
la lógica de negocio no es probabilística y debe ser consistente:

- **`server.py`** — punto de entrada MCP (FastMCP, stdio). Solo orquesta: valida
  argumentos, llama a los módulos y devuelve dicts estructurados.
- **`autodeck.py`** — construcción del `.ppttc` para `create_auto_deck`. Arma
  el JSON directamente (no usa la librería `thinkcell`, que no expone celdas
  `date`/`percentage`/`fill`) y usa la plantilla oficial incluida
  `templates/thinkcell_auto.pptx`, resuelta de forma absoluta.
- **`charts/`** — un módulo por tipo de gráfico. Cada uno subclasea
  `ChartBuilder` (`charts/base.py`): valida la entrada para ese tipo y produce
  el par `(categories, series)` que espera la librería `thinkcell`.
- **`converter.py`** — wrapper de `ppttc.exe` con manejo de errores
  estructurado; nunca lanza excepción para los fallos documentados.
- **`validator.py`** — validación estructural del esquema `.ppttc` (JSON).
- **`diagnostics.py`** — chequeos **solo-lectura** de registro/archivos de
  Windows para diagnosticar el entorno think-cell. No lanza procesos.

**Cómo funciona la automatización think-cell:** think-cell **no** crea gráficos
de la nada y no expone una API general para "controlar" su add-in. El único
flujo soportado es: un template `.pptx` que ya contiene gráficos think-cell
**nombrados** → un `.ppttc` (JSON) que mapea cada nombre a una tabla de datos →
`ppttc.exe` empuja los datos en los elementos que coinciden por nombre. Si los
nombres no coinciden, `ppttc.exe` corre pero no actualiza nada.

## Principios de operación

1. **Revisa primero si existe la herramienta.** Antes de escribir código nuevo,
   revisa `charts/`, `server.py` y los módulos existentes.
2. **Auto-corrección cuando algo falla.** Lee el error y el stack trace, corrige,
   y vuelve a probar. No ejecutes `convert_to_pptx` repetidamente sin entender
   el fallo: usa `diagnose_thinkcell` primero.
3. **Prueba antes de dar por terminado.** Ejecuta `python tests.py` tras
   cualquier cambio en `charts/`, `validator.py` o `server.py`.
4. **Actualiza la documentación a medida que aprendes.** README y esta sección
   de aprendizajes son documentos vivos. No los descartes; mejóralos.

## Organización de archivos

```
Thinkcell/
  server.py          7 herramientas MCP (FastMCP, stdio)
  autodeck.py        constructor de .ppttc sin-setup (create_auto_deck)
  converter.py       wrapper de ppttc.exe
  validator.py       validación estructural de .ppttc
  diagnostics.py     diagnóstico del entorno think-cell
  tests.py           suite de pruebas (python tests.py)
  charts/            base.py + 8 módulos (uno por tipo) + __init__.py
  templates/         thinkcell_auto.pptx (plantilla oficial pre-nombrada)
  output/            .ppttc / .pptx generados (regenerables, no versionar)
  template.pptx      template de PowerPoint con gráficos think-cell nombrados
  requirements.txt   dependencias (mcp[cli], thinkcell)
  README.md          documentación completa
```

**Principio clave:** los archivos en `output/` siempre pueden borrarse y deben
ser reproducibles ejecutando las herramientas de nuevo, nunca editados a mano.

## Entorno

- **Python:** 3.10 o superior.
- **ppttc.exe:** dentro del directorio de think-cell (por defecto
  `C:\Program Files (x86)\think-cell`). Si think-cell está instalado en otra
  ruta, define la variable de entorno `THINKCELL_DIR`; `converter.py`,
  `diagnostics.py` y `autodeck.py` la respetan.
- **Requiere Windows** con think-cell y PowerPoint instalados para
  `convert_to_pptx` y `diagnose_thinkcell`. La construcción y validación de
  `.ppttc` funciona en cualquier sistema operativo.
- **Plantilla de `create_auto_deck`:** no se versiona. `autodeck.py` la copia
  desde la instalación local de think-cell (`<THINKCELL_DIR>/ppttc/template.pptx`)
  la primera vez que se usa; no se redistribuye el archivo de think-cell.
