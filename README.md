# IAG Transform slides (plugin)

Plugin con una skill que genera presentaciones PowerPoint (.pptx) con el formato corporativo de IAG Transform: plantilla oficial, paleta, Arial, logos, iconos y patrones de slide de referencia.

## Qué incluye
- `skills/iag-slides/SKILL.md` — flujo de trabajo y reglas de la skill.
- `skills/iag-slides/assets/template.pptx` — plantilla oficial (58 layouts, sin slides).
- `skills/iag-slides/scripts/iag_deck.py` — librería de componentes (KPI band, paneles, gráficos nativos, tablas comparativas, tarjetas numeradas, pasos, fondos con foto…).
- `skills/iag-slides/references/` — sistema de diseño, catálogo de patrones, assets y layouts.
- `skills/iag-slides/assets/brand/` — logos (IAG, OpCos y proveedores: Boeing, Collins, Honeywell, Lufthansa Technik, Moog, AJW, Turkish Technic, Air France-KLM E&M), fotos de marca. La skill coloca automáticamente el logo de cualquier empresa que se mencione. `assets/fa-icons/` — 88 iconos blancos.
- `skills/iag-slides/assets/examples/` — dos presentaciones de ejemplo con su vista previa.

## Requisitos
Antes de instalar, comprueba que tienes:
- **Claude Code** (app de escritorio o terminal).
- **Git**: el marketplace se descarga desde GitHub. Windows: `winget install --id Git.Git -e`.
- **Windows y Mac con chip Apple (M1 o posterior): no hace falta instalar nada más.** El plugin incluye su propio Python portátil con las librerías (`skills/iag-slides/runtime/win/` y `runtime/macos-arm64/`); Claude usa el que corresponde a tu ordenador.
- **Mac con Intel o Linux**: necesitan **Python 3** con las librerías (`pip install python-pptx Pillow lxml`). Claude puede instalarlas si se lo pides.
- **LibreOffice** (opcional pero recomendado): permite revisar visualmente las slides antes de entregarlas. Sin él, el .pptx se genera igual pero no se renderiza para revisión.

## Instalación
**Desde la app de escritorio de Claude Code (gestor de plugins):**
1. Abre el gestor de plugins (escribe `/plugin` en la caja de mensajes o usa el menú de plugins).
2. Elige **Añadir marketplace** y, entre las opciones, **GitHub**. Escribe `PRemealntt/iag-slides` y confirma.
3. Abre el marketplace **iag-transform**, localiza **iag-slides** e instálalo con **Instalar por mí** (queda disponible en todos tus proyectos).
4. Abre una **sesión nueva** y ya puedes usarlo.

**Desde la terminal (`claude`):**
```
/plugin marketplace add PRemealntt/iag-slides
/plugin install iag-slides@iag-transform
```
Después reinicia Claude Code o ejecuta `/reload-plugins`.

## Actualización
Cuando se publique una versión nueva (el número de versión sube) hay que hacer dos cosas, en este orden:
1. En el gestor de plugins, **sincroniza el marketplace** `iag-transform` (botón *Sincronizar*). Sin esto, el plugin no sabe que hay algo nuevo.
2. Abre **iag-slides** y pulsa **Actualizar**.

Después abre una **sesión nueva**: las sesiones que ya estaban abiertas siguen con la versión anterior. Si el botón de actualizar no aparece, desinstala el plugin, elimina el marketplace y vuelve a añadirlo.

Por terminal:
```
/plugin marketplace update iag-transform
/plugin update iag-slides@iag-transform
```

## Cómo usarlo
Instala el plugin y pide, por ejemplo: «Hazme unas slides con el formato de IAG Transform sobre …» (adjunta datos, notas o un documento). La skill propone la estructura, genera el .pptx y lo revisa visualmente antes de entregarlo.

## Mantenimiento
- Nueva versión de la plantilla: sustituye `assets/template.pptx` (sin slides) y revisa `references/layouts.md`.
- Nuevo logo de proveedor: guarda un PNG recortado y transparente en `skills/iag-slides/assets/brand/logos/suppliers/<nombre-en-minusculas>.png`; se detecta solo por el nombre.
- Nuevo patrón de referencia: añade la presentación a `assets/examples/`, su preview JPG y una entrada en `references/patterns.md`.
- Los ejemplos incluidos contienen información interna de la compañía: no distribuir fuera de la organización.
