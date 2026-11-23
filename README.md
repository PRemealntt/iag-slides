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
- **Python 3** con las librerías: `pip install python-pptx Pillow lxml` (en Windows, `py -m pip install python-pptx Pillow lxml`).
- **LibreOffice** (opcional pero recomendado): permite revisar visualmente las slides antes de entregarlas. Sin él, el .pptx se genera igual pero no se renderiza para revisión.

## Instalación
Escribe estos comandos dentro de Claude Code, uno a uno:

```
/plugin marketplace add PRemealntt/iag-slides
/plugin install iag-slides@iag-transform
```

Reinicia Claude Code (o ejecuta `/reload-plugins`) y ya está disponible.

## Actualización
Cuando se publique una versión nueva:

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
