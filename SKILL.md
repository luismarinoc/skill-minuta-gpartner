---
name: crearminutaGpartnerc
description: >
  Genera minutas de reunión profesionales con branding de GPartner Consulting a partir de un
  link de documento ClickUp. Extrae automáticamente asistentes, agenda, puntos tratados, acuerdos
  y próximos pasos desde la transcripción del documento, y produce un archivo .docx descargable
  con logo oficial GPartner, colores corporativos azul/verde/rojo y formato ejecutivo completo.

  ÚSALA SIEMPRE cuando el usuario mencione: "minuta GPartner", "crear minuta desde ClickUp",
  "generar minuta", "minuta de reunión profesional", "redactar minuta", "minuta con logo",
  o cuando comparta un link de app.clickup.com/docs. También cuando diga "minuta del proyecto"
  o "minuta profesional" en contexto SAP/consultoría GPartner.
compatibility:
  tools:
    - ClickUp MCP (clickup_list_document_pages, clickup_get_document_pages)
    - bash_tool
    - present_files
---

# Skill: crearminutaGpartnerc

Genera minutas ejecutivas de GPartner Consulting desde documentos ClickUp.

## PASO 1 — Extraer IDs del URL ClickUp

Del URL `https://app.clickup.com/{workspace_id}/docs/{doc_id}`:
- `workspace_id` = número antes de `/docs/`  
- `doc_id` = código después de `/docs/`

Ejemplo: `https://app.clickup.com/9013598423/docs/8cm166q-16113`
→ workspace_id=`9013598423`, doc_id=`8cm166q-16113`

## PASO 2 — Obtener contenido del documento

```
clickup_list_document_pages(document_id=doc_id, workspace_id=workspace_id)
```
Luego:
```
clickup_get_document_pages(document_id=doc_id, page_ids=[...todas...], content_format="text/md")
```

## PASO 3 — Extraer datos estructurados

Del contenido, identificar:

| Campo | Descripción |
|-------|-------------|
| `titulo` | Nombre del doc / primera línea |
| `subtitulo` | Tipo de reunión (técnica, directorio, etc.) |
| `proyecto` | Cliente/proyecto (ej: "Proyecto Bash") |
| `fecha` | Fecha en español largo (ej: "Miércoles, 4 de junio de 2026") |
| `hora` | Hora de inicio (ej: "15:00") |
| `plataforma` | Teams, Zoom, presencial, etc. |
| `facilitador` | Quien dirigió |
| `redactor` | Quien levanta minuta (default: "GPartner Consulting") |
| `ref_clickup` | "Doc " + doc_id |
| `asistentes` | [{nombre, cargo, empresa}] — inferir empresa del contexto |
| `agenda` | [string] puntos de agenda |
| `puntos_tratados` | [{titulo, desarrollo, resultado}] |
| `acuerdos` | [{num, descripcion, responsable, plazo}] |
| `proximos_pasos` | Siguiente reunión o acciones |
| `observaciones` | Notas adicionales |

**Si la reunión fue muy breve o sin decisiones explícitas:** igualmente generar minuta con lo disponible, indicar "Sin decisiones formales" en acuerdos y documentar el contexto.

## PASO 4 — Generar el .docx

### Setup rápido
```bash
npm list -g docx 2>/dev/null | grep docx || npm install -g docx
```

### Escribir script con datos reales
```bash
cat > /home/claude/gen_minuta.js << 'JSEOF'
[PEGAR SCRIPT COMPLETO DEL PASO 5 CON DATOS REEMPLAZADOS]
JSEOF
node /home/claude/gen_minuta.js /mnt/user-data/outputs/minuta_gpartner.docx
```

## PASO 5 — Script Node.js completo (COPIAR Y ADAPTAR)

El logo GPartner está embebido en base64 — **NO depende de archivos externos**.

```javascript
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak
} = require('docx');
const fs = require('fs');

// ── LOGO GPARTNER (embebido en base64, no requiere archivo externo) ──
const LOGO_B64 = "/9j/4AAQSkZJRgABAQAAAQABAAD...";
// [Logo base64 completo — ver archivo references/docx-template.md]
const logoBuffer = Buffer.from(LOGO_B64, 'base64');

// ── DATOS DE LA REUNIÓN — REEMPLAZAR ESTOS VALORES ──────────────────
const data = {
  titulo: "REEMPLAZAR: Título de la reunión",
  subtitulo: "REEMPLAZAR: Tipo de reunión",
  proyecto: "REEMPLAZAR: Nombre del proyecto/cliente",
  fecha: "REEMPLAZAR: Día, DD de mes de YYYY",
  hora: "REEMPLAZAR: HH:MM",
  plataforma: "REEMPLAZAR: Teams / Zoom / Presencial",
  facilitador: "REEMPLAZAR: Nombre",
  redactor: "GPartner Consulting",
  ref_clickup: "REEMPLAZAR: Doc XXXXX",
  asistentes: [
    { nombre: "REEMPLAZAR", cargo: "REEMPLAZAR", empresa: "REEMPLAZAR" }
  ],
  agenda: ["REEMPLAZAR punto 1", "REEMPLAZAR punto 2"],
  puntos_tratados: [
    { titulo: "REEMPLAZAR", desarrollo: "REEMPLAZAR", resultado: "REEMPLAZAR" }
  ],
  acuerdos: [
    { num: 1, descripcion: "REEMPLAZAR", responsable: "REEMPLAZAR", plazo: "REEMPLAZAR" }
  ],
  proximos_pasos: "REEMPLAZAR",
  observaciones: ""
};
// Ver script completo en references/docx-template.md
```

> **Nota:** El script completo con el logo embebido y todo el código está en `references/docx-template.md`.

## PASO 6 — Entregar resultado

```bash
# Validar
python3 /mnt/skills/public/docx/scripts/office/validate.py /mnt/user-data/outputs/minuta_gpartner.docx
```

Luego usar `present_files` para compartir el `.docx`.

## Colores GPartner

| Elemento | Hex |
|----------|-----|
| Azul principal | `#1B3A6B` |
| Verde | `#2D8A4E` |
| Rojo acento | `#C0392B` |
| Gris texto | `#4A4A4A` |

## Instalación en Claude Code

```bash
# Clonar repo
git clone https://github.com/luismarinoc/skill-minuta-gpartner.git /path/to/skills/user/crearminutagpartnerc
```
