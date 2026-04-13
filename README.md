# skill-minuta-gpartner

Skill para Claude que genera minutas ejecutivas de **GPartner Consulting** a partir de documentos ClickUp.

## ¿Qué hace?

- Lee el contenido de un documento ClickUp (transcripción de reunión)
- Extrae asistentes, agenda, puntos tratados, acuerdos y próximos pasos
- Genera un `.docx` profesional con:
  - Logo GPartner leído desde `assets/logo_gpartner.jpg`
  - Colores corporativos (azul `#1B3A6B`, verde `#2D8A4E`, rojo `#C0392B`)
  - Tablas de asistentes, agenda y acuerdos con formato ejecutivo
  - Footer con numeración de páginas

## Estructura

```
skill-minuta-gpartner/
├── SKILL.md                    # Instrucciones para Claude
├── references/
│   └── docx-template.md        # Script Node.js completo
├── assets/
│   └── logo_gpartner.jpg       # Logo GPartner
└── README.md
```

## Instalación

```bash
git clone https://github.com/luismarinoc/skill-minuta-gpartner.git ~/.claude/skills/user/crearminutagpartnerc
```

El logo se busca automáticamente en las siguientes rutas (en orden):
1. `~/.claude/skills/user/crearminutagpartnerc/assets/logo_gpartner.jpg`
2. `/mnt/skills/user/crearminutagpartnerc/assets/logo_gpartner.jpg`
3. `/home/claude/skill-minuta-gpartner/assets/logo_gpartner.jpg`

## Uso

Compartir el link de un documento ClickUp a Claude y decir:

> "Genera la minuta de esta reunión" o "crea minuta desde ClickUp"

## Tecnologías

- Node.js + `docx` npm package
- ClickUp MCP (`clickup_list_document_pages`, `clickup_get_document_pages`)
- Logo GPartner leído desde asset del repo (no base64 embebido)
