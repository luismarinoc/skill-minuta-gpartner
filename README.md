# skill-minuta-gpartner

Skill para Claude que genera minutas ejecutivas de **GPartner Consulting** a partir de documentos ClickUp.

## ¿Qué hace?

- Lee el contenido de un documento ClickUp (transcripción de reunión)
- Extrae asistentes, agenda, puntos tratados, acuerdos y próximos pasos
- Genera un `.docx` profesional con:
  - Logo GPartner embebido
  - Colores corporativos (azul `#1B3A6B`, verde `#2D8A4E`, rojo `#C0392B`)
  - Tablas de asistentes, agenda y acuerdos con formato ejecutivo
  - Footer con numeración de páginas

## Estructura

```
skill-minuta-gpartner/
├── SKILL.md                    # Instrucciones para Claude
├── references/
│   └── docx-template.md        # Script Node.js completo con logo embebido
└── README.md
```

## Instalación

```bash
git clone https://github.com/luismarinoc/skill-minuta-gpartner.git ~/.claude/skills/user/crearminutagpartnerc
```

## Uso

Compartir el link de un documento ClickUp a Claude y decir:

> "Genera la minuta de esta reunión" o "crea minuta desde ClickUp"

## Tecnologías

- Node.js + `docx` npm package
- ClickUp MCP (`clickup_list_document_pages`, `clickup_get_document_pages`)
- Logo GPartner embebido en base64 (no requiere archivos externos)
