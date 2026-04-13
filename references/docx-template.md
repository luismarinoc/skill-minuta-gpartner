# Template Node.js para Minuta GPartner

Guardar como `/home/claude/gen_minuta.js` y ejecutar con `node /home/claude/gen_minuta.js`.

## Variables a reemplazar antes de ejecutar

Sustituir el objeto `data` con los valores extraídos del documento ClickUp.

## Script completo

```javascript
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, HeadingLevel,
  TabStopType
} = require('docx');
const fs = require('fs');
const path = require('path');

// ─── DATOS DE LA REUNIÓN ────────────────────────────────────────────
// REEMPLAZAR ESTOS VALORES con los extraídos del documento ClickUp
const data = {
  titulo: "Título de la Reunión",
  subtitulo: "Reunión técnica de integración",
  proyecto: "Proyecto Cliente",
  fecha: "Miércoles, 4 de junio de 2026",
  hora: "15:00",
  plataforma: "Reunión virtual (grabada)",
  facilitador: "Nombre Facilitador",
  redactor: "GPartner Consulting",
  ref_clickup: "Doc 8cm166q-16113",
  asistentes: [
    { nombre: "Nombre Apellido", cargo: "Cargo / Rol", empresa: "Empresa" }
  ],
  agenda: [
    "Punto de agenda 1",
    "Punto de agenda 2"
  ],
  puntos_tratados: [
    {
      titulo: "Título del punto",
      desarrollo: "Desarrollo de lo discutido en este punto.",
      resultado: "Resultado o conclusión alcanzada."
    }
  ],
  acuerdos: [
    { num: 1, descripcion: "Descripción del acuerdo", responsable: "Nombre", plazo: "DD/MM/YYYY" }
  ],
  proximos_pasos: "Próxima reunión: fecha y modalidad.",
  observaciones: ""
};
// ────────────────────────────────────────────────────────────────────

// Colores GPartner
const C = {
  azul: "1B3A6B",
  verde: "2D8A4E",
  rojo: "C0392B",
  gris: "4A4A4A",
  grisBorde: "CCCCCC",
  azulClaro: "EBF0F8",
  blanco: "FFFFFF"
};

const PAGE_W = 11906; // A4 width DXA
const MARGIN = 1134;  // ~2cm margins
const CONTENT_W = PAGE_W - MARGIN * 2; // 9638 DXA

const border = (color = C.grisBorde) => ({
  top: { style: BorderStyle.SINGLE, size: 1, color },
  bottom: { style: BorderStyle.SINGLE, size: 1, color },
  left: { style: BorderStyle.SINGLE, size: 1, color },
  right: { style: BorderStyle.SINGLE, size: 1, color }
});

const noBorder = () => ({
  top: { style: BorderStyle.NONE, size: 0, color: C.blanco },
  bottom: { style: BorderStyle.NONE, size: 0, color: C.blanco },
  left: { style: BorderStyle.NONE, size: 0, color: C.blanco },
  right: { style: BorderStyle.NONE, size: 0, color: C.blanco }
});

// ── Helpers ──────────────────────────────────────────────────────────
function txt(text, opts = {}) {
  return new TextRun({ text, font: "Arial", size: opts.size || 20, ...opts });
}

function para(children, opts = {}) {
  return new Paragraph({ children: Array.isArray(children) ? children : [children], ...opts });
}

function spacer(lines = 1) {
  return para([txt("")], { spacing: { before: 0, after: lines * 80 } });
}

function sectionHeader(title) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: C.azul },
              left: { style: BorderStyle.SINGLE, size: 8, color: C.azul },
              right: { style: BorderStyle.NONE }
            },
            shading: { fill: C.azulClaro, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 180, right: 80 },
            width: { size: CONTENT_W, type: WidthType.DXA },
            children: [para([
              txt(title.toUpperCase(), { bold: true, size: 18, color: C.azul })
            ])]
          })
        ]
      })
    ]
  });
}

function cell(content, opts = {}) {
  const {
    w = 2000, bold = false, size = 19, color = C.gris,
    align = AlignmentType.LEFT, shade = null, colSpan,
    borders: brd = border()
  } = opts;
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: brd,
    shading: shade ? { fill: shade, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: colSpan,
    children: [para([txt(content, { bold, size, color })], { alignment: align })]
  });
}

// ── Logo ─────────────────────────────────────────────────────────────
// Buscar logo: primero en skill instalado, luego en ruta de trabajo
let logoPath = null;
const logoCandidates = [
  '/mnt/skills/user/crearminutaGpartnerc/assets/logo_gpartner.png',
  '/mnt/skills/examples/crearminutaGpartnerc/assets/logo_gpartner.png',
  '/home/claude/skill-crearminutaGpartnerc/assets/logo_gpartner.png',
];
for (const p of logoCandidates) {
  if (fs.existsSync(p)) { logoPath = p; break; }
}

let logoRun = null;
if (logoPath) {
  const logoData = fs.readFileSync(logoPath);
  logoRun = new ImageRun({
    data: logoData,
    transformation: { width: 160, height: 28 },
    type: "png"
  });
}

// ── Header del documento ─────────────────────────────────────────────
const headerTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [Math.round(CONTENT_W * 0.45), Math.round(CONTENT_W * 0.55)],
  rows: [
    new TableRow({
      children: [
        new TableCell({
          borders: noBorder(),
          width: { size: Math.round(CONTENT_W * 0.45), type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [para(logoRun ? [logoRun] : [txt("GPARTNER CONSULTING", { bold: true, size: 22, color: C.azul })], {
            alignment: AlignmentType.LEFT
          })]
        }),
        new TableCell({
          borders: noBorder(),
          width: { size: Math.round(CONTENT_W * 0.55), type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [
            para([txt("MINUTA DE REUNIÓN", { bold: true, size: 20, color: C.azul })], { alignment: AlignmentType.RIGHT }),
            para([txt("Documento confidencial", { size: 16, color: C.gris, italics: true })], { alignment: AlignmentType.RIGHT }),
            para([txt("GPARTNER CONSULTING", { size: 16, color: C.gris })], { alignment: AlignmentType.RIGHT }),
          ]
        })
      ]
    })
  ]
});

// ── Línea divisoria principal ─────────────────────────────────────────
const divider = para([], {
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.azul, space: 1 } },
  spacing: { before: 80, after: 160 }
});

// ── Línea tricolor ────────────────────────────────────────────────────
const tricolor = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [Math.round(CONTENT_W/3), Math.round(CONTENT_W/3), CONTENT_W - Math.round(CONTENT_W/3)*2],
  rows: [
    new TableRow({
      height: { value: 40, rule: "exact" },
      children: [
        new TableCell({ borders: noBorder(), shading: { fill: C.azul, type: ShadingType.CLEAR }, width: { size: Math.round(CONTENT_W/3), type: WidthType.DXA }, children: [para([])] }),
        new TableCell({ borders: noBorder(), shading: { fill: C.rojo, type: ShadingType.CLEAR }, width: { size: Math.round(CONTENT_W/3), type: WidthType.DXA }, children: [para([])] }),
        new TableCell({ borders: noBorder(), shading: { fill: C.verde, type: ShadingType.CLEAR }, width: { size: CONTENT_W - Math.round(CONTENT_W/3)*2, type: WidthType.DXA }, children: [para([])] }),
      ]
    })
  ]
});

// ── Tabla de datos de reunión ─────────────────────────────────────────
const LABEL_W = Math.round(CONTENT_W * 0.13);
const VAL_W   = Math.round(CONTENT_W * 0.37);

function labelCell(text) {
  return new TableCell({
    width: { size: LABEL_W, type: WidthType.DXA },
    borders: border(C.grisBorde),
    shading: { fill: "F5F7FA", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [para([txt(text, { bold: true, size: 17, color: C.azul })])]
  });
}
function valueCell(text, w = VAL_W) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: border(C.grisBorde),
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [para([txt(text, { size: 18, color: C.gris })])]
  });
}

const infoTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [LABEL_W, VAL_W, LABEL_W, VAL_W],
  rows: [
    new TableRow({ children: [labelCell("FECHA"), valueCell(data.fecha), labelCell("PLATAFORMA"), valueCell(data.plataforma)] }),
    new TableRow({ children: [labelCell("FACILITADOR"), valueCell(data.facilitador), labelCell("N° ASISTENTES"), valueCell(`${data.asistentes.length} personas`)] }),
    new TableRow({ children: [labelCell("REDACTOR"), valueCell(data.redactor), labelCell("REFERENCIA CLICKUP"), valueCell(data.ref_clickup)] }),
  ]
});

// ── Tabla de asistentes ───────────────────────────────────────────────
const COL_NUM = 500, COL_NOM = Math.round((CONTENT_W - 500) * 0.35),
      COL_CARGO = Math.round((CONTENT_W - 500) * 0.35),
      COL_EMP  = CONTENT_W - 500 - Math.round((CONTENT_W-500)*0.35)*2;

const asistentesHeader = new TableRow({
  tableHeader: true,
  children: [
    cell("#",         { w: COL_NUM, bold: true, size: 18, color: C.blanco, shade: C.azul }),
    cell("Nombre",    { w: COL_NOM, bold: true, size: 18, color: C.blanco, shade: C.azul }),
    cell("Cargo / Rol",{ w: COL_CARGO, bold: true, size: 18, color: C.blanco, shade: C.azul }),
    cell("Empresa",   { w: COL_EMP, bold: true, size: 18, color: C.blanco, shade: C.azul }),
  ]
});

const asistentesRows = data.asistentes.map((a, i) =>
  new TableRow({
    children: [
      cell(String(i + 1), { w: COL_NUM, bold: true, size: 18, align: AlignmentType.CENTER, shade: i % 2 === 1 ? C.azulClaro : null }),
      cell(a.nombre,  { w: COL_NOM, bold: true, size: 18, shade: i % 2 === 1 ? C.azulClaro : null }),
      cell(a.cargo,   { w: COL_CARGO, size: 18, shade: i % 2 === 1 ? C.azulClaro : null }),
      cell(a.empresa, { w: COL_EMP, size: 18, shade: i % 2 === 1 ? C.azulClaro : null }),
    ]
  })
);

const asistentesTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [COL_NUM, COL_NOM, COL_CARGO, COL_EMP],
  rows: [asistentesHeader, ...asistentesRows]
});

// ── Agenda ────────────────────────────────────────────────────────────
const agendaItems = data.agenda.map((item, i) =>
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [460, CONTENT_W - 460],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 460, type: WidthType.DXA },
            borders: noBorder(),
            margins: { top: 60, bottom: 60, left: 0, right: 100 },
            children: [para([txt(String(i+1), { bold: true, size: 18, color: C.blanco })], { alignment: AlignmentType.CENTER })],
            shading: { fill: C.azul, type: ShadingType.CLEAR },
            verticalAlign: VerticalAlign.CENTER
          }),
          new TableCell({
            width: { size: CONTENT_W - 460, type: WidthType.DXA },
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.SINGLE, size: 1, color: C.grisBorde }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
            margins: { top: 80, bottom: 80, left: 160, right: 0 },
            children: [para([txt(item, { size: 19, color: C.gris })])]
          })
        ]
      })
    ]
  })
);

// ── Puntos tratados ───────────────────────────────────────────────────
const puntosContent = data.puntos_tratados.flatMap((p, i) => [
  para([
    txt(`${i+1}. `, { bold: true, size: 20, color: C.azul }),
    txt(p.titulo, { bold: true, size: 20, color: C.azul })
  ], { spacing: { before: 160, after: 60 } }),
  para([txt(p.desarrollo, { size: 19, color: C.gris })], { spacing: { before: 0, after: 60 }, indent: { left: 360 } }),
  ...(p.resultado ? [
    para([
      txt("Resultado: ", { bold: true, size: 18, color: C.verde }),
      txt(p.resultado, { size: 18, color: C.gris, italics: true })
    ], { spacing: { before: 40, after: 80 }, indent: { left: 360 } })
  ] : [])
]);

// ── Tabla de acuerdos ─────────────────────────────────────────────────
const AC_NUM = 400, AC_DESC = Math.round((CONTENT_W - 400) * 0.50),
      AC_RESP = Math.round((CONTENT_W - 400) * 0.28),
      AC_PLAZO = CONTENT_W - 400 - Math.round((CONTENT_W-400)*0.50) - Math.round((CONTENT_W-400)*0.28);

const acuerdosHeader = new TableRow({
  tableHeader: true,
  children: [
    cell("#",           { w: AC_NUM,  bold: true, size: 18, color: C.blanco, shade: C.azul, align: AlignmentType.CENTER }),
    cell("Acuerdo / Tarea", { w: AC_DESC, bold: true, size: 18, color: C.blanco, shade: C.azul }),
    cell("Responsable", { w: AC_RESP, bold: true, size: 18, color: C.blanco, shade: C.azul }),
    cell("Plazo",       { w: AC_PLAZO,bold: true, size: 18, color: C.blanco, shade: C.azul }),
  ]
});

const acuerdosRows = data.acuerdos.map((a, i) =>
  new TableRow({
    children: [
      cell(String(a.num), { w: AC_NUM,  bold: true, size: 18, align: AlignmentType.CENTER, shade: i % 2 === 1 ? C.azulClaro : null }),
      cell(a.descripcion, { w: AC_DESC, size: 18, shade: i % 2 === 1 ? C.azulClaro : null }),
      cell(a.responsable || "—", { w: AC_RESP, size: 18, shade: i % 2 === 1 ? C.azulClaro : null }),
      cell(a.plazo || "—",       { w: AC_PLAZO, size: 18, shade: i % 2 === 1 ? C.azulClaro : null }),
    ]
  })
);

const acuerdosTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [AC_NUM, AC_DESC, AC_RESP, AC_PLAZO],
  rows: [acuerdosHeader, ...acuerdosRows]
});

// ── Footer ────────────────────────────────────────────────────────────
const footer = new Footer({
  children: [
    para([
      txt("GPartner Consulting  —  Documento Confidencial", { size: 16, color: C.gris, italics: true }),
      new TextRun({ children: ["   Página ", PageNumber.CURRENT, " de ", PageNumber.TOTAL_PAGES], font: "Arial", size: 16, color: C.gris })
    ], { alignment: AlignmentType.CENTER,
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.azul, space: 4 } }
    })
  ]
});

// ── Armado del documento ──────────────────────────────────────────────
const children = [
  headerTable,
  divider,
  // Título
  para([txt(data.titulo, { bold: true, size: 36, color: C.azul })], {
    alignment: AlignmentType.CENTER, spacing: { before: 160, after: 80 }
  }),
  para([txt(data.subtitulo + (data.proyecto ? ` · ${data.proyecto}` : ""), { size: 20, color: C.gris })], {
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 160 }
  }),
  tricolor,
  spacer(1),
  infoTable,
  spacer(1),
  sectionHeader("Asistentes"),
  spacer(0.5),
  asistentesTable,
  spacer(1),
  sectionHeader("Agenda"),
  spacer(0.5),
  ...agendaItems.flatMap(t => [t, spacer(0.3)]),
  spacer(0.5),
  sectionHeader("Desarrollo de la Reunión"),
  spacer(0.5),
  ...puntosContent,
  spacer(1),
  sectionHeader("Acuerdos y Compromisos"),
  spacer(0.5),
  acuerdosTable,
  spacer(1),
];

if (data.proximos_pasos) {
  children.push(sectionHeader("Próximos Pasos / Cierre"));
  children.push(spacer(0.5));
  children.push(para([txt(data.proximos_pasos, { size: 19, color: C.gris })], { indent: { left: 360 } }));
  children.push(spacer(1));
}

if (data.observaciones) {
  children.push(sectionHeader("Observaciones"));
  children.push(spacer(0.5));
  children.push(para([txt(data.observaciones, { size: 19, color: C.gris })], { indent: { left: 360 } }));
  children.push(spacer(1));
}

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: 16838 }, // A4
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
      }
    },
    headers: { default: new Header({ children: [] }) },
    footers: { default: footer },
    children
  }]
});

const outputPath = process.argv[2] || '/mnt/user-data/outputs/minuta_gpartner.docx';
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outputPath, buf);
  console.log("✓ Minuta generada:", outputPath);
}).catch(e => { console.error("Error:", e); process.exit(1); });
```

## Instrucciones de uso

1. Instalar dependencia (si no está):
   ```bash
   npm install -g docx 2>/dev/null || npm install docx
   ```

2. Escribir el script con los datos extraídos:
   ```bash
   cat > /home/claude/gen_minuta.js << 'JSEOF'
   [PEGAR SCRIPT COMPLETO CON DATOS REEMPLAZADOS]
   JSEOF
   ```

3. Ejecutar:
   ```bash
   node /home/claude/gen_minuta.js /mnt/user-data/outputs/minuta_gpartner.docx
   ```

4. Validar con:
   ```bash
   python /mnt/skills/public/docx/scripts/office/validate.py /mnt/user-data/outputs/minuta_gpartner.docx
   ```
