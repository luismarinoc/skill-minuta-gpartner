# Template Node.js para Minuta GPartner

Guardar como `/home/claude/gen_minuta.js` y ejecutar con `node /home/claude/gen_minuta.js`.

> **Logo:** El archivo `logo_gpartner.jpg` existe en `/mnt/skills/user/crearminutagpartnerc/assets/` cuando el skill está instalado. No se sube al repo como binario — se lee desde el filesystem en tiempo de ejecución.

## Variables a reemplazar antes de ejecutar

Sustituir el objeto `data` con los valores extraídos del documento ClickUp.

## Script completo

```javascript
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber
} = require('docx');
const fs = require('fs');

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
  azul: "1B3A6B", verde: "2D8A4E", rojo: "C0392B",
  gris: "4A4A4A", grisBorde: "CCCCCC", azulClaro: "EBF0F8", blanco: "FFFFFF"
};

const PAGE_W = 11906, MARGIN = 1134, CONTENT_W = PAGE_W - MARGIN * 2;

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
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [new TableCell({
      borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.SINGLE, size: 4, color: C.azul }, left: { style: BorderStyle.SINGLE, size: 8, color: C.azul }, right: { style: BorderStyle.NONE } },
      shading: { fill: C.azulClaro, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 180, right: 80 },
      width: { size: CONTENT_W, type: WidthType.DXA },
      children: [para([txt(title.toUpperCase(), { bold: true, size: 18, color: C.azul })])]
    })] })]
  });
}
function cell(content, opts = {}) {
  const { w = 2000, bold = false, size = 19, color = C.gris, align = AlignmentType.LEFT, shade = null, colSpan, borders: brd = border() } = opts;
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, borders: brd,
    shading: shade ? { fill: shade, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER, columnSpan: colSpan,
    children: [para([txt(content, { bold, size, color })], { alignment: align })]
  });
}

// ── Logo — busca JPG en rutas del skill instalado ────────────────────
let logoRun = null;
const logoCandidates = [
  '/mnt/skills/user/crearminutagpartnerc/assets/logo_gpartner.jpg',
  '/mnt/skills/user/crearminutaGpartnerc/assets/logo_gpartner.jpg',
  '/mnt/skills/examples/crearminutagpartnerc/assets/logo_gpartner.jpg',
];
for (const p of logoCandidates) {
  if (fs.existsSync(p)) {
    logoRun = new ImageRun({ data: fs.readFileSync(p), transformation: { width: 170, height: 30 }, type: "jpg" });
    break;
  }
}

// ── Header ───────────────────────────────────────────────────────────
const w45 = Math.round(CONTENT_W * 0.45), w55 = Math.round(CONTENT_W * 0.55);
const headerTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [w45, w55],
  rows: [new TableRow({ children: [
    new TableCell({ borders: noBorder(), width: { size: w45, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children: [para(logoRun ? [logoRun] : [txt("GPARTNER CONSULTING", { bold: true, size: 22, color: C.azul })], { alignment: AlignmentType.LEFT })] }),
    new TableCell({ borders: noBorder(), width: { size: w55, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children: [
        para([txt("MINUTA DE REUNIÓN", { bold: true, size: 20, color: C.azul })], { alignment: AlignmentType.RIGHT }),
        para([txt("Documento confidencial", { size: 16, color: C.gris, italics: true })], { alignment: AlignmentType.RIGHT }),
        para([txt("GPARTNER CONSULTING", { size: 16, color: C.gris })], { alignment: AlignmentType.RIGHT }),
      ] }),
  ] })]
});

const divider = para([], { border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.azul, space: 1 } }, spacing: { before: 80, after: 160 } });

const w3 = Math.round(CONTENT_W / 3);
const tricolor = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [w3, w3, CONTENT_W - w3 * 2],
  rows: [new TableRow({ height: { value: 40, rule: "exact" }, children: [
    new TableCell({ borders: noBorder(), shading: { fill: C.azul, type: ShadingType.CLEAR }, width: { size: w3, type: WidthType.DXA }, children: [para([])] }),
    new TableCell({ borders: noBorder(), shading: { fill: C.rojo, type: ShadingType.CLEAR }, width: { size: w3, type: WidthType.DXA }, children: [para([])] }),
    new TableCell({ borders: noBorder(), shading: { fill: C.verde, type: ShadingType.CLEAR }, width: { size: CONTENT_W - w3 * 2, type: WidthType.DXA }, children: [para([])] }),
  ] })]
});

const LBL = Math.round(CONTENT_W * 0.13), VAL = Math.round(CONTENT_W * 0.37);
const lc = tx => new TableCell({ width: { size: LBL, type: WidthType.DXA }, borders: border(C.grisBorde), shading: { fill: "F5F7FA", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [para([txt(tx, { bold: true, size: 17, color: C.azul })])] });
const vc = (tx, w = VAL) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: border(C.grisBorde), margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [para([txt(tx, { size: 18, color: C.gris })])] });
const infoTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [LBL, VAL, LBL, VAL],
  rows: [
    new TableRow({ children: [lc("FECHA"), vc(data.fecha), lc("PLATAFORMA"), vc(data.plataforma)] }),
    new TableRow({ children: [lc("FACILITADOR"), vc(data.facilitador), lc("N° ASISTENTES"), vc(`${data.asistentes.length} personas`)] }),
    new TableRow({ children: [lc("REDACTOR"), vc(data.redactor), lc("REF. CLICKUP"), vc(data.ref_clickup)] }),
  ]
});

const CN = 500, CNOM = Math.round((CONTENT_W - 500) * 0.35), CCARG = Math.round((CONTENT_W - 500) * 0.35), CEMP = CONTENT_W - 500 - CNOM - CCARG;
const astTbl = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CN, CNOM, CCARG, CEMP],
  rows: [
    new TableRow({ tableHeader: true, children: [cell("#", { w: CN, bold: true, size: 18, color: C.blanco, shade: C.azul, align: AlignmentType.CENTER }), cell("Nombre", { w: CNOM, bold: true, size: 18, color: C.blanco, shade: C.azul }), cell("Cargo / Rol", { w: CCARG, bold: true, size: 18, color: C.blanco, shade: C.azul }), cell("Empresa", { w: CEMP, bold: true, size: 18, color: C.blanco, shade: C.azul })] }),
    ...data.asistentes.map((a, i) => new TableRow({ children: [cell(String(i + 1), { w: CN, bold: true, size: 18, align: AlignmentType.CENTER, shade: i % 2 === 1 ? C.azulClaro : null }), cell(a.nombre, { w: CNOM, bold: true, size: 18, shade: i % 2 === 1 ? C.azulClaro : null }), cell(a.cargo, { w: CCARG, size: 18, shade: i % 2 === 1 ? C.azulClaro : null }), cell(a.empresa, { w: CEMP, size: 18, shade: i % 2 === 1 ? C.azulClaro : null })] }))
  ]
});

const agItems = data.agenda.map((item, i) => new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [460, CONTENT_W - 460],
  rows: [new TableRow({ children: [
    new TableCell({ width: { size: 460, type: WidthType.DXA }, borders: noBorder(), margins: { top: 60, bottom: 60, left: 0, right: 100 }, shading: { fill: C.azul, type: ShadingType.CLEAR }, verticalAlign: VerticalAlign.CENTER, children: [para([txt(String(i + 1), { bold: true, size: 18, color: C.blanco })], { alignment: AlignmentType.CENTER })] }),
    new TableCell({ width: { size: CONTENT_W - 460, type: WidthType.DXA }, borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.SINGLE, size: 1, color: C.grisBorde }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, margins: { top: 80, bottom: 80, left: 160, right: 0 }, children: [para([txt(item, { size: 19, color: C.gris })])] }),
  ] })]
}));

const puntosEl = data.puntos_tratados.flatMap((pt, i) => [
  para([txt(`${i + 1}. `, { bold: true, size: 20, color: C.azul }), txt(pt.titulo, { bold: true, size: 20, color: C.azul })], { spacing: { before: 160, after: 60 } }),
  para([txt(pt.desarrollo, { size: 19, color: C.gris })], { spacing: { before: 0, after: 60 }, indent: { left: 360 } }),
  ...(pt.resultado ? [para([txt("Resultado: ", { bold: true, size: 18, color: C.verde }), txt(pt.resultado, { size: 18, color: C.gris, italics: true })], { spacing: { before: 40, after: 80 }, indent: { left: 360 } })] : [])
]);

const AN = 400, AD = Math.round((CONTENT_W - 400) * 0.50), AR = Math.round((CONTENT_W - 400) * 0.28), AP = CONTENT_W - 400 - AD - AR;
const acTbl = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [AN, AD, AR, AP],
  rows: [
    new TableRow({ tableHeader: true, children: [cell("#", { w: AN, bold: true, size: 18, color: C.blanco, shade: C.azul, align: AlignmentType.CENTER }), cell("Acuerdo / Tarea", { w: AD, bold: true, size: 18, color: C.blanco, shade: C.azul }), cell("Responsable", { w: AR, bold: true, size: 18, color: C.blanco, shade: C.azul }), cell("Plazo", { w: AP, bold: true, size: 18, color: C.blanco, shade: C.azul })] }),
    ...data.acuerdos.map((a, i) => new TableRow({ children: [cell(String(a.num), { w: AN, bold: true, size: 18, align: AlignmentType.CENTER, shade: i % 2 === 1 ? C.azulClaro : null }), cell(a.descripcion, { w: AD, size: 18, shade: i % 2 === 1 ? C.azulClaro : null }), cell(a.responsable || "—", { w: AR, size: 18, shade: i % 2 === 1 ? C.azulClaro : null }), cell(a.plazo || "—", { w: AP, size: 18, shade: i % 2 === 1 ? C.azulClaro : null })] }))
  ]
});

const footer = new Footer({ children: [para([txt("GPartner Consulting  —  Documento Confidencial  |  ", { size: 16, color: C.gris, italics: true }), new TextRun({ children: ["Página ", PageNumber.CURRENT, " de ", PageNumber.TOTAL_PAGES], font: "Arial", size: 16, color: C.gris })], { alignment: AlignmentType.CENTER, border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.azul, space: 4 } } })] });

const children = [
  headerTable, divider,
  para([txt(data.titulo, { bold: true, size: 36, color: C.azul })], { alignment: AlignmentType.CENTER, spacing: { before: 160, after: 80 } }),
  para([txt(`${data.subtitulo} · ${data.proyecto}`, { size: 20, color: C.gris })], { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 160 } }),
  tricolor, spacer(1), infoTable, spacer(1),
  sectionHeader("Asistentes"), spacer(0.5), astTbl, spacer(1),
  sectionHeader("Agenda"), spacer(0.5),
  ...agItems.flatMap(t => [t, spacer(0.3)]), spacer(0.5),
  sectionHeader("Desarrollo de la Reunión"), spacer(0.5),
  ...puntosEl, spacer(1),
  sectionHeader("Acuerdos y Compromisos"), spacer(0.5), acTbl, spacer(1),
];
if (data.proximos_pasos) { children.push(sectionHeader("Próximos Pasos / Cierre"), spacer(0.5), para([txt(data.proximos_pasos, { size: 19, color: C.gris })], { indent: { left: 360 } }), spacer(1)); }
if (data.observaciones) { children.push(sectionHeader("Observaciones"), spacer(0.5), para([txt(data.observaciones, { size: 19, color: C.gris })], { indent: { left: 360 } })); }

const doc = new Document({ sections: [{ properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } } }, headers: { default: new Header({ children: [] }) }, footers: { default: footer }, children }] });
const out = process.argv[2] || '/mnt/user-data/outputs/minuta_gpartner.docx';
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(out, buf); console.log("✓ Minuta generada:", out); }).catch(e => { console.error(e); process.exit(1); });
```

## Instrucciones de uso

1. Instalar dependencia:
   ```bash
   npm list -g docx 2>/dev/null | grep docx || npm install -g docx
   ```

2. Copiar el script completo, reemplazar el objeto `data`, y guardar en `/home/claude/gen_minuta.js`

3. Ejecutar:
   ```bash
   node /home/claude/gen_minuta.js /mnt/user-data/outputs/minuta_gpartner.docx
   ```

4. Validar:
   ```bash
   python3 /mnt/skills/public/docx/scripts/office/validate.py /mnt/user-data/outputs/minuta_gpartner.docx
   ```
