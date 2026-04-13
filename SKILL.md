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

```javascript
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak
} = require('docx');
const fs = require('fs');

// ── LOGO GPARTNER — lee desde assets del skill instalado ──────────────
let _logoData = null;
const _logoPaths = [
  '/mnt/skills/user/crearminutagpartnerc/assets/logo_gpartner.jpg',
  '/mnt/skills/examples/crearminutagpartnerc/assets/logo_gpartner.jpg',
  '/home/claude/skill-minuta-gpartner/assets/logo_gpartner.jpg',
];
for (const _p of _logoPaths) { if (fs.existsSync(_p)) { _logoData = fs.readFileSync(_p); break; } }

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
// ────────────────────────────────────────────────────────────────────

const C = {
  azul:"1B3A6B", verde:"2D8A4E", rojo:"C0392B",
  gris:"4A4A4A", grisBorde:"CCCCCC", azulClaro:"EBF0F8", blanco:"FFFFFF"
};
const PAGE_W=11906, MARGIN=1134, CONTENT_W=PAGE_W-MARGIN*2;

const brd=(c=C.grisBorde)=>({top:{style:BorderStyle.SINGLE,size:1,color:c},bottom:{style:BorderStyle.SINGLE,size:1,color:c},left:{style:BorderStyle.SINGLE,size:1,color:c},right:{style:BorderStyle.SINGLE,size:1,color:c}});
const noBrd=()=>({top:{style:BorderStyle.NONE,size:0,color:C.blanco},bottom:{style:BorderStyle.NONE,size:0,color:C.blanco},left:{style:BorderStyle.NONE,size:0,color:C.blanco},right:{style:BorderStyle.NONE,size:0,color:C.blanco}});
const t=(text,o={})=>new TextRun({text,font:"Arial",size:o.size||20,...o});
const p=(ch,o={})=>new Paragraph({children:Array.isArray(ch)?ch:[ch],...o});
const sp=(n=1)=>p([t("")],{spacing:{before:0,after:n*80}});

function secHdr(title){
  return new Table({width:{size:CONTENT_W,type:WidthType.DXA},columnWidths:[CONTENT_W],rows:[new TableRow({children:[new TableCell({borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.SINGLE,size:4,color:C.azul},left:{style:BorderStyle.SINGLE,size:8,color:C.azul},right:{style:BorderStyle.NONE}},shading:{fill:C.azulClaro,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:180,right:80},width:{size:CONTENT_W,type:WidthType.DXA},children:[p([t(title.toUpperCase(),{bold:true,size:18,color:C.azul})])]})]})]});
}
function cl(text,o={}){
  const{w=2000,bold=false,size=19,color=C.gris,align=AlignmentType.LEFT,shade=null,colSpan,borders:b=brd()}=o;
  return new TableCell({width:{size:w,type:WidthType.DXA},borders:b,shading:shade?{fill:shade,type:ShadingType.CLEAR}:undefined,margins:{top:80,bottom:80,left:120,right:120},verticalAlign:VerticalAlign.CENTER,columnSpan:colSpan,children:[p([t(text,{bold,size,color})],{alignment:align})]});
}

const logoRun=_logoData ? new ImageRun({data:_logoData,transformation:{width:170,height:30},type:"jpg"}) : null;
const w45=Math.round(CONTENT_W*0.45),w55=Math.round(CONTENT_W*0.55);
const headerTbl=new Table({width:{size:CONTENT_W,type:WidthType.DXA},columnWidths:[w45,w55],rows:[new TableRow({children:[
  new TableCell({borders:noBrd(),width:{size:w45,type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,margins:{top:0,bottom:0,left:0,right:0},children:[p(logoRun?[logoRun]:[t("GPARTNER CONSULTING",{bold:true,size:22,color:C.azul})],{alignment:AlignmentType.LEFT})]}),
  new TableCell({borders:noBrd(),width:{size:w55,type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,margins:{top:0,bottom:0,left:0,right:0},children:[
    p([t("MINUTA DE REUNIÓN",{bold:true,size:20,color:C.azul})],{alignment:AlignmentType.RIGHT}),
    p([t("Documento confidencial",{size:16,color:C.gris,italics:true})],{alignment:AlignmentType.RIGHT}),
    p([t("GPARTNER CONSULTING",{size:16,color:C.gris})],{alignment:AlignmentType.RIGHT}),
  ]})
]})]});

const divider=p([],{border:{bottom:{style:BorderStyle.SINGLE,size:8,color:C.azul,space:1}},spacing:{before:80,after:160}});
const w3=Math.round(CONTENT_W/3);
const tricolor=new Table({width:{size:CONTENT_W,type:WidthType.DXA},columnWidths:[w3,w3,CONTENT_W-w3*2],rows:[new TableRow({height:{value:40,rule:"exact"},children:[
  new TableCell({borders:noBrd(),shading:{fill:C.azul,type:ShadingType.CLEAR},width:{size:w3,type:WidthType.DXA},children:[p([])]}),
  new TableCell({borders:noBrd(),shading:{fill:C.rojo,type:ShadingType.CLEAR},width:{size:w3,type:WidthType.DXA},children:[p([])]}),
  new TableCell({borders:noBrd(),shading:{fill:C.verde,type:ShadingType.CLEAR},width:{size:CONTENT_W-w3*2,type:WidthType.DXA},children:[p([])]}),
]})]});

const LBL=Math.round(CONTENT_W*0.13),VAL=Math.round(CONTENT_W*0.37);
const lc=tx=>new TableCell({width:{size:LBL,type:WidthType.DXA},borders:brd(C.grisBorde),shading:{fill:"F5F7FA",type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},children:[p([t(tx,{bold:true,size:17,color:C.azul})])]});
const vc=(tx,w=VAL)=>new TableCell({width:{size:w,type:WidthType.DXA},borders:brd(C.grisBorde),margins:{top:80,bottom:80,left:120,right:120},children:[p([t(tx,{size:18,color:C.gris})])]});
const infoTbl=new Table({width:{size:CONTENT_W,type:WidthType.DXA},columnWidths:[LBL,VAL,LBL,VAL],rows:[
  new TableRow({children:[lc("FECHA"),vc(data.fecha),lc("PLATAFORMA"),vc(data.plataforma)]}),
  new TableRow({children:[lc("FACILITADOR"),vc(data.facilitador),lc("N° ASISTENTES"),vc(`${data.asistentes.length} personas`)]}),
  new TableRow({children:[lc("REDACTOR"),vc(data.redactor),lc("REF. CLICKUP"),vc(data.ref_clickup)]}),
]});

const CN=500,CNOM=Math.round((CONTENT_W-500)*0.35),CCARG=Math.round((CONTENT_W-500)*0.35),CEMP=CONTENT_W-500-CNOM-CCARG;
const astTbl=new Table({width:{size:CONTENT_W,type:WidthType.DXA},columnWidths:[CN,CNOM,CCARG,CEMP],rows:[
  new TableRow({tableHeader:true,children:[cl("#",{w:CN,bold:true,size:18,color:C.blanco,shade:C.azul,align:AlignmentType.CENTER}),cl("Nombre",{w:CNOM,bold:true,size:18,color:C.blanco,shade:C.azul}),cl("Cargo / Rol",{w:CCARG,bold:true,size:18,color:C.blanco,shade:C.azul}),cl("Empresa",{w:CEMP,bold:true,size:18,color:C.blanco,shade:C.azul})]}),
  ...data.asistentes.map((a,i)=>new TableRow({children:[cl(String(i+1),{w:CN,bold:true,size:18,align:AlignmentType.CENTER,shade:i%2===1?C.azulClaro:null}),cl(a.nombre,{w:CNOM,bold:true,size:18,shade:i%2===1?C.azulClaro:null}),cl(a.cargo,{w:CCARG,size:18,shade:i%2===1?C.azulClaro:null}),cl(a.empresa,{w:CEMP,size:18,shade:i%2===1?C.azulClaro:null})]}))
]});

const agItems=data.agenda.map((item,i)=>new Table({width:{size:CONTENT_W,type:WidthType.DXA},columnWidths:[460,CONTENT_W-460],rows:[new TableRow({children:[
  new TableCell({width:{size:460,type:WidthType.DXA},borders:noBrd(),margins:{top:60,bottom:60,left:0,right:100},shading:{fill:C.azul,type:ShadingType.CLEAR},verticalAlign:VerticalAlign.CENTER,children:[p([t(String(i+1),{bold:true,size:18,color:C.blanco})],{alignment:AlignmentType.CENTER})]}),
  new TableCell({width:{size:CONTENT_W-460,type:WidthType.DXA},borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.SINGLE,size:1,color:C.grisBorde},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},margins:{top:80,bottom:80,left:160,right:0},children:[p([t(item,{size:19,color:C.gris})])]}),
]})]});

const puntosEl=data.puntos_tratados.flatMap((pt,i)=>[
  p([t(`${i+1}. `,{bold:true,size:20,color:C.azul}),t(pt.titulo,{bold:true,size:20,color:C.azul})],{spacing:{before:160,after:60}}),
  p([t(pt.desarrollo,{size:19,color:C.gris})],{spacing:{before:0,after:60},indent:{left:360}}),
  ...(pt.resultado?[p([t("Resultado: ",{bold:true,size:18,color:C.verde}),t(pt.resultado,{size:18,color:C.gris,italics:true})],{spacing:{before:40,after:80},indent:{left:360}})]:[]
)]);

const AN=400,AD=Math.round((CONTENT_W-400)*0.50),AR=Math.round((CONTENT_W-400)*0.28),AP=CONTENT_W-400-AD-AR;
const acTbl=new Table({width:{size:CONTENT_W,type:WidthType.DXA},columnWidths:[AN,AD,AR,AP],rows:[
  new TableRow({tableHeader:true,children:[cl("#",{w:AN,bold:true,size:18,color:C.blanco,shade:C.azul,align:AlignmentType.CENTER}),cl("Acuerdo / Tarea",{w:AD,bold:true,size:18,color:C.blanco,shade:C.azul}),cl("Responsable",{w:AR,bold:true,size:18,color:C.blanco,shade:C.azul}),cl("Plazo",{w:AP,bold:true,size:18,color:C.blanco,shade:C.azul})]}),
  ...data.acuerdos.map((a,i)=>new TableRow({children:[cl(String(a.num),{w:AN,bold:true,size:18,align:AlignmentType.CENTER,shade:i%2===1?C.azulClaro:null}),cl(a.descripcion,{w:AD,size:18,shade:i%2===1?C.azulClaro:null}),cl(a.responsable||"—",{w:AR,size:18,shade:i%2===1?C.azulClaro:null}),cl(a.plazo||"—",{w:AP,size:18,shade:i%2===1?C.azulClaro:null})]}))
]});

const footer=new Footer({children:[p([t("GPartner Consulting  —  Documento Confidencial  |  ",{size:16,color:C.gris,italics:true}),new TextRun({children:["Página ",PageNumber.CURRENT," de ",PageNumber.TOTAL_PAGES],font:"Arial",size:16,color:C.gris})],{alignment:AlignmentType.CENTER,border:{top:{style:BorderStyle.SINGLE,size:4,color:C.azul,space:4}}})]});

const children=[
  headerTbl,divider,
  p([t(data.titulo,{bold:true,size:36,color:C.azul})],{alignment:AlignmentType.CENTER,spacing:{before:160,after:80}}),
  p([t(`${data.subtitulo} · ${data.proyecto}`,{size:20,color:C.gris})],{alignment:AlignmentType.CENTER,spacing:{before:0,after:160}}),
  tricolor,sp(1),infoTbl,sp(1),
  secHdr("Asistentes"),sp(0.5),astTbl,sp(1),
  secHdr("Agenda"),sp(0.5),
  ...agItems.flatMap(t=>[t,sp(0.3)]),sp(0.5),
  new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}}),
  secHdr("Desarrollo de la Reunión"),sp(0.5),
  ...puntosEl,sp(1),
  secHdr("Acuerdos y Compromisos"),sp(0.5),
  acTbl,sp(1),
  ...(data.proximos_pasos?[secHdr("Próximos Pasos / Cierre"),sp(0.5),p([t(data.proximos_pasos,{size:19,color:C.gris})],{indent:{left:360},spacing:{after:80}}),sp(1)]:[]),
  ...(data.observaciones?[secHdr("Observaciones"),sp(0.5),p([t(data.observaciones,{size:19,color:C.gris})],{indent:{left:360}})]:[]),
];

const doc=new Document({sections:[{properties:{page:{size:{width:PAGE_W,height:16838},margin:{top:MARGIN,right:MARGIN,bottom:MARGIN,left:MARGIN}}},headers:{default:new Header({children:[]})},footers:{default:footer},children}]});
const out=process.argv[2]||'/mnt/user-data/outputs/minuta_gpartner.docx';
Packer.toBuffer(doc).then(buf=>{fs.writeFileSync(out,buf);console.log("✓ Minuta generada:",out);}).catch(e=>{console.error(e);process.exit(1);});
```

## PASO 6 — Entregar resultado

```bash
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

## Instalación

```bash
git clone https://github.com/luismarinoc/skill-minuta-gpartner.git ~/.claude/skills/user/crearminutagpartnerc
```
