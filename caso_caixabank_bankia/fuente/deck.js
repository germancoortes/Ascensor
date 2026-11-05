const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");
const fs = require("fs");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2];
const NAVY = "0B2545", TEAL = "0F8B8D", AMBER = "F4A259", INK = "1B2433", MUTED = "5C6B7A",
  LIGHT = "EEF3F7", WHITE = "FFFFFF", CORAL = "D9534F", GREEN = "2E9E6B", TEAL_L = "D7EFEF", AMBER_L = "FDEBD9", CORAL_L = "F8DEDD";
const THEME = {
  name: "Fusion RRHH", headFontFace: "Calibri", bodyFontFace: "Calibri",
  colors: { dk1: INK, lt1: WHITE, dk2: NAVY, lt2: LIGHT, accent1: TEAL, accent2: AMBER, accent3: "8DA9C4",
    accent4: CORAL, accent5: GREEN, accent6: MUTED, hlink: TEAL, folHlink: "5C6B7A" },
};

async function icon(name, color, size = 256) {
  const C = md[name]; if (!C) throw new Error("icon " + name);
  const svg = RDS.renderToStaticMarkup(React.createElement(C, { color: "#" + color, size }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const NOTES = [];
(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: "Calibri", bodyFontFace: "Calibri" };
  pres.title = "Fusión CaixaBank–Bankia: APT y planificación de RRHH";
  pres.author = "Grupo de trabajo – Dirección de RRHH II";

  pres.defineSlideMaster({
    title: "CONTENT", background: { color: WHITE },
    objects: [
      { text: { text: "Caso Práctico 1 · Fusión CaixaBank–Bankia · Dirección de RRHH II", options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: MUTED, margin: 0 } } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.62, w: 12.1, h: 0.75, fontSize: 30, bold: true, color: NAVY, margin: 0, valign: "top", align: "left" }, text: "" } },
    ],
    slideNumber: { x: 12.23, y: 6.98, w: 0.5, h: 0.3, fontSize: 10, color: MUTED, align: "right" },
  });
  pres.defineSlideMaster({ title: "DARK", background: { color: NAVY }, objects: [] });

  let n = 0;
  const kicker = (s, t) => s.addText(t, { x: 0.6, y: 0.3, w: 9, h: 0.3, fontSize: 12, bold: true, color: TEAL, charSpacing: 2, margin: 0, isTextBox: true });
  const content = (title, kick, notes, sect) => {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sect });
    s.addText(title, { placeholder: "title" });
    kicker(s, kick); s.addNotes(notes); NOTES.push({ n: ++n, title, notes }); return s;
  };
  const dark = (notes, title, sect) => { const s = pres.addSlide({ masterName: "DARK", sectionTitle: sect }); s.addNotes(notes); NOTES.push({ n: ++n, title, notes }); return s; };
  const shadow = () => ({ type: "outer", color: "000000", opacity: 0.12, blur: 6, offset: 2, angle: 90 });
  const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: INK, valign: "top" }, o));
  const R = (s, o) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ rectRadius: 0.08, line: { type: "none" } }, o));
  const circleIcon = async (s, ic, x, y, d, bg, fg, name) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" }, objectName: name });
    s.addImage({ data: await icon(ic, fg), x: x + d * 0.22, y: y + d * 0.22, w: d * 0.56, h: d * 0.56, objectName: name + "i" });
  };
  const section = (num, title, sub, notes, sect) => {
    pres.addSection({ title: sect });
    const s = dark(notes, title, sect);
    T(s, num, { x: 0.8, y: 1.6, w: 4, h: 1.9, fontSize: 120, bold: true, color: AMBER, objectName: "a1_num" });
    T(s, title, { x: 0.8, y: 3.6, w: 11.5, h: 1.0, fontSize: 40, bold: true, color: WHITE, objectName: "a2_t" });
    T(s, sub, { x: 0.8, y: 4.65, w: 11, h: 1.2, fontSize: 18, color: "BFD0E0", objectName: "a3_s" });
    return s;
  };

  // ---------- 1. PORTADA ----------
  pres.addSection({ title: "Introducción" });
  let s = dark("Buenos días. Presentamos el Caso Práctico 1 de Dirección de Recursos Humanos II: la fusión entre CaixaBank y Bankia, anunciada en septiembre de 2020. Vamos a analizarla desde dos herramientas del Tema 2 —el Análisis de Puestos de Trabajo y la Planificación de Recursos Humanos— y responderemos a las cuatro preguntas del caso.", "Portada", "Introducción");
  T(s, "CASO PRÁCTICO 1 · DIRECCIÓN DE RRHH II · 4º GADE", { x: 0.8, y: 0.9, w: 8, h: 0.35, fontSize: 13, bold: true, color: AMBER, charSpacing: 2, objectName: "a1_k" });
  T(s, "Fusión CaixaBank – Bankia", { x: 0.8, y: 1.5, w: 7.2, h: 1.9, fontSize: 48, bold: true, color: WHITE, valign: "top", objectName: "a2_t" });
  T(s, "Análisis de Puestos de Trabajo y Planificación de Recursos Humanos ante la integración de dos plantillas", { x: 0.8, y: 3.5, w: 6.6, h: 1.1, fontSize: 19, color: "BFD0E0", objectName: "a3_s" });
  T(s, "Universidad de Cádiz · Curso 2026/27", { x: 0.8, y: 6.4, w: 6, h: 0.35, fontSize: 13, color: "8DA9C4", objectName: "a3_u" });
  s.addShape(pres.shapes.OVAL, { x: 7.9, y: 1.55, w: 3.3, h: 3.3, fill: { color: TEAL, transparency: 15 }, line: { type: "none" }, objectName: "a4_c1" });
  s.addShape(pres.shapes.OVAL, { x: 9.75, y: 2.35, w: 2.6, h: 2.6, fill: { color: AMBER, transparency: 15 }, line: { type: "none" }, objectName: "a5_c2" });
  T(s, [{ text: "≈35.600", options: { fontSize: 30, bold: true, breakLine: true } }, { text: "empleados CaixaBank", options: { fontSize: 13 } }], { x: 8.05, y: 2.65, w: 2.3, h: 1.1, color: WHITE, align: "center", objectName: "a4_l1" });
  T(s, [{ text: "≈16.000", options: { fontSize: 26, bold: true, breakLine: true } }, { text: "empleados Bankia", options: { fontSize: 13 } }], { x: 10.55, y: 3.15, w: 1.9, h: 1.0, color: NAVY, align: "center", objectName: "a5_l2" });
  T(s, [{ text: "+51.000 personas", options: { bold: true, fontSize: 22, color: WHITE, breakLine: true } }, { text: "una sola plantilla que hay que analizar y planificar", options: { fontSize: 14, color: "BFD0E0" } }], { x: 7.9, y: 5.2, w: 4.6, h: 0.9, objectName: "a6_tot" });

  // ---------- 2. AGENDA ----------
  s = content("Cuatro preguntas, un mismo hilo: personas en una fusión", "HOJA DE RUTA",
    "Esta es nuestra hoja de ruta. Las cuatro preguntas siguen una lógica: primero conocer los puestos (APT), después cómo hacer ese análisis, luego planificar la plantilla futura y, por último, valorar qué sirve de lo que cada banco ya tenía planificado. Todo se apoya en el Tema 2, Procesos Básicos, y lo conectamos con el resto de temas.", "Introducción");
  const qs = [["01", "El papel del APT en la fusión y en el ajuste de plantilla"], ["02", "Los pasos del APT aplicados a las dos entidades"], ["03", "¿Nueva planificación de RRHH? Sí: pasos a seguir"], ["04", "Utilidad de las planificaciones previas a la fusión"]];
  qs.forEach(([num, q], i) => {
    const x = 0.6 + i * 3.08;
    R(s, { x, y: 1.85, w: 2.85, h: 3.9, fill: { color: i % 2 ? LIGHT : NAVY }, shadow: shadow(), objectName: `a${i + 1}_c` });
    T(s, num, { x: x + 0.3, y: 2.15, w: 2.2, h: 1.0, fontSize: 54, bold: true, color: i % 2 ? TEAL : AMBER, objectName: `a${i + 1}_n` });
    T(s, q, { x: x + 0.3, y: 3.4, w: 2.3, h: 2.1, fontSize: 18, bold: true, color: i % 2 ? NAVY : WHITE, objectName: `a${i + 1}_q` });
  });
  T(s, "Marco teórico: Tema 2 · Procesos básicos (Dessler, 2009; Gómez-Mejía et al., 2008), conectado con los Temas 1, 3, 4 y 5", { x: 0.6, y: 6.1, w: 12, h: 0.4, fontSize: 13, italic: true, color: MUTED, objectName: "a5_f" });

  // ---------- 3. CONTEXTO ----------
  s = content("La operación: nace el mayor banco de España", "CONTEXTO · SEPTIEMBRE 2020",
    "El 17 y 18 de septiembre de 2020 los consejos aprueban la fusión: CaixaBank absorbe a Bankia. La entidad resultante supera los 51.000 empleados y los 664.000 millones de activos. El objetivo económico es claro: unos 770 millones de euros anuales de ahorro de costes por sinergias. Y aquí está la tensión del caso: CaixaBank afirma que las personas son su principal activo, pero las sinergias pasan, en gran parte, por reducir duplicidades de plantilla y oficinas.", "Introducción");
  const stats = [["+51.000", "empleados entre ambas entidades"], ["≈6.300", "oficinas, muchas solapadas"], ["664.000 M€", "en activos: nº 1 en España"], ["770 M€", "ahorro anual de costes previsto"]];
  stats.forEach(([v, l], i) => {
    const x = 0.6 + i * 3.08;
    R(s, { x, y: 1.75, w: 2.85, h: 2.2, fill: { color: LIGHT }, objectName: `a${i + 1}_b` });
    T(s, v, { x: x + 0.25, y: 2.0, w: 2.45, h: 0.9, fontSize: 36, bold: true, color: i === 3 ? AMBER : TEAL, objectName: `a${i + 1}_v`, fit: "shrink" });
    T(s, l, { x: x + 0.25, y: 2.95, w: 2.4, h: 0.8, fontSize: 14, color: INK, objectName: `a${i + 1}_l` });
  });
  R(s, { x: 0.6, y: 4.35, w: 12.13, h: 2.25, fill: { color: NAVY }, objectName: "a5_q" });
  T(s, "“", { x: 0.9, y: 4.3, w: 0.8, h: 1.0, fontSize: 72, bold: true, color: AMBER, objectName: "a5_qq" });
  T(s, "Las personas son el pilar fundamental de su actividad y su principal activo: contar con personas motivadas, preparadas y comprometidas resulta determinante.", { x: 1.7, y: 4.65, w: 10.6, h: 1.1, fontSize: 19, italic: true, color: WHITE, objectName: "a5_qt" });
  T(s, "CaixaBank, según el enunciado del caso  ·  Reto: lograr sinergias sin perder ese activo", { x: 1.7, y: 5.85, w: 10.6, h: 0.4, fontSize: 13, color: AMBER, objectName: "a5_qs" });
  T(s, "Fuentes: comunicado CaixaBank (18/09/2020); Banco de España, IEF 2021; Bolsamanía.", { x: 0.6, y: 6.65, w: 12, h: 0.3, fontSize: 10, color: MUTED });

  // ---------- 4. CRONOLOGÍA ----------
  s = content("Del anuncio al ajuste de plantilla", "CRONOLOGÍA",
    "En septiembre de 2020, cuando se publica la noticia, el tamaño del ajuste aún no estaba decidido: había que analizar solapamientos y negociar con los sindicatos. Lo que ocurrió después confirma la importancia del caso: la fusión jurídica se completa en marzo de 2021, en abril el banco propone 8.291 salidas y, tras negociar, el 1 de julio de 2021 se firman 6.452 salidas voluntarias, un 22 % menos. Es decir, el ajuste se fue ajustando con información y negociación.", "Introducción");
  s.addShape(pres.shapes.LINE, { x: 0.9, y: 3.55, w: 11.5, h: 0, line: { color: "B8C7D6", width: 3 }, objectName: "a1_line" });
  const tl = [["4 sep 2020", "Se confirman las negociaciones de fusión"], ["17–18 sep 2020", "Los consejos aprueban la fusión por absorción"], ["26 mar 2021", "Fusión jurídica efectiva"], ["Abr 2021", "Propuesta inicial: 8.291 salidas"], ["1 jul 2021", "Acuerdo sindical: 6.452 salidas voluntarias"], ["Nov 2021 →", "Salidas escalonadas por ventanas"]];
  tl.forEach(([d, t], i) => {
    const cx = 1.5 + i * 2.05 - 0.73, up = i % 2 === 0;
    s.addShape(pres.shapes.OVAL, { x: cx + 0.55, y: 3.37, w: 0.36, h: 0.36, fill: { color: i >= 3 ? AMBER : TEAL }, line: { color: WHITE, width: 2 }, objectName: `a${i + 2}_dot` });
    T(s, d, { x: cx - 0.25, y: up ? 1.85 : 4.05, w: 2.0, h: 0.4, fontSize: 16, bold: true, color: i >= 3 ? "C77A2E" : TEAL, align: "center", objectName: `a${i + 2}_d` });
    T(s, t, { x: cx - 0.25, y: up ? 2.3 : 4.5, w: 2.0, h: 0.95, fontSize: 14, color: INK, align: "center", objectName: `a${i + 2}_t` });
  });
  R(s, { x: 0.6, y: 5.75, w: 12.13, h: 0.85, fill: { color: AMBER_L }, objectName: "a8_m" });
  T(s, "En la noticia (sep. 2020) el tamaño del ajuste estaba pendiente: primero había que analizar solapamientos y negociar con la representación de los trabajadores.", { x: 0.85, y: 5.87, w: 11.7, h: 0.65, fontSize: 14, color: INK, valign: "middle", objectName: "a8_mt" });

  // ---------- 5. SECCIÓN 1 ----------
  section("01", "El papel del APT en la fusión", "¿Qué aporta el Análisis de Puestos de Trabajo y sirve para ajustar la plantilla?",
    "Primera pregunta: ¿qué papel juega el Análisis de Puestos de Trabajo en esta fusión y si sirve para el ajuste de plantilla?", "P1 · APT");

  // ---------- 6. QUÉ ES ----------
  s = content("El APT: qué es y qué información aporta", "PREGUNTA 1 · MARCO TEÓRICO",
    "Según el Tema 2, el APT es el procedimiento para establecer las obligaciones y habilidades que requiere un puesto y el tipo de persona que debe ocuparlo. Produce dos documentos: la descripción —qué se hace, cómo y en qué condiciones— y la especificación —los requisitos humanos: estudios, habilidades, personalidad—. Ojo: no se analiza a la persona, sino el puesto. Recoge seis tipos de información, que veis a la derecha.", "P1 · APT");
  T(s, "Procedimiento para establecer las obligaciones y habilidades que requiere un puesto y el tipo de persona que debe ocuparlo. Se analiza el puesto, no a la persona.", { x: 0.6, y: 1.65, w: 5.6, h: 1.1, fontSize: 15, color: INK, objectName: "a1_def" });
  [["Descripción del puesto", "Qué hace, cómo y en qué condiciones: funciones, responsabilidades, autoridad, estándares y condiciones laborales.", TEAL, TEAL_L], ["Especificación del puesto", "Requisitos humanos: estudios, conocimientos, habilidades, personalidad.", "C77A2E", AMBER_L]].forEach(([h, b, c, bg], i) => {
    const y = 2.95 + i * 1.8;
    R(s, { x: 0.6, y, w: 5.6, h: 1.6, fill: { color: bg }, objectName: `a${i + 2}_b` });
    T(s, h, { x: 0.85, y: y + 0.18, w: 5.1, h: 0.4, fontSize: 17, bold: true, color: c, objectName: `a${i + 2}_h` });
    T(s, b, { x: 0.85, y: y + 0.62, w: 5.1, h: 0.9, fontSize: 14, color: INK, objectName: `a${i + 2}_t` });
  });
  const info = [["MdWork", "Actividades laborales"], ["MdDirectionsRun", "Conductas humanas"], ["MdBuild", "Equipos y herramientas"], ["MdSpeed", "Estándares de desempeño"], ["MdBusiness", "Contexto del puesto"], ["MdPerson", "Requisitos humanos"]];
  for (let i = 0; i < 6; i++) {
    const col = i % 2, row = Math.floor(i / 2), x = 6.75 + col * 3.0, y = 1.75 + row * 1.6;
    await circleIcon(s, info[i][0], x, y, 0.85, i % 2 ? AMBER : TEAL, WHITE, `a${4 + row}_ic${i}`);
    T(s, info[i][1], { x: x + 1.0, y: y + 0.1, w: 1.85, h: 0.7, fontSize: 15, bold: true, color: NAVY, valign: "middle", objectName: `a${4 + row}_it${i}` });
  }
  T(s, "Fuente: Tema 2, apdo. 2.1.1 (Dessler, 2009)", { x: 6.75, y: 6.45, w: 5.9, h: 0.3, fontSize: 10, color: MUTED });

  // ---------- 7. PAPEL EN LA FUSIÓN ----------
  s = content("En la fusión, el APT es el mapa de las dos plantillas", "PREGUNTA 1 · PAPEL DEL APT",
    "Aplicado a la fusión, el APT sirve para comparar los puestos de ambas entidades. Detecta cuatro cosas: duplicidades, sobre todo en servicios centrales y oficinas cercanas; diferencias organizativas, porque un mismo nombre de puesto puede esconder funciones distintas y al revés; competencias aprovechables en la plantilla; y posibilidades de reorganizar, ampliar o redistribuir funciones. Por ejemplo, si los dos bancos tenían departamentos administrativos parecidos en servicios centrales, el APT compara sus tareas reales antes de decidir cómo organizar el trabajo.", "P1 · APT");
  const roles = [["MdContentCopy", "Duplicidades", "Puestos con funciones iguales o muy similares en oficinas y servicios centrales."], ["MdCompareArrows", "Diferencias organizativas", "Mismo nombre y distintas funciones, o nombres distintos para tareas equivalentes."], ["MdSchool", "Competencias", "Conocimientos y habilidades aprovechables para nuevas responsabilidades."], ["MdAccountTree", "Reorganización", "Funciones que pueden integrarse, ampliarse o redistribuirse."]];
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + i * 3.08;
    R(s, { x, y: 1.75, w: 2.85, h: 3.6, fill: { color: WHITE }, line: { color: "D5DEE7", width: 1 }, shadow: shadow(), objectName: `a${i + 1}_c` });
    await circleIcon(s, roles[i][0], x + 0.3, 2.0, 0.9, i % 2 ? AMBER : TEAL, WHITE, `a${i + 1}_i`);
    T(s, roles[i][1], { x: x + 0.3, y: 3.1, w: 2.3, h: 0.7, fontSize: 18, bold: true, color: NAVY, objectName: `a${i + 1}_h` });
    T(s, roles[i][2], { x: x + 0.3, y: 3.8, w: 2.3, h: 1.45, fontSize: 14, color: INK, objectName: `a${i + 1}_t` });
  }
  R(s, { x: 0.6, y: 5.65, w: 12.13, h: 0.95, fill: { color: LIGHT }, objectName: "a5_e" });
  T(s, [{ text: "Ejemplo: ", options: { bold: true, color: TEAL } }, { text: "dos áreas administrativas similares en los servicios centrales de Barcelona/Valencia y Madrid → el APT compara tareas y responsabilidades reales antes de decidir cómo se organiza el trabajo." }], { x: 0.85, y: 5.75, w: 11.7, h: 0.75, fontSize: 14, valign: "middle", objectName: "a5_et" });

  // ---------- 8. ÚTIL PARA AJUSTE ----------
  s = content("¿Útil para el ajuste de plantilla? Sí, pero no decide solo", "PREGUNTA 1 · AJUSTE DE PLANTILLA",
    "Sí, el APT es muy útil para el ajuste, porque da información objetiva: qué puestos son necesarios, cuáles se solapan y quién puede asumir otras funciones. Así las decisiones se basan en las necesidades reales y no solo en número de empleados o coste salarial. Pero tiene límites: el APT no dice cuántas personas deben salir. Eso requiere la planificación de RRHH y, además, cualquier medida debe respetar la ley y negociarse con la representación de los trabajadores, como de hecho ocurrió con el ERE.", "P1 · APT");
  R(s, { x: 0.6, y: 1.75, w: 2.6, h: 4.85, fill: { color: NAVY }, objectName: "a1_b" });
  T(s, "SÍ", { x: 0.6, y: 2.3, w: 2.6, h: 1.5, fontSize: 80, bold: true, color: AMBER, align: "center", objectName: "a1_y" });
  T(s, "con información objetiva, no solo con costes", { x: 0.85, y: 4.0, w: 2.1, h: 1.6, fontSize: 16, color: WHITE, align: "center", objectName: "a1_yt" });
  const col = (x, w, head, c, bg, items, step, ic) => {
    R(s, { x, y: 1.75, w, h: 4.85, fill: { color: bg }, objectName: `a${step}_bg` });
    T(s, head, { x: x + 0.3, y: 1.95, w: w - 0.6, h: 0.45, fontSize: 18, bold: true, color: c, objectName: `a${step}_h` });
    T(s, items.map((t, i) => ({ text: t, options: { bullet: { code: ic }, breakLine: i < items.length - 1, paraSpaceAfter: 8 } })), { x: x + 0.3, y: 2.55, w: w - 0.6, h: 3.9, fontSize: 15, color: INK, objectName: `a${step}_l` });
  };
  col(3.45, 4.55, "Lo que aporta", GREEN, "E3F3EB", ["Qué puestos son necesarios en la nueva estructura", "Qué puestos presentan solapamientos", "Qué personas pueden asumir otras funciones (movilidad, formación)", "Criterios objetivos y defendibles ante la plantilla y los sindicatos"], 2, "2714");
  col(8.18, 4.55, "Sus límites", "B5452F", CORAL_L, ["No determina por sí solo cuántas personas deben salir", "Debe completarse con la planificación de RRHH y las necesidades futuras", "Toda medida exige respetar la ley y negociar con la representación de los trabajadores"], 3, "26A0");

  // ---------- 9. HUB ----------
  s = content("El APT alimenta todos los procesos de RRHH de la nueva entidad", "PREGUNTA 1 · CONEXIÓN CON LA ASIGNATURA",
    "El APT no es un fin en sí mismo: es la base de los demás procesos de RRHH, como vimos en el Tema 2. En la fusión conecta con todo el temario: con el Tema 1, porque integra la gestión de personas en la nueva estrategia; con la planificación del Tema 2; con el Tema 3, para reubicar y socializar a las personas en sus nuevos puestos; con el Tema 4, para formar, recualificar y gestionar traslados y ascensos; y con el Tema 5, porque la valoración de puestos es la base para homologar los salarios de la plantilla de Bankia, algo que se pactó de forma progresiva en cinco años.", "P1 · APT");
  s.addShape(pres.shapes.OVAL, { x: 5.42, y: 2.95, w: 2.5, h: 2.5, fill: { color: NAVY }, line: { type: "none" }, objectName: "a1_hub" });
  T(s, [{ text: "APT", options: { fontSize: 34, bold: true, breakLine: true } }, { text: "descripción + especificación", options: { fontSize: 12 } }], { x: 5.55, y: 3.55, w: 2.25, h: 1.3, color: WHITE, align: "center", objectName: "a1_hubt" });
  const spokes = [[0.6, 1.65, "T1 · Estrategia", "Integrar la gestión de personas en la estrategia de la entidad fusionada"], [8.73, 1.65, "T2 · Planificación", "Base para prever necesidades y oferta interna de personal"], [0.6, 3.6, "T3 · Afectación", "Reubicación, selección interna y socialización en los nuevos puestos"], [8.73, 3.6, "T4 · Desarrollo", "Formación, recualificación, ascensos y transferencias"], [4.67, 5.65, "T5 · Retribución", "Valoración de puestos para homologar salarios (equidad interna)"]];
  spokes.forEach(([x, y, h, t], i) => {
    R(s, { x, y, w: 4.0, h: 1.2, fill: { color: i % 2 ? AMBER_L : TEAL_L }, objectName: `a${i + 2}_s` });
    T(s, h, { x: x + 0.2, y: y + 0.1, w: 3.6, h: 0.38, fontSize: 16, bold: true, color: i % 2 ? "B0661E" : TEAL, objectName: `a${i + 2}_h` });
    T(s, t, { x: x + 0.2, y: y + 0.48, w: 3.6, h: 0.68, fontSize: 13, color: INK, objectName: `a${i + 2}_t` });
  });
  [[4.6, 2.25, 0.95, 1.0, false], [7.82, 2.25, 0.91, 1.0, true], [4.6, 4.2, 0.82, 0, false], [7.92, 4.2, 0.81, 0, true], [6.67, 5.45, 0, 0.2, false]].forEach(([x, y, w, h, flip], i) =>
    s.addShape(pres.shapes.LINE, { x, y, w, h, flipH: flip, line: { color: "B8C7D6", width: 2 }, objectName: `a1_ln${i}` }));

  // ---------- 10. SECCIÓN 2 ----------
  section("02", "Pasos del APT aplicados a la fusión", "Seis pasos (Tema 2) y los métodos para recoger la información en una plantilla de más de 51.000 personas",
    "Segunda pregunta: ¿qué pasos seguiríamos para hacer el APT en la fusión?", "P2 · Pasos APT");

  // ---------- 11. SEIS PASOS ----------
  s = content("Seis pasos del APT aplicados a CaixaBank y Bankia", "PREGUNTA 2 · PROCESO",
    "El Tema 2 propone seis pasos. Uno: decidir para qué se usará la información; aquí, mapear puestos, detectar duplicidades y apoyar la reorganización. Dos: revisar organigramas y estructura de ambos bancos. Tres: elegir puestos representativos, porque no podemos estudiar 51.000 puestos uno a uno; empezaríamos por atención al cliente, gestores comerciales, administrativos y servicios centrales. Cuatro: analizar los puestos con entrevistas, cuestionarios, observación y diarios. Cinco: verificar la información con empleados y supervisores, la triangulación, porque en una fusión lo formal y lo real pueden no coincidir. Y seis: redactar descripciones y especificaciones que permitan comparar y fijar los puestos de la nueva estructura.", "P2 · Pasos APT");
  const steps = [["Determinar el uso", "Mapear puestos, detectar duplicidades y apoyar la reorganización"], ["Revisar información básica", "Organigramas y estructura de ambos bancos: relaciones jerárquicas y solapamientos"], ["Seleccionar puestos representativos", "Atención al cliente, gestores comerciales, administrativos y servicios centrales"], ["Analizar los puestos", "Actividades, conductas, condiciones y habilidades con métodos combinados"], ["Verificar con empleado y supervisor", "Triangulación: lo formal frente a lo que realmente se hace"], ["Descripción y especificación", "Fichas comparables para fijar los puestos de la nueva estructura"]];
  steps.forEach(([h, t], i) => {
    const row = Math.floor(i / 3), c = i % 3, x = 0.6 + c * 4.12, y = 1.7 + row * 2.55;
    R(s, { x, y, w: 3.75, h: 2.25, fill: { color: row ? LIGHT : WHITE }, line: { color: "D5DEE7", width: 1 }, shadow: shadow(), objectName: `a${i + 1}_c` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: y + 0.25, w: 0.7, h: 0.7, fill: { color: i < 3 ? TEAL : AMBER }, line: { type: "none" }, objectName: `a${i + 1}_o` });
    T(s, String(i + 1), { x: x + 0.25, y: y + 0.25, w: 0.7, h: 0.7, fontSize: 24, bold: true, color: WHITE, align: "center", valign: "middle", objectName: `a${i + 1}_n` });
    T(s, h, { x: x + 1.1, y: y + 0.22, w: 2.5, h: 0.8, fontSize: 16, bold: true, color: NAVY, valign: "middle", objectName: `a${i + 1}_h` });
    T(s, t, { x: x + 0.25, y: y + 1.1, w: 3.3, h: 1.05, fontSize: 14, color: INK, objectName: `a${i + 1}_t` });
    if (c < 2) s.addShape(pres.shapes.RIGHT_TRIANGLE ? pres.shapes.CHEVRON : pres.shapes.CHEVRON, { x: x + 3.83, y: y + 0.95, w: 0.22, h: 0.35, fill: { color: "B8C7D6" }, line: { type: "none" }, objectName: `a${i + 1}_ar` });
  });

  // ---------- 12. MÉTODOS ----------
  s = content("Cómo recoger la información en una plantilla de 51.000 personas", "PREGUNTA 2 · MÉTODOS DE RECOGIDA",
    "Para el paso cuatro combinamos los cuatro métodos del tema. Las entrevistas, individuales o en grupo, con trabajadores y supervisores dan profundidad, aunque pueden distorsionarse: en una fusión hay miedo a perder el puesto y la gente puede exagerar sus funciones. Los cuestionarios, enviados por la intranet a toda la red de oficinas, son rápidos y permiten comparar ambas plantillas. La observación funciona en tareas rutinarias de oficina. Y los diarios de actividad muestran el tiempo real dedicado a cada tarea. La clave es la triangulación entre trabajador, supervisor y RRHH.", "P2 · Pasos APT");
  const meth = [["MdQuestionAnswer", "Entrevista", "Trabajadores y supervisores, individual o en grupo", "+ Profundidad y rapidez", "– Distorsión: miedo a perder el puesto"], ["MdAssignment", "Cuestionario", "Vía intranet a toda la red: datos comparables", "+ Rápido y eficiente", "– Diseño y validación costosos"], ["MdVisibility", "Observación", "Tareas rutinarias de oficina durante un ciclo completo", "+ Lo que realmente se hace", "– Reactividad del observado"], ["MdEditNote", "Diario de actividades", "Tareas y tiempo dedicado a cada una", "+ Cronológico y detallado", "– Depende de la constancia"]];
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + (i % 2) * 6.15, y = 1.7 + Math.floor(i / 2) * 2.15;
    R(s, { x, y, w: 5.95, h: 1.95, fill: { color: LIGHT }, objectName: `a${i + 1}_c` });
    await circleIcon(s, meth[i][0], x + 0.25, y + 0.3, 0.85, i % 2 ? AMBER : TEAL, WHITE, `a${i + 1}_i`);
    T(s, meth[i][1], { x: x + 1.3, y: y + 0.2, w: 4.4, h: 0.4, fontSize: 17, bold: true, color: NAVY, objectName: `a${i + 1}_h` });
    T(s, meth[i][2], { x: x + 1.3, y: y + 0.62, w: 4.45, h: 0.6, fontSize: 14, color: INK, objectName: `a${i + 1}_t` });
    T(s, [{ text: meth[i][3], options: { color: GREEN, bold: true } }, { text: "    " }, { text: meth[i][4], options: { color: "B5452F", bold: true } }], { x: x + 1.3, y: y + 1.3, w: 4.5, h: 0.5, fontSize: 13, objectName: `a${i + 1}_p` });
  }
  R(s, { x: 0.6, y: 6.05, w: 12.13, h: 0.65, fill: { color: NAVY }, objectName: "a5_k" });
  T(s, [{ text: "Clave: ", options: { bold: true, color: AMBER } }, { text: "triangulación trabajador – supervisor – RRHH y explicar el objetivo del análisis para reducir la desconfianza.", options: { color: WHITE } }], { x: 0.85, y: 6.1, w: 11.7, h: 0.55, fontSize: 14, valign: "middle", objectName: "a5_kt" });

  // ---------- 13. FICHA COMPARADA ----------
  s = content("Resultado: fichas de puesto comparables", "PREGUNTA 2 · PASO 6 · EJEMPLO ILUSTRATIVO",
    "Este es el resultado del paso seis, con un ejemplo ilustrativo, no con datos reales. Comparamos el gestor comercial de oficina de cada banco en los apartados de la descripción: funciones, autoridad, estándares y requisitos. A partir de ahí se define un único puesto unificado para la nueva entidad, y se ve quién cumple ya la especificación y quién necesita formación, lo que conecta con el Tema 4.", "P2 · Pasos APT");
  const hdr = (t, c) => ({ text: t, options: { bold: true, color: WHITE, fill: { color: c }, fontSize: 14 } });
  const rows = [
    [hdr("Apartado", NAVY), hdr("Gestor/a comercial · CaixaBank", TEAL), hdr("Gestor/a comercial · Bankia", "B0661E"), hdr("Puesto unificado (nueva entidad)", NAVY)],
    ["Funciones", "Cartera de particulares, venta de productos, asesoramiento", "Cartera de particulares y pymes, operaciones de caja", "Cartera mixta y asesoramiento; caja en canal digital"],
    ["Autoridad", "Aprobación de riesgos hasta un límite", "Propuesta de riesgos al director", "Límite único de aprobación para todos"],
    ["Estándares", "Objetivos comerciales trimestrales", "Objetivos anuales por oficina", "Objetivos comunes y medibles"],
    ["Especificación", "Grado + certificación MiFID II", "Grado; MiFID II en curso", "Grado + MiFID II + competencias digitales"],
  ];
  s.addTable(rows, { x: 0.6, y: 1.7, w: 12.13, colW: [1.9, 3.4, 3.4, 3.43], fontSize: 13, color: INK, border: { type: "solid", color: "D5DEE7", pt: 1 }, fill: { color: WHITE }, rowH: 0.72, valign: "middle", margin: 0.08, objectName: "a1_tab" });
  R(s, { x: 0.6, y: 5.6, w: 12.13, h: 0.95, fill: { color: AMBER_L }, objectName: "a2_n" });
  T(s, [{ text: "Uso: ", options: { bold: true, color: "B0661E" } }, { text: "detectar equivalencias y brechas → quién cumple ya la especificación y quién necesita formación (T4). Datos de ejemplo, no reales." }], { x: 0.85, y: 5.7, w: 11.7, h: 0.75, fontSize: 14, valign: "middle", objectName: "a2_nt" });

  // ---------- 14. SECCIÓN 3 ----------
  section("03", "Una nueva planificación de RRHH", "¿Es recomendable? Sí. Siete pasos, de la estrategia al control",
    "Tercera pregunta: ¿debería la nueva entidad hacer una nueva planificación de recursos humanos? Nuestra respuesta es sí.", "P3 · Planificación");

  // ---------- 15. POR QUÉ ----------
  s = content("¿Nueva planificación? Sí: la fusión lo cambia todo", "PREGUNTA 3 · JUSTIFICACIÓN",
    "Es muy recomendable porque la fusión cambia la estrategia, la estructura y las necesidades de personal. El Tema 2 lo dice claramente: la planificación surge de los planes estratégicos, como reducir costes, que afectan a los tipos de puestos que la empresa necesita. Aquí la estrategia es aprovechar sinergias de 770 millones; la estructura cambia por oficinas próximas y servicios centrales duplicados; y las necesidades cambian por la digitalización. Mantener dos planificaciones separadas no tendría sentido.", "P3 · Planificación");
  const why = [["MdFlag", "Estrategia nueva", "Reducir costes y aprovechar sinergias (≈770 M€/año) e integrar dos entidades."], ["MdAccountTree", "Estructura nueva", "Oficinas próximas entre sí y servicios centrales duplicados."], ["MdInsights", "Necesidades nuevas", "Digitalización, menos uso de la oficina y perfiles tecnológicos."]];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.12;
    R(s, { x, y: 1.75, w: 3.85, h: 2.85, fill: { color: WHITE }, line: { color: "D5DEE7", width: 1 }, shadow: shadow(), objectName: `a${i + 1}_c` });
    await circleIcon(s, why[i][0], x + 0.3, 2.0, 0.85, i === 1 ? AMBER : TEAL, WHITE, `a${i + 1}_i`);
    T(s, why[i][1], { x: x + 0.3, y: 3.0, w: 3.3, h: 0.45, fontSize: 18, bold: true, color: NAVY, objectName: `a${i + 1}_h` });
    T(s, why[i][2], { x: x + 0.3, y: 3.5, w: 3.3, h: 1.0, fontSize: 14, color: INK, objectName: `a${i + 1}_t` });
  }
  R(s, { x: 0.6, y: 4.95, w: 12.13, h: 1.65, fill: { color: NAVY }, objectName: "a4_q" });
  T(s, "“La planificación surge de los planes estratégicos de la empresa, tales como entrar en nuevos negocios, construir plantas o reducir costes, que afectan a los tipos de puestos que la empresa necesita.”", { x: 0.95, y: 5.1, w: 11.4, h: 1.0, fontSize: 16, italic: true, color: WHITE, objectName: "a4_qt" });
  T(s, "Tema 2, apdo. 2.2.1", { x: 0.95, y: 6.12, w: 6, h: 0.35, fontSize: 12, color: AMBER, objectName: "a4_qs" });

  // ---------- 16. SIETE PASOS ----------
  s = content("Siete pasos para planificar la plantilla integrada", "PREGUNTA 3 · PROCESO",
    "Los pasos serían siete. Primero, partir de los objetivos de la nueva entidad. Segundo, pronosticar cuántas personas y con qué competencias necesitará, usando las herramientas del tema: análisis de tendencias, de razón, diagrama de dispersión o pronósticos informatizados, sin olvidar el juicio de la dirección. Tercero, estimar la oferta interna con inventarios de cualificaciones y gráficos de reemplazo. Cuarto, la oferta externa, solo si dentro no hay los perfiles necesarios, por ejemplo en tecnología o ciberseguridad. Quinto, comparar demanda y oferta. Sexto, diseñar las medidas. Y séptimo, controlar y evaluar, porque la planificación es un proceso vivo.", "P3 · Planificación");
  const p7 = [["Objetivos y estrategia", "Eficiencia, reducción de costes y sinergias de la integración"], ["Pronóstico de necesidades", "Actividad, productividad, tecnología, red de oficinas, rotación (tendencias, razón, dispersión)"], ["Oferta interna", "Inventarios de cualificaciones, sistemas de información y gráficos de reemplazo"], ["Oferta externa", "Solo si faltan perfiles: tecnología, ciberseguridad, servicios digitales"], ["Comparar demanda y oferta", "Cantidad y cualificación: exceso, déficit o equilibrio"], ["Diseñar y aplicar medidas", "Movilidad, formación, salidas voluntarias negociadas, contratación"], ["Controlar y evaluar", "Revisión periódica: la planificación no es un proceso cerrado"]];
  p7.forEach(([h, t], i) => {
    const c = i < 4 ? 0 : 1, r = i < 4 ? i : i - 4, x = 0.6 + c * 6.15, y = 1.7 + r * 1.27;
    R(s, { x, y, w: 5.95, h: 1.1, fill: { color: c ? AMBER_L : TEAL_L }, objectName: `a${i + 1}_c` });
    T(s, String(i + 1), { x: x + 0.2, y: y + 0.15, w: 0.65, h: 0.8, fontSize: 32, bold: true, color: c ? "B0661E" : TEAL, align: "center", valign: "middle", objectName: `a${i + 1}_n` });
    T(s, h, { x: x + 1.0, y: y + 0.1, w: 4.8, h: 0.38, fontSize: 16, bold: true, color: NAVY, objectName: `a${i + 1}_h` });
    T(s, t, { x: x + 1.0, y: y + 0.48, w: 4.8, h: 0.58, fontSize: 13, color: INK, objectName: `a${i + 1}_t` });
  });
  R(s, { x: 6.75, y: 5.51, w: 5.95, h: 1.1, fill: { color: NAVY }, objectName: "a8_k" });
  T(s, "APT (P1–P2) → datos de puestos para los pasos 2, 3 y 5", { x: 7.0, y: 5.6, w: 5.5, h: 0.9, fontSize: 15, bold: true, color: WHITE, valign: "middle", objectName: "a8_kt" });

  // ---------- 17. ESCENARIOS ----------
  s = content("Comparar demanda y oferta: tres escenarios a la vez", "PREGUNTA 3 · PASOS 5 Y 6",
    "Al comparar demanda y oferta pueden darse tres situaciones, y en una fusión se dan las tres a la vez. Exceso en funciones administrativas y oficinas solapadas: movilidad funcional y geográfica, reasignación, salidas voluntarias incentivadas o prejubilaciones, siempre negociadas. Déficit en perfiles digitales y especializados: formación, promoción interna y, si hace falta, contratación externa. Y equilibrio en otras áreas, donde basta con reubicar y estandarizar los puestos.", "P3 · Planificación");
  const sc = [["MdTrendingUp", "Exceso de plantilla", "Funciones administrativas y oficinas solapadas", ["Movilidad funcional y geográfica", "Reasignación de tareas", "Salidas voluntarias incentivadas", "Prejubilaciones (negociadas)"], CORAL, CORAL_L, "B5452F"], ["MdTrendingDown", "Déficit de plantilla", "Perfiles digitales, tecnológicos y especializados", ["Formación y recualificación (T4)", "Promoción interna", "Contratación externa si es necesaria (T3)"], AMBER, AMBER_L, "B0661E"], ["MdBalance", "Equilibrio", "Áreas donde la plantilla se ajusta a lo previsto", ["Reubicar y estandarizar puestos", "Socializar a los equipos integrados (T3)", "Homologar condiciones (T5)"], GREEN, "E3F3EB", GREEN]];
  for (let i = 0; i < 3; i++) {
    const [ic, h, sub, items, c, bg, tc] = sc[i], x = 0.6 + i * 4.12;
    R(s, { x, y: 1.7, w: 3.85, h: 4.9, fill: { color: bg }, objectName: `a${i + 1}_c` });
    await circleIcon(s, ic, x + 0.3, 1.95, 0.85, c, WHITE, `a${i + 1}_i`);
    T(s, h, { x: x + 1.3, y: 2.05, w: 2.4, h: 0.65, fontSize: 18, bold: true, color: tc, valign: "middle", objectName: `a${i + 1}_h` });
    T(s, sub, { x: x + 0.3, y: 3.0, w: 3.3, h: 0.7, fontSize: 14, italic: true, color: INK, objectName: `a${i + 1}_s` });
    T(s, items.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < items.length - 1, paraSpaceAfter: 6 } })), { x: x + 0.3, y: 3.8, w: 3.35, h: 2.65, fontSize: 14, color: INK, objectName: `a${i + 1}_l` });
  }

  // ---------- 18. AJUSTE REAL ----------
  s = content("El ajuste real: la negociación redujo las salidas un 22 %", "PREGUNTA 3 · LO QUE OCURRIÓ",
    "¿Qué pasó en la realidad? En abril de 2021 CaixaBank propuso 8.291 salidas. Tras negociar con los sindicatos, el 1 de julio se firmaron 6.452 salidas, un 22 % menos, con carácter voluntario y un coste estimado de unos 1.900 millones de euros. Además se pactó homologar progresivamente las condiciones de la plantilla de Bankia en cinco años. La lección: una planificación apoyada en el APT da una base objetiva para negociar, y la negociación forma parte del paso seis.", "P3 · Planificación");
  s.addChart(pres.charts.BAR, [{ name: "Salidas", labels: ["Propuesta inicial (abr 2021)", "Acuerdo final (1 jul 2021)"], values: [8291, 6452] }], {
    x: 0.6, y: 1.65, w: 6.6, h: 4.95, barDir: "col", chartColors: ["8DA9C4", TEAL], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 16, dataLabelFontBold: true, dataLabelColor: INK, dataLabelFormatCode: "#,##0",
    showLegend: false, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, catAxisLabelColor: MUTED, catAxisLabelFontSize: 13, catAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", barGapWidthPct: 60,
    showTitle: true, title: "Salidas de empleados planteadas en el ERE", titleFontSize: 15, titleColor: NAVY, titleFontFace: "+mn-lt", valAxisMinVal: 0, valAxisMaxVal: 9500, objectName: "a1_ch" });
  const facts = [["–22 %", "salidas tras la negociación sindical"], ["≈1.900 M€", "coste estimado del ajuste"], ["5 años", "homologación progresiva de condiciones de la plantilla de Bankia"]];
  facts.forEach(([v, l], i) => {
    const y = 1.75 + i * 1.3;
    R(s, { x: 7.55, y, w: 5.18, h: 1.12, fill: { color: LIGHT }, objectName: `a${i + 2}_f` });
    T(s, v, { x: 7.8, y: y + 0.13, w: 1.95, h: 0.85, fontSize: 28, bold: true, color: i ? TEAL : AMBER, valign: "middle", objectName: `a${i + 2}_v` });
    T(s, l, { x: 9.8, y: y + 0.13, w: 2.8, h: 0.85, fontSize: 14, color: INK, valign: "middle", objectName: `a${i + 2}_l` });
  });
  R(s, { x: 7.55, y: 5.67, w: 5.18, h: 0.93, fill: { color: NAVY }, objectName: "a5_k" });
  T(s, "APT + planificación = base objetiva para negociar el ajuste", { x: 7.8, y: 5.72, w: 4.8, h: 0.83, fontSize: 15, bold: true, color: WHITE, valign: "middle", objectName: "a5_kt" });
  T(s, "Fuentes: Bolsamanía, Forbes, Ara (jul. 2021).", { x: 0.6, y: 6.65, w: 8, h: 0.3, fontSize: 10, color: MUTED });

  // ---------- 19. SECCIÓN 4 ----------
  section("04", "Las planificaciones previas a la fusión", "Útiles como punto de partida, insuficientes como plan para la nueva entidad",
    "Cuarta y última pregunta: ¿en qué medida sirven las planificaciones de RRHH que cada banco tenía antes de la fusión?", "P4 · Planes previos");

  // ---------- 20. UTILIDAD VS LÍMITES ----------
  s = content("Planes previos: punto de partida, no plan definitivo", "PREGUNTA 4 · UTILIDAD Y LÍMITES",
    "Las planificaciones previas son útiles: dan una foto de partida de cada plantilla, ayudan a detectar necesidades y excedentes al compararlas, permiten identificar personas que pueden asumir nuevas funciones gracias a los inventarios de cualificaciones, y sirven para contrastar previsiones anteriores con los nuevos objetivos. Pero tienen límites: se hicieron con la estructura y la estrategia de cada banco por separado, no contemplan las duplicidades que aparecen al integrarse y pueden quedar desfasadas por la digitalización y el menor uso de las oficinas.", "P4 · Planes previos");
  const useL = [["MdPhotoCamera", "Foto de partida", "Nº de empleados, cualificaciones, experiencia y puestos de cada plantilla"], ["MdSearch", "Necesidades y excedentes", "Al compararlas aparecen áreas similares y posibles duplicidades"], ["MdSchool", "Talento reubicable", "Inventarios de cualificaciones: movilidad y formación antes que salidas"], ["MdCompareArrows", "Contraste con objetivos", "¿Siguen siendo válidas las previsiones con la nueva estrategia?"]];
  const limL = [["MdCallSplit", "Hechas por separado", "Cada banco con su estructura, procesos y estrategia"], ["MdContentCopy", "Sin duplicidades", "No previeron los solapamientos que surgen al integrarse"], ["MdUpdate", "Desfase", "Digitalización y menor uso de la oficina cambian las necesidades"]];
  T(s, "ÚTILES PARA…", { x: 0.6, y: 1.6, w: 6, h: 0.4, fontSize: 15, bold: true, color: GREEN, objectName: "a1_h1" });
  for (let i = 0; i < 4; i++) {
    const y = 2.05 + i * 1.15;
    await circleIcon(s, useL[i][0], 0.6, y, 0.8, GREEN, WHITE, `a${i + 1}_ui`);
    T(s, useL[i][1], { x: 1.6, y: y - 0.03, w: 5.0, h: 0.36, fontSize: 16, bold: true, color: NAVY, objectName: `a${i + 1}_uh` });
    T(s, useL[i][2], { x: 1.6, y: y + 0.33, w: 5.0, h: 0.6, fontSize: 13, color: INK, objectName: `a${i + 1}_ut` });
  }
  R(s, { x: 7.05, y: 1.6, w: 5.68, h: 5.0, fill: { color: CORAL_L }, objectName: "a5_bg" });
  T(s, "PERO INSUFICIENTES PORQUE…", { x: 7.35, y: 1.8, w: 5.2, h: 0.4, fontSize: 15, bold: true, color: "B5452F", objectName: "a5_h2" });
  for (let i = 0; i < 3; i++) {
    const y = 2.4 + i * 1.35;
    await circleIcon(s, limL[i][0], 7.35, y, 0.8, CORAL, WHITE, `a${i + 6}_li`);
    T(s, limL[i][1], { x: 8.35, y: y - 0.03, w: 4.2, h: 0.36, fontSize: 16, bold: true, color: NAVY, objectName: `a${i + 6}_lh` });
    T(s, limL[i][2], { x: 8.35, y: y + 0.33, w: 4.2, h: 0.7, fontSize: 13, color: INK, objectName: `a${i + 6}_lt` });
  }

  // ---------- 21. 1+1≠2 ----------
  s = content("1 + 1 ≠ 2: de dos planes a un plan integrado", "PREGUNTA 4 · CONCLUSIÓN",
    "La idea clave de esta pregunta es que uno más uno no es igual a dos. Las necesidades de la nueva entidad no son la suma de las que tenían CaixaBank y Bankia por separado. Los planes previos son un insumo, junto con el APT conjunto y la nueva estrategia, para construir una planificación integrada a corto y largo plazo. Aprovechar lo que ya existe reduce la incertidumbre inicial y acelera la transición.", "P4 · Planes previos");
  const ins = [["Plan RRHH CaixaBank", TEAL], ["Plan RRHH Bankia", "B0661E"], ["APT conjunto", NAVY], ["Nueva estrategia", MUTED]];
  ins.forEach(([t, c], i) => {
    const y = 1.75 + i * 1.18;
    R(s, { x: 0.6, y, w: 3.6, h: 0.95, fill: { color: c }, objectName: `a${i + 1}_in` });
    T(s, t, { x: 0.6, y, w: 3.6, h: 0.95, fontSize: 17, bold: true, color: WHITE, align: "center", valign: "middle", objectName: `a${i + 1}_int` });
  });
  s.addShape(pres.shapes.RIGHT_ARROW, { x: 4.5, y: 3.4, w: 1.3, h: 1.0, fill: { color: AMBER }, line: { type: "none" }, objectName: "a5_ar" });
  s.addShape(pres.shapes.OVAL, { x: 6.1, y: 1.95, w: 3.9, h: 3.9, fill: { color: NAVY }, line: { type: "none" }, objectName: "a6_o" });
  T(s, [{ text: "Plan integrado", options: { fontSize: 26, bold: true, breakLine: true } }, { text: "de la nueva entidad", options: { fontSize: 16, breakLine: true } }, { text: "corto y largo plazo", options: { fontSize: 14, color: AMBER } }], { x: 6.4, y: 3.05, w: 3.3, h: 1.7, color: WHITE, align: "center", valign: "middle", objectName: "a6_ot" });
  T(s, "1 + 1 ≠ 2", { x: 10.25, y: 2.45, w: 2.5, h: 1.0, fontSize: 40, bold: true, color: AMBER, objectName: "a7_eq" });
  T(s, "Las necesidades de la entidad fusionada no son la suma de las que tenían ambos bancos por separado.", { x: 10.25, y: 3.5, w: 2.5, h: 2.2, fontSize: 15, color: INK, objectName: "a7_eqt" });

  // ---------- 22. CONCLUSIONES ----------
  pres.addSection({ title: "Cierre" });
  s = dark("Para cerrar, tres ideas. Una: el APT es el mapa de la fusión; sin él no se sabe qué puestos sobran, cuáles faltan ni quién puede reubicarse. Dos: la nueva planificación conecta ese mapa con la estrategia y convierte las sinergias en decisiones sobre personas, desde la movilidad y la formación hasta salidas negociadas. Y tres: los planes previos sirven como punto de partida, pero la entidad fusionada necesita una visión global. Si de verdad las personas son el principal activo, el ajuste debe basarse en análisis objetivo y negociación. Muchas gracias.", "Conclusiones", "Cierre");
  T(s, "CONCLUSIONES", { x: 0.8, y: 0.7, w: 8, h: 0.4, fontSize: 14, bold: true, color: AMBER, charSpacing: 2, objectName: "a1_k" });
  T(s, "Personas como activo, también en una fusión", { x: 0.8, y: 1.15, w: 11.8, h: 0.9, fontSize: 36, bold: true, color: WHITE, objectName: "a1_t" });
  const concl = [["01", "El APT es el mapa", "Sin él no se sabe qué puestos sobran, cuáles faltan ni quién puede reubicarse."], ["02", "La planificación decide", "Conecta el mapa con la estrategia: movilidad, formación y salidas negociadas."], ["03", "Punto de partida", "Los planes anteriores informan, pero la nueva entidad necesita una visión global."]];
  concl.forEach(([nn, h, t], i) => {
    const x = 0.8 + i * 4.0;
    R(s, { x, y: 2.5, w: 3.7, h: 3.3, fill: { color: "16365E" }, objectName: `a${i + 2}_c` });
    T(s, nn, { x: x + 0.3, y: 2.7, w: 2, h: 0.8, fontSize: 40, bold: true, color: AMBER, objectName: `a${i + 2}_n` });
    T(s, h, { x: x + 0.3, y: 3.55, w: 3.1, h: 0.5, fontSize: 20, bold: true, color: WHITE, objectName: `a${i + 2}_h` });
    T(s, t, { x: x + 0.3, y: 4.15, w: 3.1, h: 1.5, fontSize: 15, color: "BFD0E0", objectName: `a${i + 2}_t` });
  });
  T(s, "Si las personas son el principal activo, el ajuste debe apoyarse en análisis objetivo y negociación.", { x: 0.8, y: 6.15, w: 11.8, h: 0.5, fontSize: 17, italic: true, color: AMBER, objectName: "a5_f" });

  // ---------- 23. FUENTES ----------
  s = content("Fuentes", "BIBLIOGRAFÍA Y PRENSA",
    "Estas son las fuentes: la bibliografía básica de la asignatura, los temas 1 a 5, el artículo de Expansión del caso y otras fuentes de prensa y oficiales que hemos usado para los datos de la fusión y del ERE.", "Cierre");
  T(s, "Bibliografía académica", { x: 0.6, y: 1.65, w: 5.8, h: 0.4, fontSize: 16, bold: true, color: TEAL });
  T(s, ["Dessler, G. (2009). Administración de recursos humanos (11ª ed.). Pearson.", "Gómez-Mejía, L. R., Balkin, D. B. y Cardy, R. L. (2008). Dirección y gestión de recursos humanos (5ª ed.). Prentice Hall.", "Dolan, S. L., Valle, R., Jackson, S. E. y Schuler, R. S. (2007). La gestión de recursos humanos (3ª ed.). McGraw-Hill.", "Universidad de Cádiz. Dirección de RRHH II, Temas 1–5 (apuntes de la asignatura)."].map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1, paraSpaceAfter: 8 } })), { x: 0.6, y: 2.1, w: 5.8, h: 4.5, fontSize: 14 });
  T(s, "Prensa y fuentes oficiales", { x: 6.9, y: 1.65, w: 5.8, h: 0.4, fontSize: 16, bold: true, color: TEAL });
  T(s, ["Expansión (19/09/2020). Artículo facilitado en el enunciado del caso.", "CaixaBank (18/09/2020). Comunicado de anuncio de la fusión con Bankia.", "Banco de España (2021). Informe de Estabilidad Financiera, primavera 2021.", "Bolsamanía (2020–2021). Aprobación de la fusión y firma del ERE (6.452 empleados).", "Forbes España y Ara (jul. 2021). Condiciones del ERE de CaixaBank.", "Consumer Eroski (2020). Fusión de CaixaBank y Bankia: lo que debe saber el cliente."].map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1, paraSpaceAfter: 8 } })), { x: 6.9, y: 2.1, w: 5.8, h: 4.5, fontSize: 14 });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  fs.writeFileSync(OUT.replace(/\.pptx$/, "_notes.json"), JSON.stringify(NOTES, null, 1));
  console.log("ok", n);
})().catch(e => { console.error(e); process.exit(1); });
