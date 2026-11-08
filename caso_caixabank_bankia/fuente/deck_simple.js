const pptxgen = require("pptxgenjs");
const fs = require("fs");
const OUT = process.argv[2];
const ROSA = "D6457A", OSCURO = "7B2D4F", FONDO = "FDEEF3", TXT = "333333";
const NOTES = [];
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Caso práctico 1 - Fusión CaixaBank y Bankia";

let n = 0;
function slide(title, notes) {
  const s = pres.addSlide();
  s.background = { color: FONDO };
  if (title) {
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 13.333, h: 0.25, fill: { color: ROSA }, line: { type: "none" } });
    s.addText(title, { x: 0.7, y: 0.5, w: 12, h: 0.9, fontSize: 34, bold: true, color: OSCURO, fontFace: "Calibri", isTextBox: true });
    s.addText(String(n + 1), { x: 12.4, y: 6.95, w: 0.6, h: 0.35, fontSize: 12, color: "999999", align: "right", isTextBox: true });
  }
  s.addNotes(notes); n++; NOTES.push({ n, title: title || "Portada", notes });
  return s;
}
function bullets(s, items, o = {}) {
  s.addText(items.map((t, i) => {
    const sub = t.startsWith("-");
    return { text: sub ? t.slice(1).trim() : t, options: { bullet: sub ? { indent: 20 } : true, indentLevel: sub ? 1 : 0, fontSize: o.fontSize || (sub ? 21 : 24), breakLine: i < items.length - 1, paraSpaceAfter: 8 } };
  }), Object.assign({ x: 0.7, y: 1.6, w: 12, h: 5.2, color: TXT, fontFace: "Calibri", valign: "top", isTextBox: true }, o));
}

// 1 portada
let s = slide(null, "Hola, nuestro trabajo es el caso práctico 1 sobre la fusión de CaixaBank y Bankia. Vamos a responder a las cuatro preguntas usando sobre todo lo del tema 2: el análisis de puestos y la planificación de recursos humanos.");
s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.5, h: 7.5, fill: { color: ROSA }, line: { type: "none" } });
s.addText("Caso práctico 1:\nFusión CaixaBank y Bankia", { x: 1.2, y: 1.8, w: 11, h: 2.0, fontSize: 44, bold: true, color: OSCURO, fontFace: "Calibri", isTextBox: true });
s.addText("Dirección de Recursos Humanos II · 4º GADE · Universidad de Cádiz", { x: 1.2, y: 3.9, w: 11, h: 0.5, fontSize: 20, color: TXT, isTextBox: true });
s.addText("Integrantes:", { x: 1.2, y: 5.2, w: 6, h: 0.5, fontSize: 18, color: TXT, isTextBox: true });

// 2 intro
s = slide("Introducción", "En septiembre de 2020 se aprueba la fusión: CaixaBank absorbe a Bankia y sale el banco más grande de España, con más de 51.000 trabajadores. Querían ahorrar unos 770 millones al año. Lo curioso es que CaixaBank dice que las personas son su principal activo, pero gran parte del ahorro viene de quitar puestos repetidos.");
bullets(s, ["Septiembre 2020: se aprueba la fusión (CaixaBank absorbe a Bankia)", "Sale el banco más grande de España", "-Más de 51.000 trabajadores (≈35.600 CaixaBank + ≈16.000 Bankia)", "-Más de 6.000 oficinas, muchas muy cerca unas de otras", "Objetivo: ahorrar unos 770 millones al año", "CaixaBank dice que las personas son su “principal activo”… pero habrá ajuste de plantilla"]);

// 3 APT
s = slide("Pregunta 1: ¿Qué es el APT?", "El APT es el procedimiento para saber qué obligaciones y habilidades tiene un puesto y qué tipo de persona hace falta. Se analiza el puesto, no a la persona. Sirve para hacer la descripción del puesto, lo que se hace, y la especificación, lo que tiene que tener la persona.");
bullets(s, ["Procedimiento para saber las obligaciones y habilidades de un puesto y qué persona necesita (Tema 2)", "Se analiza el puesto, NO a la persona", "Se obtienen dos cosas:", "-Descripción del puesto: qué se hace, cómo y en qué condiciones", "-Especificación del puesto: requisitos de la persona (estudios, habilidades…)", "Métodos: entrevistas, cuestionarios, observación y diarios"]);

// 4 papel + ajuste
s = slide("Pregunta 1: el APT en la fusión", "En la fusión el APT sería como un mapa de las dos plantillas: sirve para ver duplicidades, puestos que se llaman igual pero no hacen lo mismo, qué sabe hacer la gente y cómo reorganizar. Para el ajuste de plantilla sí sirve, porque da información objetiva, pero no dice por sí solo cuántos tienen que salir: eso es de la planificación y hay que negociarlo con los sindicatos.");
bullets(s, ["Es como un “mapa” de las dos plantillas. Permite ver:", "-Duplicidades (puestos que hacen lo mismo)", "-Puestos con el mismo nombre pero distintas funciones", "-Competencias de los trabajadores", "-Cómo reorganizar las funciones", "¿Sirve para el ajuste de plantilla? SÍ, da información objetiva", "Pero no decide solo cuántos salen → hace falta planificación y negociar con los sindicatos"], { w: 7.4 });
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.6, y: 2.0, w: 4.1, h: 3.2, rectRadius: 0.2, fill: { color: "FFFFFF" }, line: { color: ROSA, width: 2 } });
s.addText("También es la base de los demás temas: planificación, selección, formación y sueldos (para igualar sueldos hay que saber qué puestos son equivalentes)", { x: 8.85, y: 2.2, w: 3.6, h: 2.8, fontSize: 17, italic: true, color: OSCURO, isTextBox: true, valign: "middle" });

// 5 pasos APT
s = slide("Pregunta 2: pasos del APT", "Son los seis pasos del tema 2 aplicados a la fusión. Destacamos el tercero, porque no se pueden analizar 51.000 puestos uno a uno, y el quinto, porque en una fusión la gente puede exagerar lo que hace por miedo a perder el puesto, así que hay que comprobarlo con el jefe.");
s.addText([
  "Decidir para qué se usa la información → ver duplicidades y reorganizar",
  "Revisar organigramas de los dos bancos",
  "Elegir puestos representativos (no se pueden analizar 51.000 puestos)",
  "Analizar los puestos (tareas, responsabilidades, condiciones, habilidades)",
  "Comprobarlo con el trabajador y su supervisor",
  "Hacer la descripción y la especificación de cada puesto",
].map((t, i, a) => ({ text: t, options: { bullet: { type: "number" }, breakLine: i < a.length - 1, paraSpaceAfter: 14 } })), { x: 0.7, y: 1.6, w: 12, h: 5.0, fontSize: 24, color: TXT, valign: "top", isTextBox: true });

// 6 métodos
s = slide("Pregunta 2: métodos para recoger la información", "Para analizar los puestos combinaríamos los cuatro métodos. Nos parece que los cuestionarios por la intranet son lo más práctico con tanta gente, y que con las entrevistas hay que tener cuidado porque la gente puede exagerar.");
const H = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: ROSA } } });
s.addTable([[H("Método"), H("Ventaja"), H("Inconveniente")],
  ["Entrevistas", "Dan mucho detalle", "La gente puede exagerar (miedo a perder el puesto)"],
  ["Cuestionarios (intranet)", "Rápidos, llegan a todos", "Cuesta prepararlos bien"],
  ["Observación", "Se ve lo que se hace de verdad", "La gente cambia si la observan"],
  ["Diarios", "Muy detallados", "Depende de que los rellenen bien"]],
  { x: 0.7, y: 1.7, w: 11.9, colW: [3.2, 3.7, 5.0], fontSize: 20, color: TXT, fill: { color: "FFFFFF" }, border: { type: "solid", color: "E8B4C6", pt: 1 }, rowH: 0.85, valign: "middle" });
s.addText("Lo mejor: combinarlos y comprobarlo todo con el supervisor", { x: 0.7, y: 6.0, w: 12, h: 0.5, fontSize: 18, italic: true, color: OSCURO, isTextBox: true });

// 7 nueva planificación por qué
s = slide("Pregunta 3: ¿nueva planificación? SÍ", "Creemos que sí hace falta una planificación nueva porque con la fusión cambia todo: la estrategia, la estructura y lo que se necesita. Y como dice el tema, la planificación sale de la estrategia de la empresa, por ejemplo de reducir costes, que es justo lo que pasa aquí.");
bullets(s, ["Con la fusión cambia casi todo:", "-La estrategia: ahorrar costes y aprovechar sinergias", "-La estructura: oficinas cerca unas de otras y departamentos repetidos", "-Lo que se necesita: digitalización, perfiles tecnológicos", "No tiene sentido seguir con dos planificaciones separadas", "“La planificación surge de los planes estratégicos de la empresa, como reducir costes” (Tema 2)"]);

// 8 pasos planificación
s = slide("Pregunta 3: pasos de la planificación", "Estos serían los pasos. Lo importante es que al comparar lo que se necesita con lo que hay, en una fusión pasa todo a la vez: sobra gente en administración y falta en perfiles digitales. Y las medidas para el exceso hay que negociarlas.");
s.addText([
  "Ver los objetivos y la estrategia del nuevo banco",
  "Calcular cuánta gente va a hacer falta y con qué perfil",
  "Ver qué personal hay dentro (inventarios de cualificaciones)",
  "Buscar fuera solo si falta algún perfil (ej. tecnología)",
  "Comparar lo que se necesita con lo que hay (exceso, déficit o equilibrio)",
  "Tomar medidas: movilidad, formación, bajas voluntarias, prejubilaciones…",
  "Revisar cada cierto tiempo si funciona",
].map((t, i, a) => ({ text: t, options: { bullet: { type: "number" }, breakLine: i < a.length - 1, paraSpaceAfter: 10 } })), { x: 0.7, y: 1.6, w: 12, h: 5.2, fontSize: 23, color: TXT, valign: "top", isTextBox: true });

// 9 ERE
s = slide("Lo que pasó en realidad", "En abril de 2021 CaixaBank propuso 8.291 salidas y, después de negociar con los sindicatos, en julio se firmó un ERE de 6.452 salidas voluntarias, un 22 % menos. Con esto se ve que negociar con información sirvió para reducir el impacto.");
s.addChart(pres.charts.BAR, [{ name: "Salidas", labels: ["Propuesta (abril 2021)", "Acuerdo (julio 2021)"], values: [8291, 6452] }], {
  x: 0.7, y: 1.6, w: 7.0, h: 5.2, barDir: "col", chartColors: ["F59AB5", ROSA], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 16, dataLabelFormatCode: "#,##0",
  showLegend: false, valAxisHidden: true, valGridLine: { style: "none" }, catAxisLabelFontSize: 14, catAxisLabelColor: TXT, valAxisMinVal: 0, valAxisMaxVal: 9500 });
bullets(s, ["ERE firmado con los sindicatos en julio de 2021", "22 % menos salidas de las propuestas", "Salidas voluntarias", "Coste: unos 1.900 millones €"], { x: 8.0, y: 2.0, w: 4.8, h: 4.5 });

// 10 planes previos
s = slide("Pregunta 4: ¿sirven los planes de antes?", "Los planes que tenía cada banco sirven, pero solo para empezar. Dan una foto de cómo estaba cada plantilla y ayudan a ver quién se puede recolocar. Pero se hicieron por separado y no contaban con las duplicidades. Por eso decimos que uno más uno no es dos.");
bullets(s, ["Sí sirven, pero solo como punto de partida:", "-Dan una “foto” de cada plantilla", "-Ayudan a ver duplicidades y a quién se puede recolocar o formar", "Pero no son suficientes:", "-Cada banco los hizo por su cuenta", "-No contaban con las duplicidades al juntarse", "-La digitalización cambia lo que se necesita"], { w: 8.0 });
s.addText("1 + 1 ≠ 2", { x: 8.9, y: 2.4, w: 3.8, h: 1.2, fontSize: 54, bold: true, color: ROSA, align: "center", isTextBox: true });
s.addText("Lo que necesita el nuevo banco no es la suma de lo que necesitaba cada uno", { x: 8.9, y: 3.7, w: 3.8, h: 1.5, fontSize: 17, color: OSCURO, align: "center", isTextBox: true });

// 11 conclusiones
s = slide("Conclusiones", "Para terminar: el APT es la base para saber qué sobra y qué falta, hace falta una planificación nueva, y los planes de antes solo sirven para empezar. Y si de verdad las personas son lo más importante, los ajustes tienen que hacerse con información y negociando. Gracias.");
bullets(s, ["El APT es la base para saber qué puestos sobran, cuáles faltan y quién se puede recolocar", "Hace falta una planificación nueva porque con la fusión cambia todo", "Los planes de antes solo sirven como punto de partida", "Si las personas son lo más importante, los ajustes hay que hacerlos con información y negociando"]);
s.addText("¡Gracias!", { x: 0.7, y: 5.8, w: 12, h: 0.8, fontSize: 32, bold: true, color: ROSA, isTextBox: true });

// 12 bibliografía
s = slide("Bibliografía", "Estas son las fuentes que hemos usado.");
bullets(s, ["Dessler, G. (2009). Administración de recursos humanos. Pearson.", "Gómez-Mejía, L. R., Balkin, D. B. y Cardy, R. L. (2008). Dirección y gestión de recursos humanos. Prentice Hall.", "Apuntes de Dirección de RRHH II, Temas 1 a 5 (UCA).", "Expansión (19/09/2020). Noticia de la fusión CaixaBank–Bankia.", "CaixaBank (18/09/2020). Comunicado de la fusión.", "Bolsamanía y Forbes (2021). Noticias sobre el ERE de CaixaBank."], { fontSize: 20 });

pres.writeFile({ fileName: OUT }).then(() => {
  fs.writeFileSync(OUT.replace(/\.pptx$/, "_notes.json"), JSON.stringify(NOTES, null, 1));
  console.log("ok", n);
});
