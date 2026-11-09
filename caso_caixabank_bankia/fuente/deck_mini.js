const pptxgen = require("pptxgenjs");
const fs = require("fs");
const OUT = process.argv[2];
const ROSA = "D6457A", OSCURO = "7B2D4F", FONDO = "FDEEF3", TXT = "333333";
const NOTES = [];
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "Caso práctico 1 - Fusión CaixaBank y Bankia";
let n = 0;
function slide(title, items, notes, o = {}) {
  const s = pres.addSlide();
  s.background = { color: FONDO };
  s.addText(title, { x: 0.7, y: 0.5, w: 12, h: 1.0, fontSize: 40, bold: true, color: OSCURO, isTextBox: true });
  if (items) s.addText(items.map((t, i) => ({ text: t, options: { bullet: o.num ? { type: "number" } : true, breakLine: i < items.length - 1, paraSpaceAfter: 18 } })),
    { x: 0.9, y: 1.8, w: o.w || 11.5, h: 5.0, fontSize: o.fs || 28, color: TXT, valign: "top", isTextBox: true });
  s.addNotes(notes); n++; NOTES.push({ n, title, notes });
  return s;
}

let s = pres.addSlide(); s.background = { color: ROSA };
s.addText("Fusión CaixaBank y Bankia", { x: 0.8, y: 2.2, w: 11.7, h: 1.4, fontSize: 52, bold: true, color: "FFFFFF", align: "center", isTextBox: true });
s.addText("Caso práctico 1 · Dirección de RRHH II", { x: 0.8, y: 3.7, w: 11.7, h: 0.6, fontSize: 24, color: "FFFFFF", align: "center", isTextBox: true });
s.addText("Integrantes:", { x: 0.8, y: 5.6, w: 11.7, h: 0.5, fontSize: 18, color: "FFFFFF", align: "center", isTextBox: true });
const p1 = "Hola, nuestro trabajo es sobre la fusión de CaixaBank y Bankia. Vamos a responder a las cuatro preguntas usando sobre todo el tema 2: el análisis de puestos y la planificación de recursos humanos.";
s.addNotes(p1); n++; NOTES.push({ n, title: "Portada", notes: p1 });

slide("Introducción", ["Fusión en septiembre de 2020", "Más de 51.000 trabajadores", "Objetivo: ahorrar costes", "Habrá ajuste de plantilla"],
  "En septiembre de 2020 CaixaBank absorbe a Bankia y sale el banco más grande de España, con más de 51.000 trabajadores. Querían ahorrar unos 770 millones al año, así que iba a haber ajuste de plantilla. Lo curioso es que CaixaBank dice que las personas son su principal activo.");

slide("1. ¿Qué es el APT?", ["Analiza el puesto, no a la persona", "Descripción del puesto", "Especificación del puesto"],
  "El APT es el procedimiento para saber qué obligaciones y habilidades tiene un puesto y qué tipo de persona hace falta. Se analiza el puesto, no a la persona. De ahí salen la descripción, que es lo que se hace en el puesto, y la especificación, que es lo que tiene que tener la persona.");

slide("1. El APT en la fusión", ["Es como un mapa de las dos plantillas", "Ver puestos duplicados", "¿Sirve para el ajuste? Sí", "Pero no decide quién sale"],
  "En la fusión el APT es como un mapa de las dos plantillas: sirve para ver qué puestos están repetidos, cuáles se llaman igual pero no hacen lo mismo y qué sabe hacer la gente. Para el ajuste de plantilla sí sirve, porque da información objetiva, pero por sí solo no dice cuántos tienen que salir: para eso está la planificación y hay que negociarlo con los sindicatos.");

slide("2. Pasos del APT", ["Para qué se usa", "Revisar organigramas", "Elegir puestos", "Analizar los puestos", "Comprobar con el jefe", "Descripción y especificación"],
  "Son los seis pasos del tema 2. Primero decidir para qué se usa la información, aquí para ver duplicidades. Luego revisar los organigramas de los dos bancos, elegir puestos representativos porque no se pueden mirar 51.000 puestos, analizarlos con entrevistas, cuestionarios, observación o diarios, comprobarlo con el trabajador y su jefe, porque la gente puede exagerar por miedo a perder el puesto, y por último hacer la descripción y la especificación.", { num: true, fs: 26 });

slide("3. ¿Nueva planificación?", ["SÍ", "Cambia la estrategia", "Cambia la estructura", "Cambia lo que se necesita"],
  "Creemos que sí hace falta una planificación nueva, porque con la fusión cambia todo: la estrategia, que ahora es ahorrar; la estructura, con oficinas y departamentos repetidos; y lo que se necesita, por la digitalización. Como dice el tema, la planificación sale de la estrategia de la empresa.");

slide("3. Pasos", ["Objetivos", "Cuánta gente hace falta", "Qué gente hay dentro", "Buscar fuera si falta", "Comparar", "Tomar medidas", "Revisar"],
  "Los pasos serían: ver los objetivos del nuevo banco, calcular cuánta gente va a hacer falta, ver qué personal hay dentro, buscar fuera solo si falta algún perfil, comparar lo que se necesita con lo que hay, tomar medidas como recolocar, formar, bajas voluntarias o prejubilaciones, y revisar cada cierto tiempo si funciona.", { num: true, fs: 24 });

s = slide("Lo que pasó", null,
  "En abril de 2021 CaixaBank propuso 8.291 salidas y, después de negociar con los sindicatos, en julio se firmó un ERE de 6.452 salidas voluntarias, un 22 % menos. Con esto se ve que negociar con información sirvió.");
s.addChart(pres.charts.BAR, [{ name: "Salidas", labels: ["Propuesta", "Acuerdo final"], values: [8291, 6452] }], {
  x: 1.5, y: 1.7, w: 10.3, h: 5.2, barDir: "col", chartColors: ["F59AB5", ROSA], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 20, dataLabelFormatCode: "#,##0",
  showLegend: false, valAxisHidden: true, valGridLine: { style: "none" }, catAxisLabelFontSize: 18, catAxisLabelColor: TXT, valAxisMinVal: 0, valAxisMaxVal: 9500,
  showTitle: true, title: "Salidas del ERE (2021)", titleFontSize: 18, titleColor: OSCURO });

s = slide("4. ¿Sirven los planes de antes?", ["Sí, para empezar", "Pero se hicieron por separado", "No contaban con duplicidades"],
  "Los planes que tenía cada banco antes sirven, pero solo para empezar: dan una foto de cada plantilla. El problema es que se hicieron por separado y no contaban con las duplicidades al juntarse. Por eso decimos que uno más uno no es dos: lo que necesita el banco nuevo no es la suma de lo que necesitaba cada uno.", { w: 7.5 });
s.addText("1 + 1 ≠ 2", { x: 8.5, y: 2.8, w: 4.3, h: 1.4, fontSize: 60, bold: true, color: ROSA, align: "center", isTextBox: true });

slide("Conclusiones", ["El APT es la base", "Hace falta planificar de nuevo", "Los planes de antes solo ayudan a empezar", "Ajustes negociando"],
  "Para terminar: el APT es la base para saber qué sobra y qué falta, hace falta una planificación nueva, los planes de antes solo sirven para empezar, y si las personas son lo más importante, los ajustes tienen que hacerse con información y negociando. Gracias.");

slide("Bibliografía", ["Dessler (2009). Administración de recursos humanos", "Gómez-Mejía et al. (2008). Dirección y gestión de RRHH", "Apuntes Temas 1-5 (UCA)", "Expansión (19/09/2020)", "Bolsamanía y Forbes (2021)"],
  "Estas son las fuentes que hemos usado.", { fs: 22 });

pres.writeFile({ fileName: OUT }).then(() => {
  fs.writeFileSync(OUT.replace(/\.pptx$/, "_notes.json"), JSON.stringify(NOTES, null, 1));
  console.log("ok", n);
});
