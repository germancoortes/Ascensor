const pptxgen = require("pptxgenjs"); const fs = require("fs"); const path = require("path");
const [dir, notesFile, out] = process.argv.slice(2);
const notes = JSON.parse(fs.readFileSync(notesFile));
const imgs = fs.readdirSync(dir).filter(f => /^s-\d+\.png$/.test(f)).sort();
const pres = new pptxgen(); pres.layout = "LAYOUT_WIDE"; pres.title = "Fusión CaixaBank y Bankia (versión móvil)";
imgs.forEach((f, i) => { const s = pres.addSlide(); s.background = { color: "FFFFFF" };
  s.addImage({ path: path.join(dir, f), x: 0, y: 0, w: 13.333, h: 7.5, altText: notes[i] ? notes[i].title : "Diapositiva " + (i + 1) });
  if (notes[i]) s.addNotes(notes[i].notes); });
pres.writeFile({ fileName: out }).then(() => console.log("ok", imgs.length));
