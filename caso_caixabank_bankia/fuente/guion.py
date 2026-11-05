import json, sys, html
notes = json.load(open(sys.argv[1]))
blocks = [(1, 4, "Ponente 1", "Introducción y contexto"), (5, 9, "Ponente 2", "Pregunta 1 · APT"), (10, 13, "Ponente 3", "Pregunta 2 · Pasos del APT"),
          (14, 18, "Ponente 4", "Pregunta 3 · Nueva planificación"), (19, 23, "Ponente 1", "Pregunta 4 y cierre")]
WPM = 135
total = 0
rows = []
for b0, b1, who, name in blocks:
    secs = 0; items = []
    for nt in notes[b0 - 1:b1]:
        s = max(15, round(len(nt["notes"].split()) / WPM * 60 / 5) * 5); secs += s
        items.append(f'<div class="sl"><div class="h"><span class="num">{nt["n"]}</span><span class="t">{html.escape(nt["title"])}</span><span class="tm">{s // 60}:{s % 60:02d}</span></div><p>{html.escape(nt["notes"])}</p></div>')
    total += secs
    rows.append(f'<section><div class="blk"><span class="who">{who}</span><span class="bn">{name}</span><span class="bt">Diapositivas {b0}–{b1} · ≈{round(secs / 60, 1)} min</span></div>{"".join(items)}</section>')
summary = "".join(f"<tr><td>{w}</td><td>{n}</td><td>{a}–{b}</td></tr>" for a, b, w, n in blocks)
doc = f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Guion de la exposición</title><style>
@page {{ size:A4; margin:18mm 18mm 18mm 18mm; }}
body {{ font-family:Carlito,Calibri,Arial,sans-serif; color:#1B2433; font-size:11.5pt; line-height:1.5; margin:0; }}
h1 {{ color:#0B2545; font-size:26pt; margin:0 0 1mm; }} .k {{ color:#0F8B8D; font-weight:700; letter-spacing:.12em; font-size:10pt; }}
.intro {{ background:#EEF3F7; border-radius:2mm; padding:4mm 5mm; margin:5mm 0; }}
table {{ border-collapse:collapse; width:100%; margin-top:2mm; }} td,th {{ padding:1.5mm 2.5mm; border-bottom:1px solid #D5DEE7; text-align:left; }} th {{ color:#0B2545; }}
.blk {{ background:#0B2545; color:#fff; border-radius:2mm; padding:3mm 5mm; margin:7mm 0 3mm; display:flex; gap:5mm; align-items:baseline; page-break-after:avoid; }}
.who {{ color:#F4A259; font-weight:700; }} .bn {{ font-weight:700; font-size:13pt; }} .bt {{ margin-left:auto; color:#BFD0E0; font-size:10pt; }}
.sl {{ page-break-inside:avoid; border-bottom:1px solid #E1E8EF; padding:2mm 0 1mm; }}
.h {{ display:flex; gap:3mm; align-items:center; }} .num {{ background:#0F8B8D; color:#fff; border-radius:50%; width:7mm; height:7mm; display:inline-flex; align-items:center; justify-content:center; font-weight:700; font-size:9.5pt; flex:none; }}
.t {{ font-weight:700; color:#0B2545; }} .tm {{ margin-left:auto; color:#5C6B7A; font-size:10pt; }}
p {{ margin:1.5mm 0 1mm 10mm; text-align:justify; }}
.tips li {{ margin-bottom:1mm; }}
</style></head><body>
<div class="k">CASO PRÁCTICO 1 · DIRECCIÓN DE RRHH II · 4º GADE</div>
<h1>Guion de la exposición</h1>
<div>Fusión CaixaBank – Bankia: APT y planificación de RRHH · 23 diapositivas · duración estimada ≈{round(total / 60)} min</div>
<div class="intro"><b>Reparto propuesto</b> (ajustad los nombres y el número de ponentes al grupo):
<table><tr><th>Ponente</th><th>Bloque</th><th>Diapositivas</th></tr>{summary}</table></div>
<div class="intro"><b>Consejos de exposición</b><ul class="tips">
<li>Las animaciones de cada diapositiva se lanzan solas al entrar: esperad un segundo a que aparezca el contenido antes de hablar.</li>
<li>No leáis la diapositiva: el texto de este guion está también en las notas del orador de cada diapositiva (vista Moderador en PowerPoint).</li>
<li>Remarcad la idea fuerza de cada pregunta: <i>P1 «el APT es el mapa»</i>, <i>P2 «seis pasos + triangulación»</i>, <i>P3 «sí, siete pasos y negociación»</i>, <i>P4 «1 + 1 ≠ 2»</i>.</li>
<li>Al citar datos, mencionad la fuente (pie de la diapositiva). La ficha de la diapositiva 13 es un ejemplo ilustrativo: decidlo.</li>
<li>Preguntas probables: ¿el APT decide los despidos? (no, aporta información) · ¿por qué no basta sumar los planes previos? (duplicidades y nueva estrategia) · ¿qué relación tiene con retribución? (valoración de puestos y homologación salarial).</li>
</ul></div>
{"".join(rows)}
</body></html>"""
open(sys.argv[2], "w").write(doc)
print(round(total / 60, 1), "min")
