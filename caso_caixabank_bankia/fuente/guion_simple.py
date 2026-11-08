import json, sys, html
notes = json.load(open(sys.argv[1]))
blocks = [(1, 2, "Persona 1", "Introducción"), (3, 4, "Persona 2", "Pregunta 1"), (5, 6, "Persona 3", "Pregunta 2"),
          (7, 9, "Persona 4", "Pregunta 3"), (10, 12, "Persona 1", "Pregunta 4 y conclusiones")]
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
body {{ font-family:Carlito,Calibri,Arial,sans-serif; color:#2E2228; font-size:11.5pt; line-height:1.5; margin:0; }}
h1 {{ color:#7B2D4F; font-size:26pt; margin:0 0 1mm; }} .k {{ color:#D6457A; font-weight:700; letter-spacing:.12em; font-size:10pt; }}
.intro {{ background:#FDF0F4; border-radius:2mm; padding:4mm 5mm; margin:5mm 0; }}
table {{ border-collapse:collapse; width:100%; margin-top:2mm; }} td,th {{ padding:1.5mm 2.5mm; border-bottom:1px solid #F0C9D6; text-align:left; }} th {{ color:#7B2D4F; }}
.blk {{ background:#7B2D4F; color:#fff; border-radius:2mm; padding:3mm 5mm; margin:7mm 0 3mm; display:flex; gap:5mm; align-items:baseline; page-break-after:avoid; }}
.who {{ color:#F59AB5; font-weight:700; }} .bn {{ font-weight:700; font-size:13pt; }} .bt {{ margin-left:auto; color:#F6D3DF; font-size:10pt; }}
.sl {{ page-break-inside:avoid; border-bottom:1px solid #F6DCE5; padding:2mm 0 1mm; }}
.h {{ display:flex; gap:3mm; align-items:center; }} .num {{ background:#D6457A; color:#fff; border-radius:50%; width:7mm; height:7mm; display:inline-flex; align-items:center; justify-content:center; font-weight:700; font-size:9.5pt; flex:none; }}
.t {{ font-weight:700; color:#7B2D4F; }} .tm {{ margin-left:auto; color:#8A6B77; font-size:10pt; }}
p {{ margin:1.5mm 0 1mm 10mm; text-align:justify; }}
.tips li {{ margin-bottom:1mm; }}
</style></head><body>
<div class="k">CASO PRÁCTICO 1 · DIRECCIÓN DE RRHH II</div>
<h1>Guion de la exposición</h1>
<div>Fusión CaixaBank y Bankia · 12 diapositivas · unos {round(total / 60)} minutos</div>
<div class="intro"><b>Cómo nos repartimos</b> (poned vuestros nombres):
<table><tr><th>Quién</th><th>Bloque</th><th>Diapositivas</th></tr>{summary}</table></div>
<div class="intro"><b>Cosas a tener en cuenta</b><ul class="tips">

<li>No leer la diapositiva tal cual. Lo que hay que decir está también en las notas de cada diapositiva.</li>
<li>Lo más importante de cada pregunta: P1 el APT es como un mapa · P2 los 6 pasos · P3 sí hace falta planificar de nuevo · P4 1 + 1 no es 2.</li>
<li>Posibles preguntas del profesor: ¿el APT decide a quién se despide? (no, solo da información) · ¿por qué no vale con juntar los planes de antes? (por las duplicidades y la nueva estrategia) · ¿qué tiene que ver con los sueldos? (para igualar sueldos hay que saber qué puestos son equivalentes).</li>
</ul></div>
{"".join(rows)}
</body></html>"""
open(sys.argv[2], "w").write(doc)
print(round(total / 60, 1), "min")
