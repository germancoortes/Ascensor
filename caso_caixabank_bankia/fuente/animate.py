"""Add slide transitions + automatic staggered fade-in entrance animations.
Shapes named aN_* animate at step N (same step = together)."""
import re, sys, zipfile, shutil, os, collections

src, dst = sys.argv[1], sys.argv[2]
STEP_MS, DUR = 350, 450
SECTION_TITLES = set()

def timing(groups):
    cid = [2]
    def nid():
        cid[0] += 1; return cid[0]
    pars = []
    for k, (step, ids) in enumerate(sorted(groups.items())):
        effs = []
        for j, spid in enumerate(ids):
            a, b, c = nid(), nid(), nid()
            node = "afterEffect" if j == 0 else "withEffect"
            effs.append(
                f'<p:par><p:cTn id="{a}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="{node}">'
                f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{b}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
                f'<p:to><p:strVal val="visible"/></p:to></p:set>'
                f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{c}" dur="{DUR}"/>'
                f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')
        pars.append((k, effs))
    inner = "".join(
        f'<p:par><p:cTn id="{nid()}" fill="hold"><p:stCondLst><p:cond delay="{k * STEP_MS}"/></p:stCondLst>'
        f'<p:childTnLst>{"".join(e)}</p:childTnLst></p:cTn></p:par>' for k, e in pars)
    return ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            f'<p:par><p:cTn id="{nid()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
            '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst>'
            f'<p:childTnLst>{inner}</p:childTnLst></p:cTn></p:par>'
            '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
            '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>')

zin = zipfile.ZipFile(src)
zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", item.filename)
    if m:
        xml = data.decode("utf8")
        groups = collections.defaultdict(list)
        for spid, step in re.findall(r'<p:cNvPr id="(\d+)" name="a(\d+)_[^"]*"', xml):
            groups[int(step)].append(spid)
        dark = 'name="a1_num"' in xml  # section divider
        trans = '<p:transition spd="slow"><p:push dir="u"/></p:transition>' if dark else '<p:transition spd="med"><p:fade/></p:transition>'
        add = trans + (timing(groups) if groups else "")
        if "<p:extLst>" in xml.split("</p:clrMapOvr>")[-1]:
            raise SystemExit("extLst after clrMapOvr not handled")
        xml = xml.replace("</p:sld>", add + "</p:sld>", 1)
        data = xml.encode("utf8")
    zout.writestr(item, data)
zout.close()
print("animated")
