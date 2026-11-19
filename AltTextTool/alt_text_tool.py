#!/usr/bin/env python3
"""
PPT Alt-Text Filler  (S4Carlisle)
---------------------------------
Reads the publisher's Alt-Text Excel log + a PowerPoint, auto-matches every
picture to its alt text, writes a REVIEW workbook (with thumbnails), waits for
you to check/edit it, then writes the alt text into a NEW copy of the PPT.

Usage
  python alt_text_tool.py                      (interactive prompts / drag & drop)
  python alt_text_tool.py deck.pptx log.xlsx   (non-interactive: review + apply)
  python alt_text_tool.py deck.pptx log.xlsx --no-pause

Needs:  pip install python-pptx openpyxl pillow
"""
import io, os, re, sys, copy, argparse
from pptx import Presentation
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE
from lxml import etree
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.drawing.image import Image as XLImage

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
DECOR_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}"
DECOR_NS = "http://schemas.microsoft.com/office/drawing/2017/decorative"

STOP = set("""a an the of in on at to for and or with by from is are was were be as this that
these those it its into over under shows show showing displays display depicts depict illustrates
illustration photo picture image graphic labeled reads read text figure table titled title""".split())

AUTO_THRESHOLD = 0.55   # match score needed to pre-fill
MARGIN = 0.12           # best must beat 2nd best by this much


# ----------------------------------------------------------------- helpers
def clean_path(s):
    return s.strip().strip('"').strip("'")


ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def xl(v):
    return ILLEGAL.sub(" ", v) if isinstance(v, str) else v


def toks(text):
    out = []
    for w in re.findall(r"[a-z0-9]+", (text or "").lower()):
        if w in STOP:
            continue
        if len(w) > 3 and w.endswith("s"):
            w = w[:-1]
        out.append(w)
    return out


def walk(shapes):
    for s in shapes:
        yield s
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk(s.shapes)


def is_visual(shape):
    """Pictures (and video/media pictures) = p:pic elements."""
    return shape._element.tag == f"{{{NS_P}}}pic"


def cnvpr(shape):
    return shape._element.xpath(".//p:cNvPr")[0]


def get_decorative(shape):
    return bool(shape._element.xpath('.//*[local-name()="decorative"]'))


def chapter_from_name(path):
    m = re.search(r"(?:PPT|CH|Chapter)[_\- ]?0*(\d{1,3})", os.path.basename(path), re.I)
    return m.group(1).zfill(2) if m else None


# ----------------------------------------------------------------- excel
def load_log(xlsx):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ws = None
    # prefer the sheet that has the 'Alt Text - Short' header and most rows
    best = 0
    for sh in wb.worksheets:
        hdr = [str(c.value or "").strip().lower() for c in sh[1]]
        if any(h.startswith("alt text - short") for h in hdr) and sh.max_row > best \
                and sh.title.lower() not in ("examples",):
            ws, best = sh, sh.max_row
    if ws is None:
        sys.exit("Could not find a sheet with an 'Alt Text - Short' column.")
    hdr = {str(c.value or "").strip().lower(): i for i, c in enumerate(ws[1])}

    def col(*names):
        for n in names:
            for h, i in hdr.items():
                if h.startswith(n):
                    return i
        return None

    ci = dict(prefix=col("chapter prefix"), ch=col("chapter number"), el=col("element number"),
              usage=col("usage"), dec=col("decorative"), short=col("alt text - short"),
              long=col("alt text - long"), fig=col("figure naming"), page=col("page no"))
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        g = lambda k: (str(r[ci[k]]).strip() if ci[k] is not None and r[ci[k]] is not None else "")
        if not (g("short") or g("el")):
            continue
        rows.append(dict(prefix=g("prefix"), ch=g("ch").zfill(2) if g("ch").isdigit() else g("ch"),
                         el=g("el"), usage=g("usage"), dec=g("dec").lower() in ("yes", "y"),
                         short=g("short"), long=g("long"), fig=g("fig"), page=g("page")))
    return rows


# ----------------------------------------------------------------- pptx
def slide_context(slide):
    """title, other text, and per-picture nearest caption"""
    title, texts = "", []
    for s in slide.shapes:
        if s.has_text_frame and s.text_frame.text.strip():
            t = s.text_frame.text.strip()
            if s == slide.shapes.title:
                title = t
            elif not re.match(r"^(©|courtesy|top |bottom |left |right |source)", t, re.I):
                texts.append((s, t))
    return title, texts


def nearest_text(pic, texts):
    best, bd = "", None
    try:
        cx, cy = pic.left + pic.width / 2, pic.top + pic.height / 2
    except Exception:
        return ""
    for s, t in texts:
        try:
            d = ((s.left + s.width / 2 - cx) ** 2 + (s.top + s.height / 2 - cy) ** 2) ** .5
        except Exception:
            continue
        if bd is None or d < bd:
            best, bd = t, d
    return best


def collect_pictures(prs):
    pics = []
    for n, slide in enumerate(prs.slides, 1):
        title, texts = slide_context(slide)
        for s in walk(slide.shapes):
            if not is_visual(s):
                continue
            pr = cnvpr(s)
            img = None
            try:
                img = s.image.blob
            except Exception:
                pass
            pics.append(dict(slide=n, shape=s, id=int(pr.get("id")), name=pr.get("name"),
                             alt=(pr.get("descr") or "").strip(), decor=get_decorative(s),
                             title=title, near=nearest_text(s, texts),
                             allText=" ".join(t for _, t in texts), blob=img))
    return pics


# ----------------------------------------------------------------- matching
import math
ABBR = {"ad": "additional duty", "ems": "emergency medical services", "emt": "emergency medical technician"}


def toks(text):  # noqa: F811  (expands abbreviations, keeps digits)
    t = (text or "").lower()
    for k, v in ABBR.items():
        t = re.sub(rf"\b{k}\b", v, t)
    out = []
    for w in re.findall(r"[a-z0-9]+", t):
        if w in STOP:
            continue
        if len(w) > 3 and w.endswith("s"):
            w = w[:-1]
        out.append(w)
    return out


def row_key(rows):
    seen = {}
    for r in rows:
        k = r["el"] or "?"
        seen[k] = seen.get(k, 0) + 1
        r["key"] = k if seen[k] == 1 else f"{k} #{seen[k]}"


def match(pics, rows):
    cands = [r for r in rows if r["short"]]
    row_key(cands)
    docs = [set(toks(r["short"] + " " + r["long"][:300] + " " + r["el"])) for r in cands]
    n = len(docs)
    df = {}
    for d in docs:
        for w in d:
            df[w] = df.get(w, 0) + 1
    idf = lambda w: math.log((n + 1) / (df.get(w, 0) + 0.5))
    pairs = []
    for pi, p in enumerate(pics):
        # weight: nearest caption x3, title x1.5, rest of slide x1, existing alt x4
        weights = {}
        for text, wt in ((p["near"], 3), (p["title"], 1.5), (p["allText"], 1), (p["alt"], 4)):
            for w in set(toks(text)):
                weights[w] = max(weights.get(w, 0), wt)
        tot = sum(weights[w] * (idf(w) if w in df else 0.6) for w in weights) or 1
        sc = []
        for ri, d in enumerate(docs):
            hit = sum(weights[w] * idf(w) for w in weights if w in d)
            sc.append(hit / tot)
        order = sorted(range(n), key=lambda i: -sc[i])
        best = sc[order[0]] if order else 0
        second = sc[order[1]] if len(order) > 1 else 0
        for ri in order[:4]:
            pairs.append((sc[ri], pi, ri, best - second if ri == order[0] else 0))
    pairs.sort(reverse=True)
    used_p, used_r, res = set(), set(), {}
    for sc, pi, ri, margin in pairs:
        if pi in used_p or ri in used_r:
            continue
        res[pi] = (cands[ri], sc, margin)
        used_p.add(pi); used_r.add(ri)
    runner = {}
    for sc, pi, ri, _ in pairs:
        if pi in res and cands[ri] is not res[pi][0] and pi not in runner:
            runner[pi] = cands[ri]
    return res, runner, [c for i, c in enumerate(cands) if i not in used_r], cands


# ----------------------------------------------------------------- review workbook
def write_review(path, pics, res, runner, cands):
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Review"
    L = wb.create_sheet("Alt Text List")
    L.append(["Element ID", "Page", "Usage", "Short alt text", "Long alt text", "Decorative"])
    for r in cands:
        L.append([xl(v) for v in (r["key"], r["page"], r["usage"], r["short"], r["long"], "Yes" if r["dec"] else "")])
    for col, w in zip("ABCDEF", (12, 8, 8, 80, 80, 10)):
        L.column_dimensions[col].width = w
    for c in L[1]:
        c.font = Font(bold=True)
    head = ["Slide", "Shape ID", "Picture", "Thumbnail", "Slide title / nearby text", "Existing alt",
            "Auto-match", "Score", "ELEMENT ID  (pick from list)", "OR type your own alt text",
            "Decorative? (Yes)", "Also append long desc? (Yes)", "RESULT (alt text that will be written)"]
    ws.append(head)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, w in enumerate([7, 8, 13, 22, 36, 30, 11, 7, 16, 40, 12, 14, 70]):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    ws.row_dimensions[1].height = 45
    green, amber, red = (PatternFill("solid", fgColor=x) for x in ("E2EFDA", "FFF2CC", "F8CBAD"))
    n_rows = len(cands) + 1
    dv = DataValidation(type="list", formula1=f"='Alt Text List'!$A$2:$A${n_rows}", allow_blank=True,
                        showErrorMessage=False)
    ws.add_data_validation(dv)
    for pi, p in enumerate(pics):
        r = res.get(pi)
        row_, conf, margin = r if r else (None, 0, 0)
        auto = bool(row_) and conf >= AUTO_THRESHOLD and margin >= MARGIN
        ws.append([xl(v) for v in [
            p["slide"], p["id"], p["name"], "", f'{p["title"]}\n→ {p["near"][:150]}'.strip(), p["alt"],
            f'{row_["key"]}' if row_ else "", round(conf, 2),
            row_["key"] if auto else "", "", "Yes" if p["decor"] else "", "", ""]])
        rr = ws.max_row
        ws.cell(rr, 13, f'=IF(K{rr}="Yes","(decorative)",IF(J{rr}<>"",J{rr},IFERROR(INDEX(\'Alt Text List\'!$D:$D,MATCH(I{rr},\'Alt Text List\'!$A:$A,0))'
                        f'&IF(L{rr}="Yes"," "&INDEX(\'Alt Text List\'!$E:$E,MATCH(I{rr},\'Alt Text List\'!$A:$A,0)),""),"")))')
        dv.add(ws.cell(rr, 9))
        ws.row_dimensions[rr].height = 80
        for c in ws[rr]:
            c.alignment = Alignment(wrap_text=True, vertical="top")
        fill = green if auto and conf >= 0.8 else amber if auto else red
        for col in (7, 8, 9):
            ws.cell(rr, col).fill = fill
        if p["blob"]:
            try:
                from PIL import Image
                im = Image.open(io.BytesIO(p["blob"])).convert("RGB")
                im.thumbnail((150, 100))
                b = io.BytesIO(); im.save(b, "PNG"); b.seek(0)
                ws.add_image(XLImage(b), f"D{rr}")
            except Exception:
                ws.cell(rr, 4, "(no preview - EMF/TIF)")
    ws.freeze_panes = "E2"
    n = wb.create_sheet("How to review")
    for line in ["GREEN = confident auto-match.  AMBER = probable - check the thumbnail.  RED = no match - pick an Element ID.",
                 "Column I: choose the Element ID from the dropdown (see 'Alt Text List' sheet for the texts).",
                 "Column J: or type your own alt text (overrides column I).  Column K: type Yes for decorative images.",
                 "Column M shows exactly what will be written to the PPT.",
                 "Save and CLOSE this file, return to the command window and press Enter."]:
        n.append([line])
    n.column_dimensions["A"].width = 120
    wb.save(path)


# ----------------------------------------------------------------- apply
def set_decorative(pr):
    pr.set("descr", "")
    for ext in pr.findall(f"{{{NS_A}}}extLst"):
        pr.remove(ext)
    ext = etree.SubElement(pr, f"{{{NS_A}}}extLst")
    e = etree.SubElement(ext, f"{{{NS_A}}}ext", uri=DECOR_URI)
    d = etree.SubElement(e, f"{{{DECOR_NS}}}decorative", nsmap={"adec": DECOR_NS}); d.set("val", "1")


def apply_review(pptx_in, review, pptx_out):
    """Resolves the result itself (does not rely on Excel's cached formula values)."""
    wb = openpyxl.load_workbook(review)
    look = {}
    for r in wb["Alt Text List"].iter_rows(min_row=2, values_only=True):
        if r[0]:
            look[str(r[0]).strip()] = (str(r[3] or "").strip(), str(r[4] or "").strip())
    edits = {}
    for r in wb["Review"].iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        eid, own = str(r[8] or "").strip(), str(r[9] or "").strip()
        dec = str(r[10] or "").strip().lower() in ("yes", "y")
        app = str(r[11] or "").strip().lower() in ("yes", "y")
        txt = own
        if not txt and eid in look:
            txt = look[eid][0]
            if app and look[eid][1]:
                txt = (txt + " " + look[eid][1]).strip()
        edits[(int(r[0]), int(r[1]))] = (re.sub(r"\s+", " ", txt).strip(), dec)
    prs = Presentation(pptx_in)
    done = decor = skipped = 0
    for n, slide in enumerate(prs.slides, 1):
        for s in walk(slide.shapes):
            if not is_visual(s):
                continue
            pr = cnvpr(s)
            k = (n, int(pr.get("id")))
            if k not in edits:
                continue
            txt, dec = edits[k]
            if dec:
                set_decorative(pr); decor += 1
            elif txt:
                pr.set("descr", txt); done += 1
            else:
                skipped += 1
    prs.save(pptx_out)
    return done, decor, skipped


# ----------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx", nargs="?"); ap.add_argument("xlsx", nargs="?")
    ap.add_argument("--chapter", help="chapter number to filter the log (default: from file name)")
    ap.add_argument("--no-pause", action="store_true", help="apply immediately without review pause")
    a = ap.parse_args()

    print("\n=== PPT Alt-Text Filler ===")
    pptx = clean_path(a.pptx or input("Drag your PPT here (or paste path) and press Enter:\n> "))
    xlsx = clean_path(a.xlsx or input("Drag the Alt-Text Excel here and press Enter:\n> "))
    ch = a.chapter or chapter_from_name(pptx)
    if not a.chapter and not a.pptx:
        x = input(f"Chapter number to use [{ch or 'all'}] (Enter to accept, or type e.g. 05, 'all'): ").strip()
        ch = (x.zfill(2) if x.isdigit() else None) if x else ch
        ch = None if (x or "").lower() == "all" else ch

    rows = load_log(xlsx)
    if ch:
        sub = [r for r in rows if r["ch"] == ch or r["ch"].lstrip("0") == ch.lstrip("0")]
        print(f"Chapter {ch}: {len(sub)} Excel rows (of {len(rows)} total)")
        rows = sub or rows
    prs = Presentation(pptx)
    pics = collect_pictures(prs)
    print(f"Found {len(pics)} pictures in {len(prs.slides)} slides.")

    res, runner, unused, cands = match(pics, rows)
    base = os.path.splitext(pptx)[0]
    review = base + "_AltText_REVIEW.xlsx"
    write_review(review, pics, res, runner, cands)
    auto = sum(1 for v in res.values() if v[1] >= AUTO_THRESHOLD and v[2] >= MARGIN)
    print(f"Auto-matched {auto}/{len(pics)}; the rest need a manual look.")
    print(f"\nReview file written:\n  {review}")
    if not a.no_pause:
        input("\nOpen it, check/edit columns I-L, SAVE & CLOSE it, then press Enter to apply... ")
    out = base + "_with_alt.pptx"
    done, decor, skipped = apply_review(pptx, review, out)
    print(f"\nDone. Alt text written: {done} | marked decorative: {decor} | left empty: {skipped}")
    print(f"Output: {out}\n")


if __name__ == "__main__":
    main()
