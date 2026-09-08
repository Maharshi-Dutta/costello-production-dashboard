"""Render docs/USER-GUIDE.md and docs/SUPPORT-AND-MAINTENANCE.md to print-ready
HTML (docs/pdf/*.html), which Playwright then prints to PDF.
Usage: python render_docs.py"""
import io, os, re, datetime
import markdown  # pip install markdown; then print docs/pdf/*.html to PDF from a browser (A4, background graphics on)
WEB = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOCS = [("USER-GUIDE.md", "User-Guide", "Production Dashboard", "User Guide"),
        ("SUPPORT-AND-MAINTENANCE.md", "Support-and-Maintenance", "Production Dashboard", "Support and Maintenance")]
CSS = """
@page { size: A4; margin: 18mm 16mm 20mm 16mm; }
html { font-family: "IBM Plex Sans", "Segoe UI", system-ui, sans-serif; font-size: 10.5pt; color: #17171a; line-height: 1.5; }
body { margin: 0; }
.cover { height: 250mm; display: flex; flex-direction: column; justify-content: flex-end; page-break-after: always; border-left: 6px solid #184f95; padding-left: 14mm; }
.cover .brand { font-family: "Barlow Condensed", "Arial Narrow", sans-serif; font-weight: 700; font-size: 20pt; letter-spacing: .04em; text-transform: uppercase; color: #6d6a62; }
.cover h1 { font-family: "Barlow Condensed", "Arial Narrow", sans-serif; font-size: 42pt; margin: 4mm 0 2mm; line-height: 1; }
.cover .sub { font-size: 13pt; color: #4d4a44; margin-bottom: 10mm; }
.cover .meta { font-size: 9.5pt; color: #6d6a62; }
.toc { page-break-after: always; }
.toc h2 { margin-top: 0; }
.toc ul { list-style: none; padding-left: 0; columns: 1; }
.toc li { margin: 2px 0; }
.toc ul ul { padding-left: 14px; font-size: 9.5pt; color: #4d4a44; }
.toc a { color: inherit; text-decoration: none; }
h1 { font-family: "Barlow Condensed", "Arial Narrow", sans-serif; font-size: 22pt; margin: 0 0 6mm; }
h2 { font-family: "Barlow Condensed", "Arial Narrow", sans-serif; font-size: 18pt; margin: 10mm 0 3mm; padding-bottom: 2mm; border-bottom: 2px solid #184f95; page-break-after: avoid; }
h3 { font-size: 12.5pt; margin: 7mm 0 2mm; page-break-after: avoid; }
h4 { font-size: 11pt; margin: 5mm 0 1.5mm; page-break-after: avoid; }
p, li { orphans: 3; widows: 3; }
img { max-width: 100%; height: auto; border: 1px solid #dcd9d2; border-radius: 4px; display: block; margin: 3mm auto 1mm; page-break-inside: avoid; }
img + p > em:only-child { display: block; text-align: center; font-size: 9pt; color: #6d6a62; margin-bottom: 5mm; }
table { border-collapse: collapse; width: 100%; font-size: 9.2pt; margin: 3mm 0 5mm; page-break-inside: auto; }
th, td { border: 1px solid #dcd9d2; padding: 4px 7px; vertical-align: top; text-align: left; }
th { background: #f2f1ee; font-weight: 600; }
tr { page-break-inside: avoid; }
code { font-family: Consolas, "Courier New", monospace; font-size: 9pt; background: #f2f1ee; padding: 1px 4px; border-radius: 3px; }
pre { background: #f2f1ee; padding: 8px 10px; border-radius: 4px; font-size: 8.6pt; overflow-x: auto; white-space: pre-wrap; page-break-inside: avoid; }
pre code { background: none; padding: 0; }
blockquote { border-left: 3px solid #184f95; margin: 3mm 0; padding: 1mm 4mm; color: #4d4a44; background: #f7f9fc; }
hr { border: 0; border-top: 1px solid #dcd9d2; margin: 6mm 0; }
.foot { position: fixed; bottom: -12mm; right: 0; font-size: 8pt; color: #8e8b84; }
"""
def slug(t): return re.sub(r"[^a-z0-9]+", "-", t.lower()).strip("-")
for src, outname, brand, title in DOCS:
    md = io.open(os.path.join(WEB, "docs", src), encoding="utf8").read()
    # drop the H1 (the cover carries it) and build a TOC from H2/H3
    lines = md.splitlines()
    if lines and lines[0].startswith("# "): lines = lines[1:]
    body_md = "\n".join(lines)
    toc = []
    for l in lines:
        m = re.match(r"^(##|###) (.+)$", l)
        if m: toc.append((len(m.group(1)), m.group(2).strip()))
    html = markdown.markdown(body_md, extensions=["tables", "fenced_code", "toc", "sane_lists"], extension_configs={"toc": {"slugify": lambda v, s: slug(v), "separator": "-"}})
    toc_html = "<div class='toc'><h2>Contents</h2><ul>"
    depth = 2
    for lvl, text in toc:
        if lvl == 3 and depth == 2: toc_html += "<ul>"; depth = 3
        elif lvl == 2 and depth == 3: toc_html += "</ul>"; depth = 2
        toc_html += "<li><a href='#%s'>%s</a></li>" % (slug(text), text)
    if depth == 3: toc_html += "</ul>"
    toc_html += "</ul></div>"
    today = datetime.date.today().strftime("%d %B %Y")
    page = ("<!doctype html><html><head><meta charset='utf-8'><base href='../'><title>%s — %s</title>"
            "<link rel='stylesheet' href='https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@600;700&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap'>"
            "<style>%s</style></head><body>"
            "<div class='cover'><div class='brand'>%s</div><h1>%s</h1><div class='sub'>Costello Windows production dashboard</div>"
            "<div class='meta'>Version of %s · build 20260908-1849 · source: docs/%s</div></div>%s%s</body></html>"
            % (brand, title, CSS, brand, title, today, src, toc_html, html))
    outp = os.path.join(WEB, "docs", "pdf", outname + ".html")
    io.open(outp, "w", encoding="utf8").write(page)
    print("wrote", outp, "| sections:", len(toc))
