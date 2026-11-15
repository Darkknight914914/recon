"""Builds Recon as one self-contained page: recon.html (for Claude) and index.html (GitHub Pages / offline)."""
import pathlib
root = pathlib.Path(__file__).resolve().parent
libs = "\n".join((root / f).read_text() for f in
                 ["vendor/sheet-engine.js", "vendor/sheet-io.js", "src/recon.js", "src/sample.js", "src/export.js"])
app = (root / "src/app.js").read_text()
for name, s in [("libs", libs), ("app", app)]:
    assert "</script" not in s.lower(), name
html = (root / "src/template.html").read_text().replace("/*@LIBS@*/", libs).replace("/*@APP@*/", app)
assert "�" not in html, "stray replacement character"
(root / "recon.html").write_text(html)
skeleton = ('<!doctype html><html lang="en"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
            '<style>body{margin:0}[hidden]{display:none!important}</style></head><body>')
(root / "index.html").write_text(skeleton + html + "</body></html>")
print("built recon.html and index.html", len(html) // 1024, "KB")
