#!/usr/bin/env python3
"""Rebuild index.html (single-file app) from rpdc_live.html + xlsx.full.min.js.
Run after any edit to rpdc_live.html:  python3 build.py"""
html = open("rpdc_live.html").read()
js = open("xlsx.full.min.js").read()
tag = '<script src="xlsx.full.min.js"></script>'
assert tag in html, "script tag not found in rpdc_live.html"
assert "</script" not in js, "sheetjs must not contain a closing script tag"
open("index.html", "w").write(html.replace(tag, "<script>\n" + js + "\n</script>"))
print("✓ index.html rebuilt")
