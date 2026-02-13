import pathlib

base = pathlib.Path(r"手机业务2025年终杜林报表")
css = (base / "dashboard-styles.css").read_text("utf-8")
js = (base / "dashboard-charts.js").read_text("utf-8")
html = (base / "dashboard.html").read_text("utf-8")

html = html.replace(
    '<link rel="stylesheet" href="dashboard-styles.css">',
    "<style>\n" + css + "\n</style>"
)
html = html.replace(
    '<script src="dashboard-charts.js"></script>',
    "<script>\n" + js + "\n</script>"
)

out = base / "太太享物2025投资者报告.html"
out.write_text(html, "utf-8")
print(f"All-in-one file created: {out}")
print(f"Size: {len(html)} bytes ({len(html)//1024} KB)")
