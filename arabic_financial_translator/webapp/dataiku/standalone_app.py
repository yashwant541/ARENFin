#!/usr/bin/env python3
"""
standalone_app.py — run the webapp locally, outside Dataiku, for testing.

It imports the exact backend.py routes and serves the same three DSS assets
(body.html + style.css + app.js) as one page, with a getWebAppBackendUrl shim so
app.js is identical to what you paste into DSS.

    python webapp/dataiku/standalone_app.py
    # then open http://127.0.0.1:5000

In real Dataiku you do NOT use this file — you paste backend.py / body.html /
style.css / app.js into the webapp's four tabs. See README_WEBAPP.md.
"""
import pathlib
from flask import Response

import backend  # registers all /api/* routes on backend.app
app = backend.app

HERE = pathlib.Path(__file__).parent
BODY = (HERE / "body.html").read_text(encoding="utf-8")
CSS = (HERE / "style.css").read_text(encoding="utf-8")
JS = (HERE / "app.js").read_text(encoding="utf-8")

PAGE_HEAD = """<!doctype html><html lang="en"><head><meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Arabic &#8594; English Financial Converter</title>
<style>__CSS__</style></head>
<body style="margin:0;background:#eef1f8;">
<script>
  // DSS injects this in production; shim it so app.js is unchanged locally.
  window.getWebAppBackendUrl = function (p) { return p; };
</script>
"""
PAGE_TAIL = "\n<script>__JS__</script>\n</body></html>"


def _render_page() -> str:
    head = PAGE_HEAD.replace("__CSS__", CSS)
    tail = PAGE_TAIL.replace("__JS__", JS)
    return head + BODY + tail


@app.route("/", methods=["GET"])
def index():
    return Response(_render_page(), mimetype="text/html")


if __name__ == "__main__":
    print("Glossary:", backend.TRANSLATOR.stats())
    print("Serving on http://127.0.0.1:5000  (Ctrl-C to stop)")
    app.run(host="127.0.0.1", port=5000, debug=False)
