"""Refresh the data cache version and local static build using project paths."""
from pathlib import Path
import re
import shutil
import sys


def refresh(root: Path, version: str) -> None:
    page = root / "overview.html"
    html = page.read_text(encoding="utf-8")
    html = re.sub(r'(dashboard-data\.js\?v=)[^"\']+', lambda match: match[1] + version, html)
    page.write_text(html, encoding="utf-8")
    print("[OK] Refreshed dashboard-data.js cache version", flush=True)
    destination = root / "dist"
    if destination.is_dir():
        (destination / "assets").mkdir(exist_ok=True)
        for source, target in (
            ("overview.html", "index.html"),
            ("overview.html", "overview.html"),
            ("dashboard-data.js", "dashboard-data.js"),
            ("assets/service-order-icons.js", "assets/service-order-icons.js"),
            ("assets/lucide-LICENSE", "assets/lucide-LICENSE"),
        ):
            shutil.copy2(root / source, destination / target)
        print("[OK] Synced website files to dist", flush=True)


if __name__ == "__main__":
    refresh(Path(__file__).resolve().parents[1], sys.argv[1])
