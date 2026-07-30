"""Build a JavaScript data bundle so index.html also works via file://."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
BUNDLE_PATH = DATA_DIR / "dashboard-data.js"


def write_local_bundle(data: dict[str, Any] | None = None) -> Path:
    if data is None:
        manifest = json.loads((DATA_DIR / "manifest.json").read_text(encoding="utf-8-sig"))
        shards = [
            json.loads((DATA_DIR / item["file"]).read_text(encoding="utf-8-sig"))
            for item in manifest["files"]
        ]
        data = {}
        for shard in shards:
            data.update(shard)

    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    BUNDLE_PATH.write_text(
        "window.__DASHBOARD_DATA__=" + payload + ";\n",
        encoding="utf-8",
    )
    return BUNDLE_PATH


if __name__ == "__main__":
    path = write_local_bundle()
    print(f"Wrote {path} ({path.stat().st_size:,} bytes)")
