"""Regenerate references/post-index.md from the crawler's manifest.json.

Usage: python build_post_index.py [path/to/finance/manifest.json]
"""

import json
import sys
from itertools import groupby
from pathlib import Path

DEFAULT_MANIFEST = Path(__file__).resolve().parents[3] / "manifest.json"  # workspace/manifest.json (crawler output)
OUT = Path(__file__).resolve().parents[1] / "references" / "post-index.md"


def build(manifest_path: Path) -> str:
    posts = json.loads(manifest_path.read_text(encoding="utf-8"))["posts"]
    posts = sorted(posts, key=lambda p: p["published"], reverse=True)
    lines = [
        "# Post index: Musings on Markets (all posts)",
        f"> {len(posts)} posts, newest first. Format: date | title | labels. "
        "Grep this file by keyword, company or label to find the source post.",
        "",
    ]
    for year, group in groupby(posts, key=lambda p: p["published"][:4]):
        lines += [f"## {year}", ""]
        for p in group:
            labels = f" | {', '.join(p['labels'])}" if p["labels"] else ""
            lines.append(f"- {p['published'][:10]} | [{p['title']}]({p['url']}){labels}")
        lines.append("")
    return "\n".join(lines)


if __name__ == "__main__":
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_MANIFEST
    text = build(src)
    assert text.count("\n- ") == len(json.loads(src.read_text(encoding="utf-8"))["posts"]), "post count mismatch"
    OUT.write_text(text, encoding="utf-8")
    print(f"Wrote {OUT} ({text.count(chr(10) + '- ')} posts)")
