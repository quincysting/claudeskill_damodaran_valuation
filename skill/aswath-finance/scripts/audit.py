"""Audit reference files against the source posts.

Checks, per reference file:
  numbers - every %/$/decimal figure must appear in a post cited on the same line
            (fallback: any post cited in the same section). Classifies each claim as
            OK / MISCITED (found only in other posts) / UNSUPPORTED (in no post text).
  dates   - "(YYYY-MM)" after a link must match the link's /YYYY/MM/ path (--fix-dates rewrites).
  quotes  - quoted text in "Signature lines" must appear verbatim in a cited post.

Usage: python audit.py [--fix-dates] [--ledger DIR] [--corpus DIR]
"""

import argparse
import glob
import json
import os
import re
from collections import Counter, defaultdict
from pathlib import Path

SKILL = Path(__file__).resolve().parents[1]
# post corpus (one markdown file per post): workspace/markdown next to workspace/skill/aswath-finance, or $ASWATH_CORPUS
CORPUS = Path(os.environ.get("ASWATH_CORPUS") or SKILL.parents[1] / "markdown")
URL_RE = re.compile(r"https?://aswathdamodaran\.blogspot\.com/(\d{4})/(\d{2})/([\w\-]+)\.html")
DATED_LINK_RE = re.compile(r"(\]\((https?://aswathdamodaran\.blogspot\.com/(\d{4})/(\d{2})/[\w\-]+\.html)\)[^()\[\]]{0,40}?)(\d{4})-(\d{2})")
NUM_RE = re.compile(r"(?<![\w.])\$?\d[\d,]*(?:\.\d+)?(?:%|\b)")
QUOTE_RE = re.compile(r"[\"“]([^\"”]{25,})[\"”]")


def norm(text: str) -> str:
    text = text.replace("\u2212", "-").replace("\u2013", "-").replace("&nbsp;", " ")
    text = text.translate(str.maketrans({"\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"'}))
    text = re.sub(r"(?<=\d),(?=\d{3})", "", text)
    return re.sub(r"\s+", " ", text)


def load_corpus(corpus: Path) -> dict[str, str]:
    """url slug key 'YYYY-MM-name' -> normalized post body."""
    return {Path(f).stem: norm(Path(f).read_text(encoding="utf-8")) for f in glob.glob(str(corpus / "*.md"))}


def key(m: re.Match) -> str:
    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"


def checkable(tok: str) -> str | None:
    """Return the numeric core worth checking, or None for years/tiny ints."""
    core = tok.lstrip("$").rstrip("%").replace(",", "")
    if "." not in core and not tok.endswith("%") and not tok.startswith("$"):
        return None  # bare integers (years, counts, list numbers) are too ambiguous
    if re.fullmatch(r"(19|20)\d\d", core):
        return None
    return core


def found(core: str, text: str) -> int:
    variants = {core, core.rstrip("0").rstrip(".") if "." in core else core}
    if "." in core:
        variants.add(core + "0")
    for v in variants:
        m = re.search(rf"(?<![\d.]){re.escape(v)}(?![\d])", text)
        if m:
            return m.start()
    return -1


def audit_file(path: Path, corpus: dict[str, str], fix_dates: bool) -> dict:
    lines = path.read_text(encoding="utf-8").splitlines()
    # section-level fallback citations (nearest ## / ### heading block)
    section_urls, current, sec_of = defaultdict(set), 0, []
    for i, line in enumerate(lines):
        if line.startswith("#"):
            current = i
        sec_of.append(current)
        section_urls[current].update(key(m) for m in URL_RE.finditer(line))

    stats, claims, quotes, date_fixes = Counter(), [], [], 0
    in_sig = False
    for i, line in enumerate(lines):
        if line.startswith("#"):
            in_sig = "signature" in line.lower()
        # dates
        for m in DATED_LINK_RE.finditer(line):
            if (m.group(3), m.group(4)) != (m.group(5), m.group(6)):
                stats["date_mismatch"] += 1
        if fix_dates:
            new = DATED_LINK_RE.sub(lambda m: f"{m.group(1)}{m.group(3)}-{m.group(4)}", line)
            date_fixes += new != line
            lines[i] = line = new
        cited = [key(m) for m in URL_RE.finditer(line)] or sorted(section_urls[sec_of[i]])
        cited = [k for k in cited if k in corpus]
        body = URL_RE.sub("", re.sub(r"\(\d{4}-\d{2}\)|\b\d{4}-\d{2}(-\d{2})?\b", "", line))
        # quotes
        if in_sig:
            for q in QUOTE_RE.findall(line):
                # elided quotes ("a ... b") must match fragment by fragment
                frags = [f.strip(" .,;:") for f in re.split(r"\.\.\.|…|\[\.\.\.\]", norm(q).lower())]
                frags = [f for f in frags if len(f) >= 12]
                ok = bool(frags) and any(all(f in corpus[k].lower() for f in frags) for k in cited)
                stats["quote_ok" if ok else "quote_bad"] += 1
                if not ok:
                    quotes.append({"line": i + 1, "quote": q, "cited": cited})
        # numbers
        cores = [c for t in NUM_RE.findall(norm(body)) if (c := checkable(t))]
        if not cores:
            continue
        missing, evidence = [], {}
        for c in cores:
            hit = next(((k, pos) for k in cited if (pos := found(c, corpus[k])) >= 0), None)
            if hit:
                k, pos = hit
                evidence[c] = f"{k}: …{corpus[k][max(0, pos - 160):pos + 120]}…"
            else:
                missing.append(c)
        if not missing:
            status = "OK"
        else:
            elsewhere = {c: [k for k, t in corpus.items() if found(c, t) >= 0][:3] for c in missing}
            status = "MISCITED" if all(elsewhere.values()) else "UNSUPPORTED"
            evidence["_missing"] = elsewhere
        stats[status] += 1
        claims.append({"line": i + 1, "status": status, "claim": line.strip()[:600],
                       "cited": cited, "evidence": evidence})
    if fix_dates and date_fixes:
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return {"stats": stats, "claims": claims, "quotes": quotes, "date_lines_fixed": date_fixes}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--fix-dates", action="store_true")
    ap.add_argument("--ledger", help="write per-file JSON ledgers here")
    ap.add_argument("--corpus", default=str(CORPUS))
    args = ap.parse_args()
    corpus = load_corpus(Path(args.corpus))
    if not corpus:  # the post corpus is a maintenance download, not part of the skill
        print(f"SKIP audit: no post corpus at {args.corpus} (set ASWATH_CORPUS or pass --corpus)")
        return
    total = Counter()
    print(f"{'file':32s} {'claims':>6s} {'OK':>5s} {'MISCIT':>6s} {'UNSUP':>5s} {'dates':>5s} {'badQ':>4s}")
    for f in sorted((SKILL / "references").glob("*.md")):
        if f.name in ("post-index.md", "valuation-engine.md") or f.name.startswith("tool-"):  # not from posts
            continue
        r = audit_file(f, corpus, args.fix_dates)
        s = r["stats"]
        total.update(s)
        n = s["OK"] + s["MISCITED"] + s["UNSUPPORTED"]
        print(f"{f.stem:32s} {n:6d} {s['OK']:5d} {s['MISCITED']:6d} {s['UNSUPPORTED']:5d} "
              f"{s['date_mismatch']:5d} {s['quote_bad']:4d}")
        if args.ledger:
            os.makedirs(args.ledger, exist_ok=True)
            Path(args.ledger, f"{f.stem}.json").write_text(json.dumps(r, ensure_ascii=False, indent=1), encoding="utf-8")
    n = total["OK"] + total["MISCITED"] + total["UNSUPPORTED"]
    print(f"{'TOTAL':32s} {n:6d} {total['OK']:5d} {total['MISCITED']:6d} {total['UNSUPPORTED']:5d} "
          f"{total['date_mismatch']:5d} {total['quote_bad']:4d}")
    print(f"number-supported claim rate: {total['OK'] / max(1, n):.1%} | quotes verbatim: "
          f"{total['quote_ok']}/{total['quote_ok'] + total['quote_bad']}")


if __name__ == "__main__":
    main()
