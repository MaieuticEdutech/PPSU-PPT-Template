#!/usr/bin/env python3
import hashlib
import json
import os
import random
import tempfile
import threading
import traceback
import uuid
import zipfile
from collections import Counter
from pathlib import Path

from pptx import Presentation

from flask import Flask, request, send_file, jsonify, abort
from flask_cors import CORS

from analyze_template import build_catalog, CATALOG_VERSION
from engine import build as run_engine
from engine import extract_raw_slides, KEEP_ORIGINAL
from render import render_deck_to_images, render_one_slide

APP_DIR = Path(__file__).parent
CACHE_DIR = APP_DIR / "catalog_cache"
CACHE_DIR.mkdir(exist_ok=True)
UPLOAD_DIR = Path(tempfile.gettempdir()) / "ppt_engine_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

# Live per-session build progress, kept in memory (single dev-server process).
# Read by /progress so the page can show "Deck 2 of 3 — rendering…".
PROGRESS = {}

app = Flask(__name__)

# Configure CORS: Allow localhost for local development and Netlify frontend url from environment variables
frontend_origins = [
    "http://localhost:5000",
    "http://127.0.0.1:5000",
    "http://localhost:3000",
    "http://localhost:8000",
    "http://127.0.0.1:8000",
]
frontend_url = os.environ.get("FRONTEND_URL")
if frontend_url:
    frontend_origins.append(frontend_url)
else:
    frontend_origins.append("*")  # Allow all if not specified, useful for initial Netlify setup

CORS(app, origins=frontend_origins)


def template_hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()[:16]


def get_catalog(template_path: Path) -> Path:
    h = template_hash(template_path)
    cached = CACHE_DIR / f"{h}.json"
    if cached.exists():
        try:
            if json.loads(cached.read_text()).get("version") == CATALOG_VERSION:
                return cached
        except Exception:
            pass  # unreadable / stale cache -> rebuild below
    catalog = build_catalog(template_path)
    cached.write_text(json.dumps(catalog, indent=2))
    return cached


def is_valid_pptx(path: Path) -> bool:
    """True if `path` is a readable .pptx package.

    A .pptx is a zip whose index sits at the END of the file, so a partially
    written / interrupted-sync upload keeps a valid header but has no central
    directory. Catching that here lets us name the offending file instead of
    blowing up later with an opaque path deep inside the engine."""
    try:
        with zipfile.ZipFile(path) as z:
            return "[Content_Types].xml" in z.namelist()
    except Exception:
        return False


def _work_dir(token: str) -> Path:
    # token is used to build a filesystem path, so reject anything that could
    # escape UPLOAD_DIR (path traversal / absolute paths)
    if not token or not token.isalnum():
        abort(404)
    work = UPLOAD_DIR / token
    if not work.is_dir():
        abort(404)
    return work


def _run_designs(work: Path):
    """(Re)build a finished deck for every stored raw deck in `work`, render
    each to preview images, and write meta.json. Because design selection in
    the engine is randomised, calling this again ("redesign") yields a fresh,
    non-repeating set of layouts. Returns the preview payload dict."""
    template_path = work / "template.pptx"
    catalog_path = get_catalog(template_path)
    inputs = json.loads((work / "inputs.json").read_text())
    raw_meta = inputs["raw_files"]
    skipped = list(inputs.get("skipped", []))

    # one shared rotation state across every uploaded deck, so designs are not
    # repeated across unit1.1 / 1.2 / 1.3
    usage_state = {}
    decks = []
    total = len(raw_meta)
    for i, raw_info in enumerate(raw_meta):
        raw_path = work / f"raw_{i}.pptx"
        out_path = work / f"final_{i}.pptx"
        # one unreadable/odd deck must not abort the whole batch — record it and
        # carry on so the other units still come back
        PROGRESS[work.name] = {"phase": "building", "deck_index": i + 1,
                               "deck_total": total, "deck_name": raw_info["filename"]}
        try:
            report = run_engine(raw_path, template_path, catalog_path, out_path,
                                usage_state=usage_state)
        except Exception as e:
            traceback.print_exc()
            skipped.append(f'{raw_info["filename"]} — could not be processed '
                           f'({e.__class__.__name__})')
            continue

        PROGRESS[work.name] = {"phase": "rendering", "deck_index": i + 1,
                               "deck_total": total, "deck_name": raw_info["filename"]}
        img_dir = work / "images" / str(i)
        images = render_deck_to_images(out_path, img_dir)
        slide_urls = [f"/preview/{work.name}/{i}/{p.name}" for p in images]

        # remember exactly which design each slide got, so a single slide can be
        # swapped later without re-rolling every other slide
        assignment = {str(m["raw_slide"]): m["template_slide"]
                      for m in report.get("matched", [])}
        (work / f"assign_{i}.json").write_text(json.dumps(assignment))

        decks.append({
            "idx": i,
            "stem": raw_info["stem"],
            "filename": raw_info["filename"],
            "report": report,
            "slides": slide_urls,
        })

    payload = {
        "token": work.name,
        "kind": "zip" if len(decks) > 1 else "pptx",
        "preview_available": any(d["slides"] for d in decks),
        "skipped": skipped,
        "decks": decks,
    }
    (work / "meta.json").write_text(json.dumps(payload))
    PROGRESS[work.name] = {"phase": "done", "deck_index": total, "deck_total": total}

    # Warm the design-thumbnail cache in the background (AFTER the preview is
    # ready), so the first "Change design" click is instant. This never touches
    # generation time — it runs once the deck is already on screen, and is a
    # no-op when the template's thumbnails are already cached.
    try:
        threading.Thread(target=ensure_design_thumbs, args=(work,),
                         daemon=True).start()
    except Exception:
        pass
    return payload


@app.route("/", methods=["GET"])
def index():
    return jsonify({"status": "ok", "service": "ppt-designer"})


@app.route("/progress/<token>", methods=["GET"])
def progress(token):
    """Live build progress for the page's loader ('Deck 2 of 3 — rendering…')."""
    return jsonify(PROGRESS.get(token, {"phase": "starting"}))


def _receive_upload():
    """Save the posted template + raw decks into a fresh session folder.

    Returns (work_dir, error_response). Exactly one of the two is None.
    """
    # accept one OR several raw decks (e.g. unit1.1 / 1.2 / 1.3) under "raw"
    raw_files = [f for f in request.files.getlist("raw") if f and f.filename]
    if not raw_files or "template" not in request.files:
        return None, (jsonify({"error": "Missing raw or template files"}), 400)

    template_file = request.files["template"]
    work = UPLOAD_DIR / uuid.uuid4().hex
    work.mkdir()

    template_path = work / "template.pptx"
    template_file.save(template_path)
    if not is_valid_pptx(template_path):
        return None, (jsonify({"error": (
            f'The template "{template_file.filename}" is not a readable .pptx '
            f'file. It looks incomplete or corrupted — try opening it in '
            f'PowerPoint and doing "Save As" to a fresh copy, then upload that.'
        )}), 400)

    # Validate each raw deck up front. A file that was still syncing (e.g.
    # OneDrive Files On-Demand) uploads with a valid header but a truncated
    # body; skipping it here keeps the rest of the batch working.
    raw_meta, skipped = [], []
    for raw_file in raw_files:
        candidate = work / f"raw_{len(raw_meta)}.pptx"
        raw_file.save(candidate)
        if not is_valid_pptx(candidate):
            skipped.append(f'{raw_file.filename} — not a readable .pptx '
                           f'(file is incomplete or corrupted)')
            candidate.unlink(missing_ok=True)
            continue
        stem = os.path.splitext(os.path.basename(raw_file.filename))[0]
        raw_meta.append({"stem": stem, "filename": raw_file.filename})

    if not raw_meta:
        return None, (jsonify({
            "error": "None of the uploaded files could be read as a "
                     "PowerPoint (.pptx) deck.", "skipped": skipped}), 400)

    (work / "inputs.json").write_text(
        json.dumps({"raw_files": raw_meta, "skipped": skipped}))
    return work, None


def _preflight(work: Path):
    """Fast look-ahead: what WOULD happen, without building or rendering.

    Only reads the decks and the template catalogue (seconds), so the user
    learns about point-count gaps before waiting through the slow render.
    """
    template_path = work / "template.pptx"
    catalog = json.loads(get_catalog(template_path).read_text())
    design_counts = {int(k): len(v) for k, v in catalog["patterns"].items()}

    inputs = json.loads((work / "inputs.json").read_text())
    decks, needed = [], Counter()
    for i, info in enumerate(inputs["raw_files"]):
        raw_path = work / f"raw_{i}.pptx"
        total = len(Presentation(str(raw_path)).slides)
        rows = extract_raw_slides(raw_path)
        counts = Counter(r["count"] for r in rows)
        needed.update(counts)
        usable = sum(n for c, n in counts.items() if c in design_counts)
        decks.append({
            "filename": info["filename"],
            "total_slides": total,
            "will_redesign": usable,
            "will_keep": total - usable,
            "counts": {str(c): n for c, n in sorted(counts.items())},
        })

    warnings = []
    for count, n in sorted(needed.items()):
        if count not in design_counts:
            warnings.append(
                f'{n} slide{"s" if n > 1 else ""} with {count} points — your '
                f'template has no {count}-item design, so {"they" if n > 1 else "it"} '
                f'will be kept exactly as {"they are" if n > 1 else "it is"}.')

    return {
        "token": work.name,
        "template_designs": {str(k): v for k, v in sorted(design_counts.items())},
        "decks": decks,
        "warnings": warnings,
        "skipped": inputs.get("skipped", []),
    }


@app.route("/analyze", methods=["POST"])
def analyze():
    """Upload + fast look-ahead. No building, no rendering — returns in seconds
    so point-count gaps surface before the slow work starts."""
    try:
        work, err = _receive_upload()
        if err:
            return err
        return jsonify(_preflight(work))
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500


@app.route("/build/<token>", methods=["POST"])
def build_deck(token):
    """The slow half: rebuild the decks and render the previews."""
    try:
        work = _work_dir(token)
        if not (work / "inputs.json").exists():
            return jsonify({"error": "Session expired, please upload again"}), 404
        payload = _run_designs(work)
        if not payload["decks"]:
            return jsonify({"error": "No deck could be generated.",
                            "skipped": payload.get("skipped", [])}), 400
        return jsonify(payload)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500


@app.route("/generate", methods=["POST"])
def generate():
    """Upload and build in a single call (kept for direct API use)."""
    try:
        work, err = _receive_upload()
        if err:
            return err
        payload = _run_designs(work)
        if not payload["decks"]:
            return jsonify({"error": "No deck could be generated.",
                            "skipped": payload.get("skipped", [])}), 400
        return jsonify(payload)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500


@app.route("/redesign/<token>", methods=["POST"])
def redesign(token):
    """Rebuild the same uploaded decks with a fresh (random, non-repeating) set
    of designs and return a new preview payload — no re-upload required."""
    try:
        work = _work_dir(token)
        if not (work / "inputs.json").exists():
            return jsonify({"error": "Session expired, please upload again"}), 404
        payload = _run_designs(work)
        return jsonify(payload)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500


@app.route("/preview/<token>/<int:idx>/<name>", methods=["GET"])
def preview_image(token, idx, name):
    work = _work_dir(token)
    if not name.endswith(".png") or "/" in name or "\\" in name or ".." in name:
        abort(404)
    img_path = work / "images" / str(idx) / name
    if not img_path.is_file():
        abort(404)
    return send_file(img_path, mimetype="image/png")


@app.route("/raw-image/<token>/<int:idx>/<int:raw_slide>", methods=["GET"])
def raw_image(token, idx, raw_slide):
    """Render the ORIGINAL raw slide on demand (for the before/after toggle).
    Rendered lazily and cached, so it costs nothing unless the user asks."""
    work = _work_dir(token)
    raw_path = work / f"raw_{idx}.pptx"
    if not raw_path.is_file():
        abort(404)
    img_dir = work / "raw_images" / str(idx)
    img = img_dir / f"slide_{raw_slide - 1}.png"
    if not img.is_file():
        render_one_slide(raw_path, raw_slide - 1, img)
    if not img.is_file():
        return jsonify({"error": "could not render the original slide"}), 500
    return jsonify({"image": f"/rawpreview/{token}/{idx}/{img.name}"})


@app.route("/rawpreview/<token>/<int:idx>/<name>", methods=["GET"])
def rawpreview_image(token, idx, name):
    work = _work_dir(token)
    if not name.endswith(".png") or "/" in name or "\\" in name or ".." in name:
        abort(404)
    img_path = work / "raw_images" / str(idx) / name
    if not img_path.is_file():
        abort(404)
    return send_file(img_path, mimetype="image/png")


def _deck_templates(work: Path):
    """All designs in the session's template, grouped by item count."""
    catalog = json.loads(get_catalog(work / "template.pptx").read_text())
    by_count = {}
    for cnt, entries in catalog["patterns"].items():
        for e in (entries if isinstance(entries, list) else [entries]):
            by_count.setdefault(int(cnt), []).append(e)
    return by_count


def _family(entry):
    """Design-family id: look-alike layouts share one, so 'already used' means
    the same VISUAL design, not merely the same slide file."""
    return entry.get("design_key") or entry["slide_file"]


# ---- template design thumbnails (for the "Change design" picker) ----------
# Rendered ONCE per template and cached by template hash, so they never touch
# the generation flow and are reused across every session using that template.
_THUMB_LOCKS = {}
_THUMB_GUARD = threading.Lock()


def _thumb_dir(h):
    return CACHE_DIR / "design_thumbs" / h


def _slide_index(slide_file):
    """template design 'slideN.xml' -> 0-based slide index (N-1)."""
    return int(slide_file.replace("slide", "").replace(".xml", "")) - 1


def ensure_design_thumbs(work: Path):
    """Render the template's design slides to cached thumbnails (idempotent,
    locked so two callers never render the same template twice). Returns the
    template hash, or None if nothing could be rendered."""
    template_path = work / "template.pptx"
    h = template_hash(template_path)
    d = _thumb_dir(h)
    if d.is_dir() and any(d.glob("slide_*.png")):
        return h
    with _THUMB_GUARD:
        lock = _THUMB_LOCKS.setdefault(h, threading.Lock())
    with lock:
        if d.is_dir() and any(d.glob("slide_*.png")):
            return h
        try:
            d.mkdir(parents=True, exist_ok=True)
            imgs = render_deck_to_images(template_path, d)
            return h if imgs else None
        except Exception:
            traceback.print_exc()
            return None


def _families_in_use(work: Path, skip_deck=None, skip_slide=None):
    """Design families already placed anywhere in this unit — across every deck
    (assign_*.json), optionally ignoring one slide (the one being changed)."""
    by_file = {}
    for entries in _deck_templates(work).values():
        for e in entries:
            by_file[e["slide_file"]] = e
    used = set()
    for ap in work.glob("assign_*.json"):
        try:
            this_deck = ap.stem == f"assign_{skip_deck}"
            for k, sf in json.loads(ap.read_text()).items():
                if sf == KEEP_ORIGINAL:
                    continue
                if this_deck and str(k) == str(skip_slide):
                    continue
                if sf in by_file:
                    used.add(_family(by_file[sf]))
        except Exception:
            pass
    return used


@app.route("/slide-options/<token>/<int:idx>/<int:raw_slide>", methods=["GET"])
def slide_options(token, idx, raw_slide):
    """Alternative designs available for one slide (same point count), plus the
    'keep original' option and which design it is using now."""
    work = _work_dir(token)
    meta = json.loads((work / "meta.json").read_text())
    deck = next((d for d in meta["decks"] if d["idx"] == idx), None)
    if deck is None:
        abort(404)

    row = next((s for s in deck["report"]["slides"]
                if s["raw_slide"] == raw_slide), None)
    if row is None:
        abort(404)

    match = next((m for m in deck["report"]["matched"]
                  if m["raw_slide"] == raw_slide), None)
    current = match["template_slide"] if match else None

    # point count straight from the deck, so a currently-kept slide can still be
    # designed if the template has a design for its number of points
    rows = extract_raw_slides(work / f"raw_{idx}.pptx")
    rs = next((r for r in rows if r["index"] == raw_slide), None)
    n_points = rs["count"] if rs else None

    # only offer designs NOT already used anywhere in the unit (any deck), so
    # the picker never suggests a look already present. The slide's OWN current
    # design is always kept (marked "in use"); designs used by other slides are
    # hidden. total_of_size lets the UI say "all N are already used" if empty.
    options, total_of_size = [], 0
    if rs:
        used = _families_in_use(work, skip_deck=idx, skip_slide=raw_slide)
        for e in _deck_templates(work).get(n_points, []):
            total_of_size += 1
            if _family(e) in used and e["slide_file"] != current:
                continue                       # already used elsewhere in the unit
            options.append({"design": e["slide_file"],
                            "title": e.get("title_text") or e["slide_file"],
                            "items": e["item_count"]})

    return jsonify({
        "raw_slide": raw_slide,
        "status": row["status"],
        "points": n_points,
        "current": current,
        "can_design": total_of_size > 0,
        "total_of_size": total_of_size,
        "options": options,          # unused designs of the same point count
    })


@app.route("/design-thumbs/<token>/<int:points>", methods=["GET"])
def design_thumbs(token, points):
    """Designs of a given point count, each with a small preview image (rendered
    once per template and cached). Used by the 'Change design' picker."""
    work = _work_dir(token)
    try:
        h = ensure_design_thumbs(work)
    except Exception:
        h = None
    out = []
    for e in _deck_templates(work).get(points, []):
        item = {"design": e["slide_file"],
                "title": e.get("title_text") or e["slide_file"],
                "items": e["item_count"], "image": None}
        if h:
            img = _thumb_dir(h) / f"slide_{_slide_index(e['slide_file'])}.png"
            if img.is_file():
                item["image"] = f"/design-thumb/{h}/{img.name}"
        out.append(item)
    return jsonify({"points": points, "designs": out})


@app.route("/design-thumb/<h>/<name>", methods=["GET"])
def design_thumb(h, name):
    if not h.isalnum() or not name.endswith(".png") or "/" in name \
            or "\\" in name or ".." in name:
        abort(404)
    p = _thumb_dir(h) / name
    if not p.is_file():
        abort(404)
    return send_file(p, mimetype="image/png")


@app.route("/swap/<token>/<int:idx>/<int:raw_slide>", methods=["POST"])
def swap_slide(token, idx, raw_slide):
    """Change ONE slide's design (or keep it as original / pick a random other
    one) while every other slide of the deck stays exactly as it was."""
    try:
        work = _work_dir(token)
        meta = json.loads((work / "meta.json").read_text())
        deck = next((d for d in meta["decks"] if d["idx"] == idx), None)
        if deck is None:
            return jsonify({"error": "unknown deck"}), 404

        want = (request.get_json(silent=True) or {}).get("design", "random")

        assign_path = work / f"assign_{idx}.json"
        assignment = json.loads(assign_path.read_text()) if assign_path.exists() else {}
        key = str(raw_slide)

        # point count for this slide, straight from the deck (works whether the
        # slide is currently redesigned or kept)
        rows = extract_raw_slides(work / f"raw_{idx}.pptx")
        rs = next((r for r in rows if r["index"] == raw_slide), None)
        entries = _deck_templates(work).get(rs["count"], []) if rs else []
        opts = [e["slide_file"] for e in entries]

        if want == "keep":
            assignment[key] = KEEP_ORIGINAL
        elif want == "random":
            current = assignment.get(key)
            # prefer a design whose look is NOT already used anywhere in this
            # unit (this deck or its siblings); fall back to any other design,
            # then to reuse, only once every design of this size is taken.
            used = _families_in_use(work, skip_deck=idx, skip_slide=raw_slide)
            fresh = [e["slide_file"] for e in entries
                     if _family(e) not in used and e["slide_file"] != current]
            others = [o for o in opts if o != current]
            pool = fresh or others or opts
            if not pool:
                return jsonify({"error": "no design of this point count exists"}), 400
            assignment[key] = random.choice(pool)
        elif want in opts:
            assignment[key] = want   # a specific design's slide_file
        else:
            return jsonify({"error": "that design does not fit this slide"}), 400

        # rebuild the ONE deck reproducibly (fixed assignment) — fast, no render
        template_path = work / "template.pptx"
        catalog_path = get_catalog(template_path)
        out_path = work / f"final_{idx}.pptx"
        report = run_engine(work / f"raw_{idx}.pptx", template_path, catalog_path,
                            out_path, usage_state={}, fixed=assignment)
        assignment = {str(m["raw_slide"]): m["template_slide"]
                      for m in report.get("matched", [])}
        if want == "keep":
            assignment[key] = KEEP_ORIGINAL
        assign_path.write_text(json.dumps(assignment))

        # re-render ONLY the changed slide (its position in the deck is fixed)
        out_slide = raw_slide - 1        # deck is built in raw-slide order
        img = work / "images" / str(idx) / f"slide_{out_slide}.png"
        render_one_slide(out_path, out_slide, img)

        # refresh this deck's report + a cache-busting image url
        deck["report"] = report
        base = f"/preview/{token}/{idx}/slide_{out_slide}.png"
        (work / "meta.json").write_text(json.dumps(meta))

        row = next((s for s in report["slides"] if s["raw_slide"] == raw_slide), {})
        return jsonify({
            "raw_slide": raw_slide,
            "image": f"{base}?v={uuid.uuid4().hex[:8]}",
            "status": row.get("status"),
            "reason": row.get("reason"),
            "warning": row.get("warning"),
            "report": report,
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500


@app.route("/download/<token>", methods=["GET"])
def download(token):
    work = _work_dir(token)
    meta = json.loads((work / "meta.json").read_text())

    # optional custom name from the client (extension is enforced by kind)
    raw_name = (request.args.get("name") or "").strip()
    safe = "".join(c for c in raw_name if c.isalnum() or c in " _-").strip()

    if meta["kind"] == "pptx":
        out_path = work / "final_0.pptx"
        stem = safe or f'{meta["decks"][0]["stem"]}_final'
        return send_file(out_path, as_attachment=True,
                         download_name=f"{stem}.pptx")

    # several decks -> zip them up (respecting the user's chosen zip name)
    zip_path = work / "final_decks.zip"
    seen = set()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for d in meta["decks"]:
            name = f'{d["stem"]}_final.pptx'
            n = 1
            while name in seen:  # guard against duplicate raw filenames
                name = f'{d["stem"]}_final_{n}.pptx'
                n += 1
            seen.add(name)
            zf.write(work / f'final_{d["idx"]}.pptx', name)

    zip_name = safe or "final_decks"
    return send_file(zip_path, as_attachment=True, download_name=f"{zip_name}.zip")


if __name__ == "__main__":
    # host 0.0.0.0 binds every network interface, so other machines on the same
    # network reach the API at http://<this-machine-LAN-IP>:5000
    # (PORT still overrides the default for hosted deployments like Railway).
    port = int(os.environ.get("PORT", 5000))

    # Debug is OFF by default: with the server exposed on the network the
    # Werkzeug debugger would be reachable by anyone on the LAN and its console
    # can run arbitrary code. Opt in locally with FLASK_DEBUG=1.
    debug = os.environ.get("FLASK_DEBUG") == "1"

    app.run(host="0.0.0.0", port=port, debug=debug, threaded=True)
