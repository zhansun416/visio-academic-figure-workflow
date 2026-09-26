from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import re
import shutil
import time
from urllib.parse import quote, urlencode
from urllib.request import urlopen, Request
import xml.etree.ElementTree as ET

from .common import read_json, write_json, sha256


def library_root(value=None):
    if value or os.environ.get("VISIO_FIGURE_SVG_LIBRARY"):
        return Path(value or os.environ["VISIO_FIGURE_SVG_LIBRARY"]).expanduser().resolve()
    documents = Path.home() / "Documents"
    if os.name == "nt":
        import ctypes
        buffer = ctypes.create_unicode_buffer(32768)
        if ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buffer) == 0:
            documents = Path(buffer.value)
    return documents / "Codex" / "svg-library"


@contextmanager
def library_lock(root, timeout=30):
    root = Path(root)
    root.mkdir(parents=True, exist_ok=True)
    lock = root / ".visio-python.lock"
    deadline = time.monotonic() + timeout
    while True:
        try:
            fd = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
            break
        except FileExistsError:
            if time.monotonic() >= deadline:
                raise TimeoutError(f"Library is locked: {lock}")
            time.sleep(.1)
    try:
        os.write(fd, str(os.getpid()).encode())
        yield
    finally:
        os.close(fd)
        lock.unlink()


def initialize(root):
    root = library_root(root)
    with library_lock(root):
        (root / "icons").mkdir(exist_ok=True)
        if not (root / "library.json").exists():
            write_json(root / "library.json", dict(libraryVersion=1, root=".", entries=[]))
    return dict(passed=True, libraryRoot=str(root))


def load_library(root):
    path = root / "library.json"
    if not path.exists():
        return dict(libraryVersion=1, root=".", entries=[])
    data = read_json(path)
    if isinstance(data, list):
        return dict(libraryVersion=1, root=".", entries=data)
    if "entries" not in data:
        data["entries"] = data.pop("assets", [])
    return data


def entry_path(root, entry):
    # Older manifests use Windows separators, including for relative paths.
    value = entry.get("path") or "icons/" + entry["file"]
    path = Path(value if os.name == "nt" else value.replace("\\", "/"))
    return path if path.is_absolute() else root / path


def assess(path):
    path = Path(path).resolve()
    raw = path.read_text(encoding="utf-8-sig")
    warnings = []
    tree = None
    try:
        if re.search(r"<!DOCTYPE|<!ENTITY", raw, re.I):
            raise ValueError("DTD/entity declarations are not accepted")
        tree = ET.fromstring(raw)
        if tree.tag != "{http://www.w3.org/2000/svg}svg":
            warnings.append("Missing SVG namespace/root")
        if not tree.get("viewBox"):
            warnings.append("Missing viewBox")
        ids = {e.get("id") for e in tree.iter() if e.get("id")}
        for element in tree.iter():
            local = element.tag.rsplit("}", 1)[-1]
            if local in {"script", "filter", "mask", "foreignObject", "animate", "animateMotion", "animateTransform", "set"}:
                warnings.append(f"Unsupported/risky SVG element: {local}")
            for key, value in element.attrib.items():
                key = key.rsplit("}", 1)[-1]
                if key.lower().startswith("on"):
                    warnings.append("Event handler in SVG")
                if key == "vector-effect" and value == "non-scaling-stroke":
                    warnings.append("Non-scaling stroke needs a custom import adapter")
                if key in {"href", "src"}:
                    if not value.startswith("#") or value[1:] not in ids:
                        warnings.append(f"External or unresolved reference: {value[:100]}")
            for ref in re.findall(r"url\(\s*['\"]?([^)'\"]+)", str(element.attrib) + (element.text or ""), re.I):
                if not ref.startswith("#") or ref[1:] not in ids:
                    warnings.append(f"External or unresolved CSS reference: {ref[:100]}")
        if re.search(r"@import", raw, re.I):
            warnings.append("External CSS import")
    except (ET.ParseError, ValueError) as error:
        warnings.append(str(error))
    colors = sorted(set(re.findall(r'''(?:fill|stroke)\s*=\s*["'](#[a-fA-F0-9]+|[a-zA-Z]+)["']''', raw)))
    return dict(path=str(path), sha256=sha256(path), validXml=tree is not None, viewBox=tree.get("viewBox") if tree is not None else None,
                usesCurrentColor=bool(re.search(r"\bcurrentColor\b", raw, re.I)), explicitColors=[c for c in colors if c not in {"none", "currentColor"}],
                warnings=sorted(set(warnings)), readyForVisioImport=not warnings)


def safe_copy(source, output, color="#1E2A44"):
    if not re.fullmatch(r"#[0-9a-fA-F]{6}", color):
        raise ValueError("SVG default color must be #RRGGBB")
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output:
        raise ValueError("Derived SVG must not overwrite its source")
    assessment = assess(source)
    if not assessment["readyForVisioImport"]:
        raise ValueError(f"Unsafe SVG {source}: {assessment['warnings']}")
    output.parent.mkdir(parents=True, exist_ok=True)
    raw = source.read_text(encoding="utf-8-sig")
    output.write_text(re.sub(r"\bcurrentColor\b", color, raw, flags=re.I), encoding="utf-8")
    return output


def prepare(directory, safe_dir=None, complete=False, color="#1E2A44"):
    records = []
    for path in sorted(Path(directory).glob("*.svg")):
        record = assess(path)
        if safe_dir and record["readyForVisioImport"]:
            record["visioSafePath"] = str(safe_copy(path, Path(safe_dir)/path.name, color))
        records.append(record)
    invalid = sum(not a["readyForVisioImport"] for a in records)
    return dict(passed=invalid == 0 and bool(records), assetCount=len(records), invalidAssetCount=invalid,
                skipIconfontSvgFinder=complete and bool(records) and not invalid, assets=records)


def sync(root, repair=False):
    root = library_root(root)
    def state():
        data = load_library(root)
        entries = data["entries"]
        existing = {p.name for p in (root/"icons").glob("*.svg")}
        indexed = {e["file"] for e in entries}
        missing, mismatch, duplicates = [], [], []
        seen = set()
        for e in entries:
            if e["file"] in seen:
                duplicates.append(e["file"])
            seen.add(e["file"])
            p = entry_path(root, e)
            if not p.is_file():
                missing.append(e["file"])
            elif e.get("sha256") and sha256(p) != e["sha256"].lower():
                mismatch.append(e["file"])
        return data, dict(libraryRoot=str(root), manifestEntryCount=len(entries), svgFileCount=len(existing),
                         unindexedFiles=sorted(existing-indexed), missingFiles=missing, hashMismatchFiles=mismatch,
                         duplicateEntries=duplicates, passed=not (existing-indexed or missing or mismatch or duplicates))
    if repair:
        with library_lock(root):
            data, report = state()
            new = []
            for filename in report["unindexedFiles"]:
                p = root/"icons"/filename
                if not assess(p)["readyForVisioImport"]:
                    raise ValueError(f"Cannot index risky SVG: {filename}")
                new.append(dict(file=filename, path="icons/"+filename, sha256=sha256(p), tags=re.split(r"\W+", p.stem),
                                license="unknown-review", sourceLabel="recovered-unindexed", sourceUrl=None))
            if new:
                data["entries"].extend(new)
                write_json(root/"library.json", data)
    return state()[1]


def find(root, query, limit=12):
    entries = load_library(library_root(root))["entries"]
    terms = query.lower().split()
    matches = []
    for e in entries:
        haystack = " ".join(str(e.get(k, "")) for k in ("file", "tags", "category", "sourceLabel")).lower()
        score = sum(t in haystack for t in terms)
        if score:
            matches.append(dict(e, score=score))
    return dict(matches=sorted(matches, key=lambda e: (-e["score"], e["file"]))[:limit], note="Candidates require visual matching; tags do not prove similarity.")


def register(root, source, **metadata):
    root, source = library_root(root), Path(source)
    files = sorted(source.glob("*.svg")) if source.is_dir() else [source]
    prepared = [(p, assess(p)) for p in files]
    if not prepared or any(not a["readyForVisioImport"] for p, a in prepared):
        raise ValueError("All registered files must be valid Visio-compatible SVGs")
    added, skipped = [], []
    with library_lock(root):
        data = load_library(root)
        (root/"icons").mkdir(exist_ok=True)
        for p, a in prepared:
            if any(e.get("sha256") == a["sha256"] for e in data["entries"]):
                skipped.append(p.name)
                continue
            filename = p.name
            target = root/"icons"/filename
            if target.exists() and sha256(target) != a["sha256"]:
                filename = f"{p.stem}-{a['sha256'][:12]}.svg"
                target = root/"icons"/filename
            if target.exists() and sha256(target) != a["sha256"]:
                raise ValueError(f"Asset name collision: {filename}")
            if target.resolve() != p.resolve():
                shutil.copy2(p, target)
            entry = dict(file=filename, path="icons/"+filename, sha256=a["sha256"], addedAt=datetime.now(timezone.utc).isoformat(),
                         tags=metadata.get("tags") or re.split(r"\W+", p.stem), category=metadata.get("category", "general"),
                         sourceLabel=metadata.get("sourceLabel", "user-supplied"), sourceUrl=metadata.get("sourceUrl"),
                         license=metadata.get("license", "unknown-review"), extractionMethod=metadata.get("extractionMethod", "file-copy"))
            data["entries"].append(entry)
            added.append(filename)
        if added:
            if (root/"library.json").exists():
                shutil.copy2(root/"library.json", root/"library.json.bak")
            write_json(root/"library.json", data)
    return dict(passed=True, addedFiles=added, skippedDuplicates=skipped)


def audit(root):
    root = library_root(root)
    report = sync(root)
    review = [e for e in load_library(root)["entries"] if not e.get("license") or re.search(r"unknown|review|unverified|proprietary", e["license"], re.I)]
    return dict(report, reviewRequired=review, publicReleaseReady=report["passed"] and not review)


def resolve_asset(ref, spec_dir, root=None, color="#1E2A44"):
    if ref.startswith("library:"):
        root = library_root(root)
        entries = [e for e in load_library(root)["entries"] if e["file"] == ref[8:]]
        if len(entries) != 1:
            raise ValueError(f"Expected one library entry for {ref}, found {len(entries)}")
        source = entry_path(root, entries[0])
        if entries[0].get("sha256") and sha256(source) != entries[0]["sha256"]:
            raise ValueError(f"Library hash mismatch for {ref}")
    else:
        source = Path(ref)
        if not source.is_absolute():
            source = Path(spec_dir)/source
    a = assess(source)
    if not a["readyForVisioImport"]:
        raise ValueError(f"SVG cannot be imported: {source}: {a['warnings']}")
    if a["usesCurrentColor"]:
        cache = Path(os.environ.get("LOCALAPPDATA", Path.home()/".cache"))/"Codex"/"visio-academic-figure-workflow"/"svg-cache"
        return safe_copy(source, cache/f"{source.stem}-{a['sha256'][:12]}-{color[1:]}.svg", color)
    return source.resolve()


def candidates(query, output, limit=12, download=3, prefix=None):
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=True)
    params = dict(query=query, limit=min(max(limit, 1), 64))
    if prefix:
        params["prefix"] = prefix
    def fetch(url):
        with urlopen(Request(url, headers={"User-Agent":"VisioFigureWorkflow/2"}), timeout=20) as response:
            return response.read()
    data = json.loads(fetch("https://api.iconify.design/search?"+urlencode(params)))
    records = []
    for i, icon in enumerate(data.get("icons", [])):
        collection, name = icon.split(":", 1)
        url = "https://api.iconify.design/"+quote(collection, safe="")+"/"+quote(name, safe="")+".svg"
        record = dict(icon=icon, sourceUrl=url, license=data.get("collections", {}).get(collection, {}).get("license"), matchLabel="candidate")
        if i < min(max(download, 0), 10):
            target = output/(re.sub(r"[^A-Za-z0-9_.-]", "_", icon)+".svg")
            target.write_bytes(fetch(url))
            record.update(localPath=str(target), assessment=assess(target))
        records.append(record)
    write_json(output/"iconify-candidates.json", records)
    return dict(candidates=records)
