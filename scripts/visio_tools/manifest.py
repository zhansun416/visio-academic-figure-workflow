"""Source-led reconstruction ledger and draft/delivery gates."""
from __future__ import annotations

from .spec import numbers

TYPES = {"text", "shape", "line", "connector", "group", "picture-svg", "picture-raster", "picture-unknown"}
MODES = {"native-text":{"text"}, "native-shape":{"shape"}, "native-line":{"line"},
         "native-connector":{"connector"}, "native-group":{"group"}, "svg-picture":{"picture-svg"},
         "raster-picture":{"picture-raster"}, "mixed":TYPES, "deferred":set()}


def check(manifest, package=None, mode="draft"):
    issues, mappings, ids, page_numbers = [], {}, set(), set()
    delivery = mode == "delivery"
    def error(message):
        issues.append(message)
    def fidelity(record, label, required):
        if not record:
            if required:
                error(f"Missing visual review: {label}")
            return
        if record.get("status") not in {"verified", "review-needed", "failed", "not-applicable"}:
            error(f"Invalid visual status: {label}")
        if required and (record.get("status") != "verified" or not record.get("methods") or not record.get("notes")):
            error(f"Unverified visual review: {label}")
    def bbox(value, canvas, label):
        try:
            x,y,w,h = numbers([value.get(k) for k in ("x","y","w","h")],4,label)
            if min(x,y) < 0 or min(w,h) <= 0 or x+w > canvas[0]+.01 or y+h > canvas[1]+.01:
                error(f"Bounding box outside source canvas: {label}")
        except (ValueError, AttributeError) as exc:
            error(f"Invalid bbox {label}: {exc}")
    if manifest.get("schemaVersion") != "1.0" or not isinstance(manifest.get("pages"), list) or not manifest["pages"]:
        return dict(passed=False, issues=["Expected schemaVersion 1.0 and nonempty pages"], mode=mode)
    for page in manifest["pages"]:
        p = page.get("page")
        if not isinstance(p,int) or isinstance(p,bool) or p < 1 or p in page_numbers:
            error(f"Invalid/duplicate page: {p}")
        page_numbers.add(p)
        try:
            canvas = numbers([page.get("canvas",{}).get(k) for k in ("widthPx","heightPx")],2,"Canvas")
            if min(canvas) <= 0:
                raise ValueError("Dimensions must be positive")
        except ValueError as exc:
            error(str(exc)); continue
        fidelity(page.get("sourceReview"), f"page {p} full source", delivery)
        if not page.get("regions") and not (page.get("blank") and page.get("blankReason")):
            error(f"Page {p} has no source regions")
        for region in page.get("regions", []):
            key = (p, region.get("id"))
            if not key[1] or key in ids:
                error(f"Missing/duplicate region id: {key}")
            ids.add(key)
            bbox(region.get("bbox"),canvas,str(key))
            fidelity(region.get("fidelity"),str(key),delivery and region.get("critical",False))
            if not region.get("elements"):
                error(f"Region has no source elements: {key}")
            for element in region.get("elements", []):
                eid = (p,element.get("id"))
                if not eid[1] or eid in ids:
                    error(f"Missing/duplicate element id: {eid}")
                ids.add(eid)
                bbox(element.get("bbox"),canvas,str(eid))
                completion = element.get("completion",{}).get("status")
                expected = element.get("expectedEditability")
                representation = element.get("representation",{})
                actual_editability, rep_mode = representation.get("editability"), representation.get("mode")
                if expected not in {"native","asset"} or actual_editability not in {"native","asset","partial"} or rep_mode not in MODES:
                    error(f"Invalid representation/editability: {eid}")
                if completion not in {"complete","partial","deferred"} or (delivery and completion != "complete"):
                    error(f"Incomplete element: {eid}")
                if expected == "asset" and not (element.get("assetReason") and element.get("editabilityBoundary")):
                    error(f"Missing asset boundary: {eid}")
                if (expected == "native" and actual_editability != "native") or completion in {"partial","deferred"}:
                    degradation = element.get("degradation",{})
                    if not (degradation.get("reason") and degradation.get("editabilityBoundary")):
                        error(f"Undeclared degradation: {eid}")
                fidelity(element.get("fidelity"),str(eid),delivery)
                maps = element.get("objectMap",[])
                if delivery and not maps:
                    error(f"Missing object mapping: {eid}")
                for m in maps:
                    name, kind = m.get("name"),m.get("type")
                    mk = (p,name)
                    if not name or kind not in TYPES or mk in mappings:
                        error(f"Invalid/duplicate mapping: {mk}")
                    if kind not in MODES.get(rep_mode,set()):
                        error(f"Representation does not match object type: {eid} / {kind}")
                    if kind and kind.startswith("picture-") and actual_editability == "native":
                        error(f"Picture incorrectly declared native: {eid}")
                    mappings[mk] = kind
    if package:
        actual = {}
        for obj in package["objects"]:
            key = (obj["page"], obj["name"])
            if key in actual:
                error(f"Duplicate actual object name: {key}")
            actual[key] = obj["type"]
        for key, kind in mappings.items():
            if key not in actual:
                error(f"Mapped object missing in VSDX: {key}")
            elif actual[key] != kind:
                error(f"Object type mismatch {key}: declared {kind}, actual {actual[key]}")
        if delivery:
            for key in actual.keys()-mappings.keys():
                error(f"Unmapped VSDX object: {key}")
            if {p["page"] for p in package["pages"]} != page_numbers:
                error("Manifest pages do not match VSDX pages")
    return dict(passed=not issues, mode=mode, mappingCount=len(mappings), issues=issues,
                note="A valid ledger cannot prove that all source-image content was inventoried; compare full source and complex regions.")
