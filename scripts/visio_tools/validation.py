from __future__ import annotations

from pathlib import Path
import math
import re

from .common import read_json
from .spec import expand
from .package import inspect_package
from .com_backend import application, walk, identity, value, center, page_point, endpoint


def validate(path, spec_path=None, expected_font=None, min_font=0, glued=False, no_raster=False, axis_aligned=False,
             tolerance_px=1, axis_tolerance=.0001, single_output=False, allowed_acronyms=None):
    if tolerance_px <= 0 or axis_tolerance <= 0 or not all(math.isfinite(x) for x in (tolerance_px, axis_tolerance, min_font)):
        raise ValueError("Invalid validation tolerance")
    package = inspect_package(path)
    spec = read_json(spec_path) if spec_path else None
    scene = expand(spec) if spec else None
    issues, fonts, connector_count, actual = [], set(), 0, {}
    if no_raster and package["rasterMedia"]:
        issues.append("Raster media present: "+", ".join(package["rasterMedia"]))
    if single_output:
        files = [p.resolve() for p in Path(path).resolve().parent.iterdir() if p.is_file()]
        if files != [Path(path).resolve()]:
            issues.append("Delivery directory must contain exactly the final VSDX")
    with application() as app:
        doc = app.Documents.Open(str(Path(path).resolve()))
        all_shapes, glue, parent_ids = {}, {}, {}
        for p in range(1, doc.Pages.Count+1):
            page = doc.Pages.Item(p)
            for s in walk(page.Shapes):
                sid = (p, s.ID)
                all_shapes[sid] = s
                key = identity(s)
                if key:
                    if key in actual:
                        issues.append(f"Duplicate output id: {key}")
                    actual[key] = s
                    parent = s.ContainingShape
                    parent_ids[key] = identity(parent) if parent and parent.ID else ""
                text = str(s.Text)
                if text.strip():
                    font = str(s.CellsU("Char.Font").FormulaU)
                    fonts.add(font)
                    if expected_font and expected_font.casefold() not in font.casefold():
                        issues.append(f"Font mismatch for {key or sid}: {font}")
                    if min_font and s.CellsU("Char.Size").Result("pt") < min_font:
                        issues.append(f"Small text for {key or sid}")
                    if re.search(r"`[nrt]", text):
                        issues.append(f"Literal PowerShell escape in {key or sid}")
                    if allowed_acronyms is not None:
                        for word in re.findall(r"\b[A-Z][A-Z0-9-]{2,}\b", text):
                            if word not in allowed_acronyms:
                                issues.append(f"Unlisted acronym {word} in {key or sid}")
                for i in range(1, s.Connects.Count+1):
                    connection = s.Connects.Item(i)
                    glue[(p, connection.FromSheet.ID, connection.FromCell.Name)] = (connection.ToSheet, connection.ToCell)
            for sid, s in list(all_shapes.items()):
                if sid[0] != p or not s.OneD or not s.Master or "connector" not in s.Master.NameU.lower():
                    continue
                connector_count += 1
                for end in ("Begin", "End"):
                    if glued and (*sid, end+"X") not in glue:
                        issues.append(f"Unglued {identity(s) or sid} {end}")
                targets = [glue.get((*sid, e+"X")) for e in ("Begin", "End")]
                if axis_aligned and all(targets):
                    a, b = [center(t[0]) for t in targets]
                    x, y = endpoint(s, "Begin"), endpoint(s, "End")
                    if abs(a[1]-b[1]) <= axis_tolerance and abs(a[0]-b[0]) > axis_tolerance and abs(x[1]-y[1]) > axis_tolerance:
                        issues.append(f"Slanted horizontal edge {identity(s)}")
                    if abs(a[0]-b[0]) <= axis_tolerance and abs(a[1]-b[1]) > axis_tolerance and abs(x[0]-y[0]) > axis_tolerance:
                        issues.append(f"Slanted vertical edge {identity(s)}")
        if scene:
            if doc.Pages.Count != 1:
                issues.append("Spec expects one page")
            rw, rh, w, h = spec["reference"]["widthPx"], spec["reference"]["heightPx"], spec["page"]["widthIn"], spec["page"]["heightIn"]
            def check_point(point, expected, context):
                px = (point[0]*rw/w, (h-point[1])*rh/h)
                if max(abs(px[i]-expected[i]) for i in (0, 1)) > tolerance_px:
                    issues.append(f"Position mismatch for {context}: expected {expected}, actual {px}")
            def parent_check(key, parent):
                expected = parent+"::group" if parent else ""
                if parent_ids.get(key, "") != expected:
                    issues.append(f"Group mismatch for {key}: expected {expected}")
            def text_check(s, item, key):
                normalized = lambda t:str(t).replace("\r\n", "\n").replace("\r", "\n").strip()
                if normalized(s.Text) != normalized(item.get("text", "")):
                    issues.append(f"Text mismatch for {key}")
            for item in scene.shapes:
                key = item["id"]
                if key not in actual:
                    issues.append(f"Missing shape {key}")
                    continue
                s = actual[key]
                text_check(s, item, key)
                x, y, bw, bh = item["bboxPx"]
                check_point(center(s), (x+bw/2, y+bh/2), key)
                if item["kind"] != "line" and max(abs(value(s, "Width")*rw/w-bw), abs(value(s, "Height")*rh/h-bh)) > tolerance_px:
                    issues.append(f"Size mismatch for {key}")
                parent_check(key, key if item["kind"] == "group" else item["parentId"])
            for group in scene.groups:
                key = group["id"]+"::group"
                if key not in actual:
                    issues.append(f"Missing native group {key}")
                else:
                    parent_check(key, group["parentId"])
            for edge in scene.connectors:
                key = edge["id"]
                if key not in actual:
                    issues.append(f"Missing edge {key}")
                    continue
                s = actual[key]
                parent_check(key, edge["parentId"])
                for end, prop in (("Begin", "from"), ("End", "to")):
                    target_id = edge[prop]
                    target = glue.get((1, s.ID, end+"X"))
                    if not target or target_id not in actual or target[0].ID != actual[target_id].ID:
                        issues.append(f"Wrong glue target for {key} {end}: {target_id}")
                    elif edge.get(prop+"Connection", "").startswith("port:"):
                        p = next(p for p in scene.by_id[target_id]["ports"] if p["id"] == edge[prop+"Connection"][5:])
                        if abs(target[1].ResultIU-value(target[0], "Width")*p["x"]) > 1e-6 or abs(target[0].CellsSRC(7,target[1].Row,1).ResultIU-value(target[0],"Height")*(1-p["y"])) > 1e-6:
                            issues.append(f"Wrong named port for {key} {end}")
                    expected_arrow = edge.get("arrow", "none") in {"both", end.lower()}
                    if (value(s, end+"Arrow") > 0) != expected_arrow:
                        issues.append(f"Wrong arrow for {key} {end}")
                points = edge.get("waypointsPx", [])
                if points:
                    if s.RowCount(11) != len(points)+3:
                        issues.append(f"Manual route vertex count mismatch for {key}")
                    else:
                        for j, expected in enumerate(points, 2):
                            check_point(page_point(s, s.CellsSRC(11,j,0).ResultIU, s.CellsSRC(11,j,1).ResultIU), expected, f"{key} waypoint {j-2}")
                        for end, row in (("Begin",1),("End",len(points)+2)):
                            visible = page_point(s, s.CellsSRC(11,row,0).ResultIU,s.CellsSRC(11,row,1).ResultIU)
                            e = endpoint(s,end)
                            if max(abs(visible[i]-e[i]) for i in (0,1)) > 1e-6:
                                issues.append(f"Visible route misses glued endpoint for {key} {end}")
                if edge.get("label"):
                    label = actual.get(key+"::label")
                    if label is None:
                        issues.append(f"Missing label for {key}")
                    else:
                        text_check(label, edge["label"], key+"::label")
                        x,y,bw,bh = edge["label"]["bboxPx"]
                        check_point(center(label), (x+bw/2,y+bh/2), key+"::label")
        # Release native proxies before shutting down the isolated application.
        actual_count = len(actual)
        actual.clear()
        all_shapes.clear()
        glue.clear()
    return dict(passed=package["passed"] and not issues, issues=issues, fonts=sorted(fonts), dynamicConnectorCount=connector_count,
                actualTaggedObjects=actual_count, expectedShapes=len(scene.shapes) if scene else None,
                expectedGroups=len(scene.groups) if scene else None, expectedConnectors=len(scene.connectors) if scene else None,
                note="Readback validates declared content; original-image completeness still needs visual review.")
