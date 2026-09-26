"""Windows Visio adapter. Importing this module does not start or require Visio."""
from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime
import gc
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile

from .common import number, quoted, read_json
from .spec import expand
from .assets import resolve_asset

_com_cache = None


def cell(shape, name, formula):
    shape.CellsU(name).FormulaU = str(formula)


def value(shape, name):
    return float(shape.CellsU(name).ResultIU)


def tag(shape, identifier, kind=None):
    for name, text in (("SpecId", identifier), ("SpecKind", kind)):
        if text is None:
            continue
        if not shape.CellExistsU("User."+name, 0):
            shape.AddNamedRow(242, name, 0)
        cell(shape, "User."+name, quoted(text))


def identity(shape):
    return str(shape.CellsU("User.SpecId").ResultStr("")) if shape.CellExistsU("User.SpecId", 0) else ""


def walk(collection):
    for i in range(1, collection.Count+1):
        shape = collection.Item(i)
        yield shape
        yield from walk(shape.Shapes)


def page_point(shape, x, y):
    return tuple(shape.XYToPage(float(x), float(y)))


def center(shape):
    return page_point(shape, value(shape, "LocPinX"), value(shape, "LocPinY"))


def endpoint(shape, end):
    point = (value(shape, end+"X"), value(shape, end+"Y"))
    parent = shape.ContainingShape
    return page_point(parent, *point) if parent and parent.ID else point


def connection_points(shape):
    result = []
    if not shape.SectionExists(7, 0):
        return result
    for row in range(shape.RowCount(7)):
        x, y = shape.CellsSRC(7, row, 0), shape.CellsSRC(7, row, 1)
        result.append(dict(row=row, cell=x, point=page_point(shape, x.ResultIU, y.ResultIU)))
    return result


def cardinal(shape, direction, tolerance=.001):
    points = connection_points(shape)
    if not points:
        raise ValueError(f"{identity(shape) or shape.NameU} has no native connection points")
    axis = 0 if direction in {"left", "right"} else 1
    extreme = (max if direction in {"right", "top"} else min)(p["point"][axis] for p in points)
    c = center(shape)
    candidates = [p for p in points if abs(p["point"][axis]-extreme) <= tolerance]
    return min(candidates, key=lambda p: (abs(p["point"][1-axis]-c[1-axis]), abs(p["point"][axis]-extreme), p["row"]))["cell"]


def auto_pair(source, target):
    a, b = center(source), center(target)
    dx, dy = b[0]-a[0], b[1]-a[1]
    if abs(dx) >= abs(dy):
        return ("right", "left") if dx >= 0 else ("left", "right")
    return ("top", "bottom") if dy >= 0 else ("bottom", "top")


def color(text):
    if re.fullmatch(r"#[0-9a-fA-F]{6}", text):
        return "RGB("+",".join(str(int(text[i:i+2], 16)) for i in (1, 3, 5))+")"
    return text


@contextmanager
def application(visible=False):
    if sys.platform != "win32":
        raise RuntimeError("Native VSDX rendering/COM validation requires Windows desktop Visio")
    import pythoncom
    # A process-owned makepy cache avoids broken or concurrently rewritten global
    # gen_py packages. Never delete the user's shared Office automation cache.
    global _com_cache
    if _com_cache is None:
        import win32com
        import win32com.gen_py
        _com_cache = tempfile.TemporaryDirectory(prefix="visio-python-com-")
        win32com.__gen_path__ = _com_cache.name
        win32com.gen_py.__path__ = [_com_cache.name]
    from win32com.client import DispatchEx, gencache
    pythoncom.CoInitialize()
    app = None
    try:
        # Generated wrappers expose XYToPage/XYFromPage output parameters as tuples.
        app = gencache.EnsureDispatch(DispatchEx("Visio.Application"))
        app.Visible = visible
        yield app
    finally:
        if app is not None:
            try:
                for i in range(app.Documents.Count, 0, -1):
                    document = app.Documents.Item(i)
                    document.Saved = True
                    document.Close()
                app.Quit()
            finally:
                app = None
        gc.collect()
        pythoncom.CoUninitialize()


def stencil_path(name):
    roots = [os.environ.get("VISIO_STENCIL_ROOT")]
    roots += [str(Path(os.environ[k])/"Microsoft Office") for k in ("ProgramFiles", "ProgramFiles(x86)") if k in os.environ]
    for root in filter(None, roots):
        found = next(Path(root).rglob(name), None) if Path(root).exists() else None
        if found:
            return found.resolve()
    raise FileNotFoundError(f"Missing {name}; set VISIO_STENCIL_ROOT to the installed Visio Content directory")


class Renderer:
    def __init__(self, app, spec, directory, library=None, svg_color="#1E2A44"):
        self.app, self.spec, self.directory, self.library, self.svg_color = app, spec, directory, library, svg_color
        self.scene = expand(spec)
        self.w, self.h = spec["page"]["widthIn"], spec["page"]["heightIn"]
        self.rw, self.rh = spec["reference"]["widthPx"], spec["reference"]["heightPx"]
        self.doc = app.Documents.Add("")
        self.page = self.doc.Pages.Item(1)
        cell(self.page.PageSheet, "PageWidth", number(self.w)+" in")
        cell(self.page.PageSheet, "PageHeight", number(self.h)+" in")
        self.stencils, self.shapes, self.edges, self.ports, self.groups, self.labels = {}, {}, {}, {}, {}, {}

    def log(self, phase):
        print(phase, file=sys.stderr, flush=True)

    def style(self, item):
        return dict(self.spec.get("styles", {}).get(item.get("style"), {}), **item)

    def master(self, name, stencil="BASIC_M.VSSX"):
        if stencil not in self.stencils:
            self.stencils[stencil] = self.app.Documents.Open(str(stencil_path(stencil)))
        return self.stencils[stencil].Masters.ItemU(name)

    def point(self, x, y):
        return self.w*x/self.rw, self.h-self.h*y/self.rh

    def bounds(self, box):
        x, y, w, h = box
        left, top = self.point(x, y)
        right, bottom = self.point(x+w, y+h)
        return left, bottom, right, top

    def position(self, shape, bounds):
        left, bottom, right, top = bounds
        for name, v in dict(PinX=(left+right)/2, PinY=(bottom+top)/2, Width=right-left, Height=top-bottom).items():
            cell(shape, name, number(v)+" in")

    def apply_style(self, shape, item):
        style = self.style(item)
        fill, line = style.get("fill", "none"), style.get("line", "none")
        cell(shape, "FillPattern", 0 if fill == "none" else 1)
        if fill != "none":
            cell(shape, "FillForegnd", color(fill))
        cell(shape, "LinePattern", 0 if line == "none" else style.get("linePattern", 1))
        if line != "none":
            cell(shape, "LineColor", color(line))
            cell(shape, "LineWeight", number(style.get("lineWeightPt", .8))+" pt")
        if style.get("roundingPx"):
            cell(shape, "Rounding", number(style["roundingPx"]*self.w/self.rw)+" in")

    def text(self, shape, item):
        style = self.style(item)
        if not style.get("text"):
            return
        shape.Text = str(style["text"])
        cell(shape, "Char.Font", "FONT("+quoted(style.get("font", "Times New Roman"))+")")
        cell(shape, "Char.Size", number(style.get("fontSizePt", 10))+" pt")
        cell(shape, "Char.Color", color(style.get("fontColor", "#111111")))
        cell(shape, "Char.Style", int(bool(style.get("bold")))+2*int(bool(style.get("italic"))))
        cell(shape, "Para.HorzAlign", {"left":0, "center":1, "right":2}[style.get("align", "center")])
        cell(shape, "VerticalAlign", {"top":0, "middle":1, "bottom":2}[style.get("verticalAlign", "middle")])
        for margin in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
            cell(shape, margin, number(style.get("textMarginPt", 2))+" pt")

    def arrows(self, shape, item):
        arrow = item.get("arrow", "none")
        for end in ("Begin", "End"):
            cell(shape, end+"Arrow", item.get("arrowType", 4) if arrow in {end.lower(), "both"} else 0)
            if "arrowSize" in item:
                cell(shape, end+"ArrowSize", item["arrowSize"])

    def draw(self, item):
        key, kind = item["id"], item["kind"]
        bounds = self.bounds(item["bboxPx"])
        if kind in {"rect", "oval", "circle", "diamond", "native", "group"}:
            master = item.get("master") or {"rect":"Rectangle", "oval":"Ellipse", "circle":"Circle", "diamond":"Diamond", "group":"Rectangle"}.get(kind)
            if not master:
                raise ValueError(f"Native shape {key} needs a master")
            s = self.page.Drop(self.master(master, item.get("stencil", "BASIC_M.VSSX")), 0, 0)
            self.position(s, bounds)
        elif kind in {"text", "path"}:
            s = self.page.DrawRectangle(*bounds)
            if kind == "path":
                s.DeleteSection(10)
                s.AddSection(10)
                s.AddRow(10, 0, 137)
                points = item["points"] + ([item["points"][0]] if item.get("closed") else [])
                for index, (x, y) in enumerate(points):
                    row = s.AddRow(10, -2, 138 if index == 0 else 139)
                    s.CellsSRC(10, row, 0).FormulaU = "Width*"+number(x)
                    s.CellsSRC(10, row, 1).FormulaU = "Height*"+number(1-y)
                cell(s, "Geometry1.NoFill", 0 if item.get("closed") else 1)
        elif kind == "line":
            s = self.page.DrawLine(*bounds)
        elif kind == "svg-asset":
            source = resolve_asset(item["assetRef"], self.directory, self.library, self.svg_color)
            self.log("Import: "+key)
            s = self.page.Import(str(source))
            # Visio may convert SVG to native groups. Their geometry scales via
            # group-relative formulas, but absolute stroke weights do not.
            # Scale every imported stroke once using the original import size.
            stroke_scale = min((bounds[2]-bounds[0])/value(s,"Width"),
                               (bounds[3]-bounds[1])/value(s,"Height"))
            imported = [s, *walk(s.Shapes)]
            strokes = [(part, value(part,"LineWeight")) for part in imported]
            self.position(s, bounds)
            for part, weight in strokes:
                cell(part, "LineWeight", number(weight*stroke_scale)+" in")
        else:
            raise ValueError(kind)
        if kind != "svg-asset":
            self.apply_style(s, item)
            self.text(s, item)
        if item.get("angleDeg"):
            cell(s, "Angle", number(item["angleDeg"])+" deg")
        if kind in {"text", "group"}:
            cell(s, "ShapePermeableX", 1)
            cell(s, "ShapePermeableY", 1)
        if kind == "line":
            self.arrows(s, item)
        tag(s, key, kind)
        self.shapes[key] = s
        self.ports[key] = {}
        for p in item.get("ports", []):
            row = s.AddRow(7, -2, 153)
            s.CellsSRC(7, row, 0).FormulaU = "Width*"+number(p["x"])
            s.CellsSRC(7, row, 1).FormulaU = "Height*"+number(1-p["y"])
            s.CellsSRC(7, row, 2).FormulaU = "0"
            s.CellsSRC(7, row, 3).FormulaU = "0"
            self.ports[key][p["id"]] = s.CellsSRC(7, row, 0)
        return s

    def connection(self, key, name, tolerance=.001):
        s = self.shapes[key]
        if name.startswith("port:"):
            return self.ports[key][name[5:]]
        if name in {"left", "right", "top", "bottom"}:
            return cardinal(s, name, tolerance)
        full = name if name.startswith("Connections.") else "Connections."+name
        if not s.CellExistsU(full, 0):
            raise ValueError(f"Missing connection {key}/{name}")
        return s.CellsU(full)

    def draw_edge(self, item):
        s = self.page.Drop(self.master("Dynamic connector", "SSFLOW_M.VSSX"), 0, 0)
        directions = auto_pair(self.shapes[item["from"]], self.shapes[item["to"]])
        for end, prop, direction in zip(("Begin", "End"), ("from", "to"), directions):
            name = item.get(prop+"Connection", "auto")
            s.CellsU(end+"X").GlueTo(self.connection(item[prop], direction if name == "auto" else name, item.get("connectionToleranceIn", .001)))
        self.apply_style(s, {"line":"#111111", **self.style(item)})
        self.arrows(s, item)
        routing = item.get("routing", "auto")
        if routing in {"orthogonal", "straight"}:
            cell(s, "ShapeRouteStyle", {"orthogonal":1, "straight":2}[routing])
        cell(s, "ConLineJumpCode", {"default":0, "never":1, "always":2, "other":3, "neither":4}[item.get("lineJump", "never")])
        tag(s, item["id"], "connector")
        self.edges[item["id"]] = s
        if item.get("label"):
            label = dict(item["label"], id=item["id"]+"::label", kind="text")
            self.labels[item["id"]] = self.draw(label)

    def route(self, connector, points):
        cell(connector, "ConFixedCode", 2)
        cell(connector, "Geometry1.NoShow", 1)
        connector.AddSection(11)
        connector.AddRow(11, 0, 137)
        connector.AddRow(11, -2, 138)
        def end_formulas(row, end):
            # Dynamic connector masters can clamp Height to a minimum. Neither
            # (0,0) nor (Width,Height) then equals the glued endpoint. Invert
            # the connector's own transform from parent coordinates instead.
            dx, dy = f"({end}X-PinX)", f"({end}Y-PinY)"
            connector.CellsSRC(11,row,0).FormulaU = f"LocPinX+IF(FlipX,-1,1)*(COS(Angle)*{dx}+SIN(Angle)*{dy})"
            connector.CellsSRC(11,row,1).FormulaU = f"LocPinY+IF(FlipY,-1,1)*(-SIN(Angle)*{dx}+COS(Angle)*{dy})"
        end_formulas(1, "Begin")
        for p in points:
            x, y = connector.XYFromPage(*self.point(*p))
            row = connector.AddRow(11, -2, 139)
            connector.CellsSRC(11, row, 0).FormulaU = number(x)+" in"
            connector.CellsSRC(11, row, 1).FormulaU = number(y)+" in"
        row = connector.AddRow(11, -2, 139)
        end_formulas(row, "End")
        cell(connector, "Geometry2.NoFill", 1)
        cell(connector, "Geometry2.NoShow", 0)

    def build(self):
        for item in self.scene.shapes:
            self.draw(item)
            self.log("Shape: "+item["id"])
        for edge in self.scene.connectors:
            self.draw_edge(edge)
        self.log("Phase: grouping")
        for group in sorted(self.scene.groups, key=lambda g:-g["depth"]):
            selection = self.page.CreateSelection(0)
            selection.Select(self.shapes[group["id"]], 2)
            for child in self.scene.shapes:
                if child["parentId"] == group["id"]:
                    selection.Select(self.groups[child["id"]] if child["kind"] == "group" else self.shapes[child["id"]], 2)
            for edge in self.scene.connectors:
                if edge["parentId"] == group["id"]:
                    selection.Select(self.edges[edge["id"]], 2)
                    if edge["id"] in self.labels:
                        selection.Select(self.labels[edge["id"]], 2)
            result = selection.Group()
            tag(result, group["id"]+"::group", "group-wrapper")
            cell(result, "SelectMode", 1)
            self.groups[group["id"]] = result
        self.log("Phase: routes")
        for edge in self.scene.connectors:
            s = self.edges[edge["id"]]
            if edge.get("waypointsPx"):
                self.route(s, edge["waypointsPx"])
            s.BringToFront()
            if edge["id"] in self.labels:
                self.labels[edge["id"]].BringToFront()


def render(spec_path, output, preview=None, formats=(), library=None, svg_color="#1E2A44", overwrite=False, visible=False):
    spec_path, output = Path(spec_path).resolve(), Path(output).resolve()
    spec = read_json(spec_path)
    expand(spec)
    if output.suffix.lower() != ".vsdx":
        raise ValueError("Output must end with .vsdx")
    if output.exists():
        if not overwrite:
            raise FileExistsError(output)
        shutil.copy2(output, output.with_name(output.stem+".backup-"+datetime.now().strftime("%Y%m%d-%H%M%S-%f")+".vsdx"))
    output.parent.mkdir(parents=True, exist_ok=True)
    exports = {}
    with application(visible) as app:
        renderer = Renderer(app, spec, spec_path.parent, library, svg_color)
        renderer.build()
        renderer.doc.SaveAs(str(output))
        for fmt in dict.fromkeys((["png"] if preview else [])+list(formats)):
            target = Path(preview).resolve() if fmt == "png" and preview else output.with_suffix("."+fmt)
            target.parent.mkdir(parents=True, exist_ok=True)
            if fmt == "pdf":
                renderer.doc.ExportAsFixedFormat(1, str(target), 1, 0)
            elif fmt == "png":
                export_preview(renderer.doc, target)
            elif fmt == "svg":
                # Page.Export otherwise crops to content and destroys source registration.
                # Temporary full-page bounds are added after saving, never persisted.
                background = renderer.page.DrawRectangle(0, 0, renderer.w, renderer.h)
                cell(background, "LinePattern", 0)
                cell(background, "FillPattern", 1)
                cell(background, "FillForegnd", "RGB(255,255,255)")
                background.SendToBack()
                try:
                    renderer.page.Export(str(target))
                finally:
                    background.Delete()
            else:
                raise ValueError(f"Unsupported export: {fmt}")
            exports[fmt] = str(target)
        report = dict(passed=True, output=str(output), exports=exports, **{k:v for k,v in renderer.scene.report().items() if k != "passed"})
        del renderer
    return report


def export_preview(document, output):
    """Rasterize the actual full-page Visio PDF, independent of user export settings."""
    import pypdfium2 as pdfium
    with tempfile.TemporaryDirectory(prefix="visio-preview-") as directory:
        pdf_path = Path(directory)/"page.pdf"
        document.ExportAsFixedFormat(1, str(pdf_path), 1, 0)
        pdf = pdfium.PdfDocument(str(pdf_path))
        try:
            page = pdf[0]
            try:
                bitmap = page.render(scale=2)
                try:
                    bitmap.to_pil().save(output)
                finally:
                    bitmap.close()
            finally:
                page.close()
        finally:
            pdf.close()
