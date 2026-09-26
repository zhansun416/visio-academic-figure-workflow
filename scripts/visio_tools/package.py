"""Read actual saved VSDX objects without Windows or Visio."""
from __future__ import annotations

from pathlib import Path, PurePosixPath
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile

NS = {"v":"http://schemas.microsoft.com/office/visio/2012/main", "r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


def inspect_package(path):
    objects, pages = [], []
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        def xml(name):
            return ET.fromstring(archive.read(name))
        def relationships(part):
            p = PurePosixPath(part)
            rel = str(p.parent/"_rels"/(p.name+".rels"))
            if rel not in names:
                return {}
            return {e.get("Id"):posixpath.normpath(str(p.parent)+"/"+e.get("Target", "")) for e in xml(rel) if e.get("TargetMode") != "External"}
        masters = {}
        if "visio/masters/masters.xml" in names:
            relations = relationships("visio/masters/masters.xml")
            for m in xml("visio/masters/masters.xml").findall("v:Master", NS):
                rel = m.find("v:Rel", NS)
                target = relations.get(rel.get("{"+NS["r"]+"}id")) if rel is not None else None
                if target in names:
                    root = xml(target)
                    masters[m.get("ID")] = (m.get("NameU", m.get("Name", "")), target, root)
        def user(shape, name):
            found = shape.find(f"v:Section[@N='User']/v:Row[@N='{name}']/v:Cell[@N='Value']", NS)
            return found.get("V", "") if found is not None else ""
        def visit(shape, part, page, parent="", inherited_master=None):
            key = user(shape, "SpecId") or shape.get("NameU") or f"page{page}:shape{shape.get('ID')}"
            kind_hint = user(shape, "SpecKind")
            master = masters.get(shape.get("Master")) or inherited_master
            inherited = None
            if master:
                wanted = shape.get("MasterShape")
                inherited = master[2].find(f".//v:Shape[@ID='{wanted}']", NS) if wanted else master[2].find("v:Shapes/v:Shape", NS)
            foreign, foreign_part = shape.find("v:ForeignData", NS), part
            if foreign is None and inherited is not None:
                foreign, foreign_part = inherited.find("v:ForeignData", NS), master[1]
            foreign_targets = []
            if foreign is not None:
                rels = relationships(foreign_part)
                foreign_targets = [rels.get(r.get("{"+NS["r"]+"}id"), "") for r in foreign.findall(".//v:Rel", NS)]
                if any(t.endswith(".svg") for t in foreign_targets):
                    kind = "picture-svg"
                elif any(re.search(r"\.(png|jpe?g|bmp|gif|tiff?)$", t, re.I) for t in foreign_targets):
                    kind = "picture-raster"
                else:
                    kind = "picture-unknown"
            elif shape.get("Type") == "Group":
                kind = "group"
            elif master and "connector" in master[0].lower():
                kind = "connector"
            elif shape.find("v:Cell[@N='BeginX']", NS) is not None:
                kind = "line"
            elif kind_hint == "text":
                kind = "text"
            else:
                kind = "shape"
            text_node = shape.find("v:Text", NS)
            record = dict(name=key, type=kind, page=page, shapeId=shape.get("ID"), parent=parent,
                          text="".join(text_node.itertext()) if text_node is not None else "", media=foreign_targets)
            objects.append(record)
            for child in shape.findall("v:Shapes/v:Shape", NS):
                visit(child, part, page, key, master)
        page_paths = sorted((n for n in names if re.fullmatch(r"visio/pages/page\d+\.xml", n)), key=lambda n:int(re.search(r"page(\d+)", n).group(1)))
        for index, part in enumerate(page_paths, 1):
            root = xml(part)
            pages.append(dict(page=index, part=part))
            for shape in root.findall("v:Shapes/v:Shape", NS):
                visit(shape, part, index)
        media = sorted(n for n in names if n.startswith("visio/media/"))
    return dict(passed=bool(pages and objects), file=str(Path(path).resolve()), pages=pages, objects=objects,
                media=media, rasterMedia=[m for m in media if re.search(r"\.(png|jpe?g|bmp|gif|tiff?)$", m, re.I)])
