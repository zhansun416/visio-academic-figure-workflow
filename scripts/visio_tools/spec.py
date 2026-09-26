from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
import math
import re

NATIVE = {"rect", "oval", "circle", "diamond", "native", "group"}
KINDS = NATIVE | {"text", "line", "path", "svg-asset"}


def numbers(values, count, label):
    if not isinstance(values, (list, tuple)) or len(values) != count or any(
        isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) for v in values
    ):
        raise ValueError(f"{label} requires {count} finite numbers")
    return list(values)


@dataclass
class Scene:
    shapes: list
    groups: list
    connectors: list
    by_id: dict
    warnings: list

    def report(self):
        return dict(passed=True, shapeCount=len(self.shapes), groupCount=len(self.groups),
                    connectorCount=len(self.connectors), textCount=sum(bool(s.get("text")) for s in self.shapes)
                    + sum(bool(e.get("label", {}).get("text")) for e in self.connectors),
                    manualRouteCount=sum(bool(e.get("waypointsPx")) for e in self.connectors), warnings=self.warnings)


def expand(spec):
    dimensions = numbers([spec.get("reference", {}).get("widthPx"), spec.get("reference", {}).get("heightPx"),
                          spec.get("page", {}).get("widthIn"), spec.get("page", {}).get("heightIn")], 4, "Canvas")
    if min(dimensions) <= 0:
        raise ValueError("Canvas dimensions must be positive")
    shapes, groups, edges, by_id, warnings = [], [], [], {}, []
    def unique(item):
        key = item.get("id")
        if not isinstance(key, str) or not key.strip() or "::" in key or key.casefold() in {s.casefold() for s in by_id}:
            raise ValueError(f"Missing, duplicate, or reserved id: {key!r}")
        return key
    def visit(items, offset=(0, 0), parent="", depth=0):
        for original in items:
            item = deepcopy(original)
            key = unique(item)
            kind = str(item.get("kind", "")).lower()
            if kind not in KINDS:
                raise ValueError(f"Unsupported kind {kind!r} for {key}")
            b = numbers(item.get("bboxPx"), 4, f"{key} bboxPx")
            if min(b[2:]) < 0 or (kind != "line" and min(b[2:]) == 0) or b[2:] == [0, 0]:
                raise ValueError(f"Invalid dimensions for {key}")
            if kind == "group" and item.get("angleDeg", 0):
                raise ValueError(f"Rotate individual children, not group {key}")
            if kind == "svg-asset" and not item.get("assetRef"):
                raise ValueError(f"Missing assetRef for {key}")
            if kind == "path":
                if len(item.get("points", [])) < 2:
                    raise ValueError(f"Path {key} needs at least two normalized points")
                for point in item["points"]:
                    xy = numbers(point, 2, f"{key} path point")
                    if min(xy) < 0 or max(xy) > 1:
                        raise ValueError(f"Path {key} points must lie in [0,1]")
            if "angleDeg" in item:
                numbers([item["angleDeg"]], 1, f"{key} angleDeg")
            if item.get("verticalAlign", "middle") not in {"top", "middle", "bottom"}:
                raise ValueError(f"Invalid verticalAlign for {key}")
            ports = set()
            for port in item.get("ports", []):
                name = port.get("id", "")
                if kind not in NATIVE or not re.fullmatch(r"[A-Za-z][A-Za-z0-9_-]*", name) or name in ports:
                    raise ValueError(f"Invalid port {key}/{name}")
                xy = numbers([port.get("x"), port.get("y")], 2, f"Port {key}/{name}")
                if min(xy) < 0 or max(xy) > 1:
                    raise ValueError(f"Port {key}/{name} must lie in [0,1]")
                ports.add(name)
            if parent and (min(b[:2]) < 0 or b[0]+b[2] > by_id[parent]["bboxPx"][2] or b[1]+b[3] > by_id[parent]["bboxPx"][3]):
                warnings.append(f"{key} protrudes outside {parent}; review against source")
            item.update(kind=kind, parentId=parent, bboxPx=[b[0]+offset[0], b[1]+offset[1], b[2], b[3]])
            by_id[key] = item
            shapes.append(item)
            if kind == "group":
                groups.append(dict(id=key, parentId=parent, depth=depth))
                visit(item.get("children", []), item["bboxPx"][:2], key, depth+1)
            elif item.get("children"):
                raise ValueError(f"Only groups have children: {key}")
    visit(spec.get("shapes", []))
    if not shapes:
        raise ValueError("At least one shape is required")
    for original in spec.get("connectors", []):
        edge = deepcopy(original)
        key = unique(edge)
        for end in ("from", "to"):
            node = by_id.get(edge.get(end), {})
            if node.get("kind") not in NATIVE:
                raise ValueError(f"{key} {end} must reference a native node or group")
            connection = edge.get(end+"Connection", "auto")
            if connection.startswith("port:"):
                if connection[5:] not in {p["id"] for p in node.get("ports", [])}:
                    raise ValueError(f"Missing port for {key}: {connection}")
            elif not re.fullmatch(r"auto|left|right|top|bottom|(?:Connections\.)?X[1-9]\d*", connection):
                raise ValueError(f"Invalid connection for {key}: {connection}")
        routing, points = edge.get("routing", "auto"), edge.get("waypointsPx", [])
        if routing not in {"auto", "manual", "straight", "orthogonal"}:
            raise ValueError(f"Invalid routing for {key}")
        if (routing == "manual" and not points) or (points and routing not in {"manual", "auto"}):
            raise ValueError(f"Conflicting or missing waypoints for {key}")
        for p in points:
            numbers(p, 2, f"{key} waypoint")
        if edge["from"] == edge["to"] and (not points or any(edge.get(e+"Connection", "auto") == "auto" for e in ("from", "to"))):
            raise ValueError(f"Self-loop {key} requires explicit ports and waypoints")
        if edge.get("arrow", "none") not in {"none", "begin", "end", "both"} or edge.get("lineJump", "never") not in {"default", "never", "always", "other", "neither"}:
            raise ValueError(f"Invalid arrow or lineJump for {key}")
        if "label" in edge:
            if min(numbers(edge["label"].get("bboxPx"), 4, f"{key} label")[2:]) <= 0:
                raise ValueError(f"Invalid label bbox for {key}")
        def ancestors(key):
            if by_id[key]["kind"] != "group":
                key = by_id[key]["parentId"]
            result = []
            while key:
                result.append(key)
                key = by_id[key]["parentId"]
            return result
        source_ancestors = set(ancestors(edge["from"]))
        edge["parentId"] = next((g for g in ancestors(edge["to"]) if g in source_ancestors), "")
        by_id[key] = edge
        edges.append(edge)
    return Scene(shapes, groups, edges, by_id, warnings)
