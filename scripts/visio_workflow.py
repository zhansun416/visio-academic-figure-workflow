#!/usr/bin/env python3
"""Portable Python command line; only render/validate/probe need Windows Visio."""
from __future__ import annotations

import argparse
from datetime import datetime
import importlib.util
import json
from pathlib import Path
import shutil
import subprocess
import sys

from visio_tools.common import read_json, write_json
from visio_tools.spec import expand
from visio_tools import assets


def parser():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--report", help="Write the JSON result atomically")
    sub = p.add_subparsers(dest="command", required=True)
    env = sub.add_parser("check-env")
    env.add_argument("--probe-visio", action="store_true")
    pre = sub.add_parser("preflight")
    pre.add_argument("spec")
    render = sub.add_parser("render")
    render.add_argument("spec")
    render.add_argument("output")
    render.add_argument("--preview")
    render.add_argument("--export", nargs="*", choices=["png","svg","pdf"], default=[])
    render.add_argument("--library")
    render.add_argument("--svg-color", default="#1E2A44")
    render.add_argument("--overwrite", action="store_true")
    render.add_argument("--visible", action="store_true")
    preview = sub.add_parser("preview", help="Reopen a saved VSDX and export its first page at 144 DPI")
    preview.add_argument("vsdx")
    preview.add_argument("output")
    validate = sub.add_parser("validate")
    validate.add_argument("vsdx")
    validate.add_argument("--spec")
    validate.add_argument("--expected-font")
    validate.add_argument("--min-font",type=float,default=0)
    validate.add_argument("--require-glued",action="store_true")
    validate.add_argument("--no-raster",action="store_true")
    validate.add_argument("--axis-aligned",action="store_true")
    validate.add_argument("--tolerance-px",type=float,default=1)
    validate.add_argument("--single-output",action="store_true")
    validate.add_argument("--allowed-acronyms",nargs="*")
    inspect = sub.add_parser("inspect")
    inspect.add_argument("vsdx")
    inspect.add_argument("--manifest")
    inspect.add_argument("--mode",choices=["draft","delivery"],default="draft")
    manifest = sub.add_parser("check-manifest")
    manifest.add_argument("manifest")
    manifest.add_argument("--mode",choices=["draft","delivery"],default="draft")
    svg = sub.add_parser("svg")
    svg.add_argument("--library")
    commands = svg.add_subparsers(dest="action",required=True)
    commands.add_parser("init")
    sync = commands.add_parser("sync")
    sync.add_argument("--repair",action="store_true")
    find = commands.add_parser("find")
    find.add_argument("query")
    find.add_argument("--limit",type=int,default=12)
    for name in ("inspect","prepare"):
        inspect_svg = commands.add_parser(name)
        inspect_svg.add_argument("source")
        inspect_svg.add_argument("--safe-dir")
        inspect_svg.add_argument("--complete",action="store_true")
        inspect_svg.add_argument("--color",default="#1E2A44")
    register = commands.add_parser("register")
    register.add_argument("source")
    register.add_argument("--source-url")
    register.add_argument("--source-label",default="user-supplied")
    register.add_argument("--license",default="unknown-review")
    register.add_argument("--category",default="general")
    register.add_argument("--tags",nargs="*")
    audit = commands.add_parser("audit")
    audit.add_argument("--public-release",action="store_true")
    candidates = commands.add_parser("candidates")
    candidates.add_argument("query")
    candidates.add_argument("output")
    candidates.add_argument("--limit",type=int,default=12)
    candidates.add_argument("--download",type=int,default=3)
    candidates.add_argument("--prefix")
    final = sub.add_parser("finalize")
    final.add_argument("vsdx")
    final.add_argument("archive")
    return p


def execute(args):
    if args.command == "check-env":
        modules = {m:importlib.util.find_spec(m) is not None for m in ("PIL","pypdfium2","win32com")}
        result = dict(passed=sys.version_info >= (3,10), python=sys.version.split()[0], platform=sys.platform,
                      modules=modules, visio="not probed", note="Preflight, SVG library and package/manifest inspection use the standard library; comparison needs Pillow; PNG previews need pypdfium2; native Visio operations need Windows and pywin32.")
        if args.probe_visio:
            from visio_tools.com_backend import application, stencil_path
            with application() as app:
                for filename, masters in {"BASIC_M.VSSX":["Rectangle","Ellipse","Circle","Diamond"],"SSFLOW_M.VSSX":["Dynamic connector"]}.items():
                    doc=app.Documents.Open(str(stencil_path(filename)))
                    for name in masters:
                        doc.Masters.ItemU(name)
            result["visio"]="COM and required masters available"
        return result
    if args.command == "preflight":
        return expand(read_json(args.spec)).report()
    if args.command == "render":
        from visio_tools.com_backend import render
        return render(args.spec,args.output,args.preview,args.export,args.library,args.svg_color,args.overwrite,args.visible)
    if args.command == "preview":
        from visio_tools.com_backend import application, export_preview
        source, target = Path(args.vsdx).resolve(), Path(args.output).resolve()
        if not source.is_file() or source == target or target.suffix.lower() != ".png":
            raise ValueError("Existing VSDX and a separate .png output are required")
        target.parent.mkdir(parents=True, exist_ok=True)
        with application() as app:
            doc = app.Documents.Open(str(source))
            export_preview(doc, target)
        return dict(passed=True, preview=str(target), page=1, dpi=144)
    if args.command == "validate":
        from visio_tools.validation import validate
        return validate(args.vsdx,args.spec,args.expected_font,args.min_font,args.require_glued,args.no_raster,args.axis_aligned,
                        tolerance_px=args.tolerance_px,single_output=args.single_output,allowed_acronyms=args.allowed_acronyms)
    if args.command == "inspect":
        from visio_tools.package import inspect_package
        report = inspect_package(args.vsdx)
        if args.manifest:
            from visio_tools.manifest import check
            report["manifestValidation"]=check(read_json(args.manifest),report,args.mode)
            report["passed"] &= report["manifestValidation"]["passed"]
        return report
    if args.command == "check-manifest":
        from visio_tools.manifest import check
        return check(read_json(args.manifest),mode=args.mode)
    if args.command == "svg":
        root=assets.library_root(args.library)
        if args.action == "init": return assets.initialize(root)
        if args.action == "sync": return assets.sync(root,args.repair)
        if args.action == "find": return assets.find(root,args.query,args.limit)
        if args.action in {"inspect","prepare"}:
            if Path(args.source).is_dir(): return assets.prepare(args.source,args.safe_dir,args.complete,args.color)
            report=assets.assess(args.source)
            return dict(report,passed=report["readyForVisioImport"])
        if args.action == "register":
            return assets.register(root,args.source,sourceUrl=args.source_url,sourceLabel=args.source_label,license=args.license,tags=args.tags,category=args.category)
        if args.action == "audit":
            report=assets.audit(root)
            if args.public_release: report["passed"]=report["publicReleaseReady"]
            return report
        if args.action == "candidates": return assets.candidates(args.query,args.output,args.limit,args.download,args.prefix)
    if args.command == "finalize":
        path,archive=Path(args.vsdx).resolve(),Path(args.archive).resolve()
        if not path.is_file() or path.suffix.lower() != ".vsdx" or archive == path.parent or archive.is_relative_to(path.parent):
            raise ValueError("Existing .vsdx and a separate archive directory outside its delivery directory are required")
        archive.mkdir(parents=True,exist_ok=True)
        moved=[]
        for other in path.parent.iterdir():
            if other.is_file() and other != path:
                target=archive/other.name
                if target.exists(): target=archive/(other.stem+"-"+datetime.now().strftime("%Y%m%d-%H%M%S-%f")+other.suffix)
                shutil.move(str(other),str(target))
                moved.append(other.name)
        return dict(passed=True,finalFile=str(path),archivedFiles=moved)
    raise ValueError("Unsupported command")


def main():
    args=parser().parse_args()
    try:
        result=execute(args)
    except Exception as exc:
        result=dict(passed=False,error=f"{type(exc).__name__}: {exc}")
    if args.report:
        write_json(args.report,result)
    print(json.dumps(result,ensure_ascii=False,indent=2,allow_nan=False))
    return 0 if result.get("passed",True) else 1


if __name__ == "__main__":
    raise SystemExit(main())
