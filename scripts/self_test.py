#!/usr/bin/env python3
"""Portable tests plus optional native COM regression in isolated subprocesses."""
from pathlib import Path
import argparse
import json
import subprocess
import sys

from visio_tools.common import read_json, write_json

ROOT=Path(__file__).resolve().parents[1]


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output',required=True,type=Path)
    parser.add_argument('--visio',action='store_true')
    args=parser.parse_args();output=args.output.resolve();output.mkdir(parents=True,exist_ok=True)
    def run(name,cmd,expected=0):
        result=subprocess.run([sys.executable,*map(str,cmd)],cwd=ROOT,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=300)
        (output/(name+'.log')).write_text(result.stdout+'\n'+result.stderr,encoding='utf-8')
        if (result.returncode==0) != (expected==0):raise RuntimeError(f'{name} unexpected exit {result.returncode}; see log')
        print(name+': passed',flush=True)
    run('portable',['-m','unittest','discover','-s','tests','-v'])
    if args.visio:
        cli=ROOT/'scripts/visio_workflow.py'
        dense=ROOT/'assets/templates/dense_figure_spec.json'
        run('dense-render',[cli,'render',dense,output/'dense.vsdx','--preview',output/'dense.png','--overwrite'])
        run('dense-readback',[cli,'validate',output/'dense.vsdx','--spec',dense,'--require-glued','--no-raster'])
        changed=read_json(dense)
        changed['connectors'][0]['to']=changed['connectors'][0]['from']
        changed['connectors'][0].update(fromConnection='right',toConnection='left',routing='manual',waypointsPx=[[10,10]])
        write_json(output/'wrong-target.json',changed)
        run('wrong-target-rejected',[cli,'validate',output/'dense.vsdx','--spec',output/'wrong-target.json'],expected=1)
        rich={'reference':{'widthPx':400,'heightPx':200},'page':{'widthIn':6,'heightIn':3},'shapes':[
            {'id':'outer','kind':'group','bboxPx':[10,10,380,180],'line':'#333333','children':[
                {'id':'inner','kind':'group','bboxPx':[10,10,180,140],'line':'#999999','children':[
                    {'id':'rotated','kind':'rect','bboxPx':[10,40,40,40],'angleDeg':90,'ports':[{'id':'out','x':.5,'y':1}]},
                    {'id':'small-icon','kind':'svg-asset','bboxPx':[90,25,36,36],'assetRef':str(ROOT/'assets/templates/rich_svg_fixture.svg')} ]},
                {'id':'target','kind':'rect','bboxPx':[280,60,50,40]},
                {'id':'large-icon','kind':'svg-asset','bboxPx':[200,115,60,60],'assetRef':str(ROOT/'assets/templates/rich_svg_fixture.svg')}
            ]}], 'connectors':[{'id':'across','from':'rotated','to':'target','fromConnection':'port:out','toConnection':'left','routing':'manual','waypointsPx':[[90,80],[90,100],[270,100]],'arrow':'end'}]}
        write_json(output/'rich.json',rich)
        run('rich-render',[cli,'render',output/'rich.json',output/'rich.vsdx','--preview',output/'rich.png','--overwrite'])
        run('rich-readback',[cli,'validate',output/'rich.vsdx','--spec',output/'rich.json','--require-glued','--no-raster'])
        run('rich-strokes-and-move',[ROOT/'tests/visio_native_checks.py',output/'rich.vsdx',output/'moved.vsdx'])
    write_json(output/'self-test.json',{'passed':True,'visioTested':args.visio})
    return 0


if __name__=='__main__':raise SystemExit(main())
