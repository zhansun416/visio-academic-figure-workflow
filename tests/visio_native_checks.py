"""Regression for imported stroke overflow and nested group/connector movement."""
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
from visio_tools.com_backend import application,walk,identity,value,cell,center,endpoint

with application() as app:
    doc=app.Documents.Open(str(Path(sys.argv[1]).resolve()))
    page=doc.Pages.Item(1)
    objects={identity(s):s for s in walk(page.Shapes) if identity(s)}
    icon=objects['small-icon']
    widths=[value(s,'LineWeight') for s in walk(icon.Shapes) if value(s,'LinePattern')>0]
    assert widths and max(widths)<.05, f'Imported stroke failed to shrink with 36px icon: {widths}'
    before={k:center(objects[k]) for k in ['rotated','small-icon','target']}
    edge=objects['across'];before_edge=[endpoint(edge,e) for e in ['Begin','End']]
    container=objects['outer::group']
    cell(container,'PinX',str(value(container,'PinX')+.5)+' in')
    for k,p in before.items():
        q=center(objects[k]);assert abs(q[0]-p[0]-.5)<1e-6 and abs(q[1]-p[1])<1e-6,k
    for e,p in zip(['Begin','End'],before_edge):
        q=endpoint(edge,e);assert abs(q[0]-p[0]-.5)<1e-6 and abs(q[1]-p[1])<1e-6,e
    assert edge.Connects.Count==2
    doc.SaveAs(str(Path(sys.argv[2]).resolve()))
print('SVG stroke bound and grouped movement verified')
