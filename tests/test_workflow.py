from copy import deepcopy
from pathlib import Path
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'scripts'))
from visio_tools.spec import expand
from visio_tools.manifest import check
from visio_tools.assets import assess, safe_copy, initialize, register, sync, find, audit
from visio_tools.package import inspect_package


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path=Path(self.temp.name)
        self.spec=json.loads((ROOT/'assets/templates/dense_figure_spec.json').read_text(encoding='utf-8-sig'))

    def test_nested_scene_and_manual_loops(self):
        scene=expand(self.spec)
        self.assertEqual((len(scene.groups),len(scene.connectors)),(5,15))
        self.assertEqual(scene.report()['manualRouteCount'],8)
        self.assertTrue(any(e['from']==e['to'] for e in scene.connectors))

    def test_duplicate_ids_rejected(self):
        self.spec['shapes'].append(deepcopy(self.spec['shapes'][0]))
        with self.assertRaisesRegex(ValueError,'duplicate'):expand(self.spec)

    def test_nonfinite_geometry_rejected(self):
        self.spec['shapes'][0]['bboxPx'][0]=float('nan')
        with self.assertRaisesRegex(ValueError,'finite'):expand(self.spec)

    def test_missing_port_rejected(self):
        self.spec['connectors'][0]['fromConnection']='port:missing'
        with self.assertRaisesRegex(ValueError,'Missing port'):expand(self.spec)

    def test_svg_cannot_be_endpoint(self):
        self.spec['shapes'].append({'id':'icon','kind':'svg-asset','bboxPx':[0,0,10,10],'assetRef':'a.svg'})
        self.spec['connectors'][0]['from']='icon'
        with self.assertRaisesRegex(ValueError,'native node'):expand(self.spec)

    def test_path_geometry_validated(self):
        self.spec['shapes'].append({'id':'wave','kind':'path','bboxPx':[0,0,20,10],'points':[[0,0],[.5,1],[1,0]]})
        self.assertIn('wave',expand(self.spec).by_id)
        self.spec['shapes'][-1]['points'][0]=[-1,0]
        with self.assertRaisesRegex(ValueError,'points'):expand(self.spec)

    def test_rich_svg_preserves_source_colors(self):
        source=ROOT/'assets/templates/rich_svg_fixture.svg'
        before=source.read_bytes()
        self.assertTrue(assess(source)['readyForVisioImport'])
        output=safe_copy(source,self.path/'safe.svg','#123456')
        self.assertIn('#123456',output.read_text())
        self.assertIn('#D00000',output.read_text())
        self.assertEqual(source.read_bytes(),before)
        with self.assertRaisesRegex(ValueError,'overwrite'):safe_copy(source,source)

    def test_svg_external_resources_rejected(self):
        p=self.path/'bad.svg'
        for body in ('<image href="https://example.invalid/a.png"/>','<script>alert(1)</script>','<path onclick="x()"/>','<use href="#missing"/>'):
            p.write_text('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1">'+body+'</svg>')
            self.assertFalse(assess(p)['readyForVisioImport'])

    def test_library_hash_corruption_is_not_silently_repaired(self):
        initialize(self.path)
        register(self.path,ROOT/'assets/templates/rich_svg_fixture.svg',license='MIT',sourceLabel='regression',sourceUrl='https://example.invalid/source',tags=['database'])
        self.assertEqual(len(find(self.path,'database')['matches']),1)
        self.assertTrue(sync(self.path)['passed'])
        icon=next((self.path/'icons').glob('*.svg'));icon.write_text(icon.read_text()+'\n')
        before=(self.path/'library.json').read_bytes()
        self.assertFalse(sync(self.path,repair=True)['passed'])
        self.assertEqual(before,(self.path/'library.json').read_bytes())

    def manifest(self):
        verified={'status':'verified','methods':['regional-crop'],'notes':'Test fixture only'}
        return {'schemaVersion':'1.0','pages':[{'page':1,'canvas':{'widthPx':100,'heightPx':100},'sourceReview':verified,'regions':[{'id':'region','bbox':{'x':0,'y':0,'w':100,'h':100},'critical':True,'fidelity':verified,'elements':[{'id':'element','bbox':{'x':0,'y':0,'w':100,'h':100},'expectedEditability':'native','representation':{'editability':'native','mode':'native-shape'},'completion':{'status':'complete'},'fidelity':verified,'objectMap':[{'name':'node','type':'shape'}]}]}]}]}

    def test_delivery_requires_source_review(self):
        m=self.manifest();m['pages'][0]['sourceReview']={'status':'review-needed'}
        self.assertTrue(check(m)['passed'])
        self.assertFalse(check(m,mode='delivery')['passed'])

    def test_delivery_checks_actual_types_and_unmapped_objects(self):
        m=self.manifest();package={'pages':[{'page':1}],'objects':[{'page':1,'name':'node','type':'picture-svg'},{'page':1,'name':'omitted','type':'shape'}]}
        result=check(m,package,mode='delivery')
        self.assertFalse(result['passed'])
        self.assertTrue(any('type mismatch' in s for s in result['issues']))
        self.assertTrue(any('Unmapped' in s for s in result['issues']))

    def test_picture_cannot_masquerade_as_native(self):
        m=self.manifest();e=m['pages'][0]['regions'][0]['elements'][0]
        e['representation']['mode']='mixed';e['objectMap'][0]['type']='picture-svg'
        self.assertFalse(check(m)['passed'])

    def test_deferred_content_blocks_delivery(self):
        m=self.manifest();e=m['pages'][0]['regions'][0]['elements'][0]
        e['completion']['status']='deferred';e['degradation']={'reason':'source unreadable','editabilityBoundary':'not built'}
        self.assertTrue(check(m)['passed']);self.assertFalse(check(m,mode='delivery')['passed'])

    def test_foreign_data_readback_uses_actual_media(self):
        p=self.path/'picture.vsdx'
        with zipfile.ZipFile(p,'w') as z:
            z.writestr('visio/pages/page1.xml','<PageContents xmlns="http://schemas.microsoft.com/office/visio/2012/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><Shapes><Shape ID="1" NameU="image"><ForeignData><Rel r:id="rId1"/></ForeignData></Shape></Shapes></PageContents>')
            z.writestr('visio/pages/_rels/page1.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="../media/image.svg"/></Relationships>')
            z.writestr('visio/media/image.svg','<svg/>')
        self.assertEqual(inspect_package(p)['objects'][0]['type'],'picture-svg')

    def test_comparison_refuses_aspect_distortion(self):
        from PIL import Image
        a,b=self.path/'a.png',self.path/'b.png'
        Image.new('RGB',(100,100),'white').save(a);Image.new('RGB',(150,100),'white').save(b)
        result=subprocess.run([sys.executable,str(ROOT/'scripts/compare_renders.py'),str(a),str(b),str(self.path/'comparison')],capture_output=True,text=True)
        self.assertNotEqual(result.returncode,0)
        report=json.loads((self.path/'comparison/comparison.json').read_text())
        self.assertEqual(report['status'],'aspect-ratio-mismatch')
        self.assertNotIn('fullPage',report)


if __name__=='__main__':unittest.main()
