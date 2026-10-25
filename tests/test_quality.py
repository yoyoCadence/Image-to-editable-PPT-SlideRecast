import json
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from io import BytesIO
import xml.etree.ElementTree as ET

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'skills/image-to-editable-ppt/scripts'))
from inspect_pptx import inspect
from compare import compare
from prepare_pdf import prepare

class PackageTests(unittest.TestCase):
    def test_both_engines_exact_text_and_native_objects(self):
        expected=json.loads((ROOT/'examples/source/expected-text.json').read_text())
        for engine in ['office-kit','pptxgenjs']:
            with self.subTest(engine=engine):
                report=inspect(ROOT/f'examples/results/{engine}.pptx',expected)
                self.assertTrue(report['valid_package_checks'],report['errors'])
                self.assertEqual(len(report['slides']),3)
                self.assertEqual(sum(p['native_text_objects'] for p in report['slides']),39)
                self.assertTrue(all(p['native_shapes']>p['native_text_objects'] for p in report['slides']))
                self.assertEqual(sum(p['pictures'] for p in report['slides']),0)
                self.assertEqual(len(report['charts']),0)  # default fidelity fixtures use shapes

    def test_invalid_relationship_detected(self):
        with tempfile.TemporaryDirectory() as folder:
            dest=Path(folder)/'bad.pptx'
            with ZipFile(ROOT/'examples/results/office-kit.pptx') as original,ZipFile(dest,'w',ZIP_DEFLATED) as output:
                for name in original.namelist():
                    if name!='ppt/slides/slide2.xml':output.writestr(name,original.read(name))
            report=inspect(dest)
            self.assertFalse(report['valid_package_checks'])
            self.assertTrue(any('Dangling relationship' in e for e in report['errors']))

    def test_wrong_text_and_wrong_page_order_fail(self):
        expected=json.loads((ROOT/'examples/source/expected-text.json').read_text())
        expected.reverse()
        report=inspect(ROOT/'examples/results/office-kit.pptx',expected)
        self.assertFalse(report['valid_package_checks'])
        self.assertFalse(report['expected_text'][0]['exact'])

    def test_chart_cache_and_embedded_workbook_edit(self):
        from openpyxl import load_workbook
        for engine in ['office-kit','pptxgenjs']:
            with self.subTest(engine=engine):
                file=ROOT/f'examples/results/{engine}-chart-edited.pptx';report=inspect(file)
                self.assertTrue(report['valid_package_checks'],report['errors'])
                self.assertEqual(report['charts'][0]['cached_values'],['33','59','77','93'])
                with ZipFile(file) as z:
                    workbooks=[n for n in z.namelist() if n.startswith('ppt/embeddings/') and n.endswith('.xlsx')]
                    self.assertEqual(len(workbooks),1)
                    book=load_workbook(BytesIO(z.read(workbooks[0])),data_only=True)
                    values=[c.value for row in book.active for c in row]
                    self.assertTrue(all(x in values for x in [33,59,77,93]),values)

    def test_edit_proof_retains_originals(self):
        proofs=json.loads((ROOT/'examples/results/edit-proof.json').read_text())
        for proof in proofs:
            report=inspect(ROOT/f'examples/results/{proof["engine"]}.pptx')
            self.assertEqual(report['sha256'],proof['original_sha256'])
            self.assertTrue(proof['original_retained'])
            edited=inspect(ROOT/f'examples/results/{proof["engine"]}-edited.pptx')
            self.assertIn('EDIT PROOF / 原生文字已修改',edited['slides'][0]['texts'])

    def test_png_embedding_with_overridden_dependency(self):
        report=inspect(ROOT/'examples/results/png-embedding.pptx')
        self.assertTrue(report['valid_package_checks'],report['errors'])
        self.assertEqual(report['slides'][0]['pictures'],1)
        self.assertFalse(report['slides'][0]['large_picture'])

class PixelAndPdfTests(unittest.TestCase):
    def test_identity_pixels_and_wrong_aspect(self):
        from PIL import Image
        with tempfile.TemporaryDirectory() as folder:
            p=Path(folder);Image.new('RGB',(120,80),'white').save(p/'a.png');Image.new('RGB',(80,120),'white').save(p/'b.png')
            report=compare(p/'a.png',p/'a.png',p/'out')
            self.assertEqual(report['normalized_rgb_mae'],0)
            with self.assertRaises(ValueError):compare(p/'a.png',p/'b.png',p/'wrong')

    def test_pdf_page_order_and_selectable_text(self):
        with tempfile.TemporaryDirectory() as folder:
            manifest=prepare(ROOT/'examples/source/selectable-text.pdf',folder)
            self.assertEqual(len(manifest),2)
            self.assertEqual((Path(folder)/manifest[0]['text']).read_text().strip(),'First page')
            self.assertEqual((Path(folder)/manifest[1]['text']).read_text().strip(),'Second page')
            self.assertTrue(all(p['has_extractable_text'] for p in manifest))

if __name__=='__main__':unittest.main()
