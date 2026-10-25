import json
import re
import unittest
import tempfile
from pathlib import Path
from zipfile import ZipFile, BadZipFile

ROOT=Path(__file__).resolve().parents[1]

class DistributionTests(unittest.TestCase):
    def test_truncated_archive_is_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            dest=Path(folder)/'truncated.zip'
            dest.write_bytes((ROOT/'dist/image-to-editable-ppt-claude-alpha.zip').read_bytes()[:100])
            with self.assertRaises(BadZipFile):ZipFile(dest)

    def test_manifest_assets_and_skill(self):
        manifest=json.loads((ROOT/'plugin.json').read_text())
        self.assertEqual(manifest['$schema'],'https://agent-plugins.org/schemas/1.0.0/plugin.schema.json')
        self.assertEqual(manifest['name'],'image-to-editable-ppt')
        interface=manifest['extensions']['com.openai']['interface']
        for key in ['logo','composerIcon']:
            self.assertTrue((ROOT/interface[key]).is_file())
        self.assertNotIn('mcpServers',manifest)
        skill=(ROOT/'skills/image-to-editable-ppt/SKILL.md').read_text()
        description=re.search(r'^description: "(.*)"$',skill,re.M).group(1)
        self.assertLessEqual(len(description),200)
        self.assertIn('Native charts and data extraction are opt-in',skill)

    def test_archive_structure_and_native_binaries_excluded(self):
        for file,entry in [('image-to-editable-ppt-claude-alpha.zip','image-to-editable-ppt/skill.md'),('sliderecast-openai-plugin-alpha.zip','skills/image-to-editable-ppt/SKILL.md')]:
            with self.subTest(file=file):
                with ZipFile(ROOT/'dist'/file) as z:
                    names=z.namelist();self.assertIn(entry,names);self.assertIsNone(z.testzip())
                    self.assertFalse(any(n.endswith(('.node','.dll','.exe','.so','.ttf','.otf','.ttc')) for n in names))
                    self.assertTrue(any(n.endswith('THIRD-PARTY-NOTICES.txt') for n in names))
                    self.assertFalse(any('..' in Path(n).parts for n in names))
                    self.assertTrue(any(n.endswith('pptxgenjs.cjs') for n in names))
                    self.assertTrue(any(n.endswith('office-kit.mjs') for n in names))

if __name__=='__main__':unittest.main()
