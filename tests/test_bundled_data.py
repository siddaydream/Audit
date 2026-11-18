import hashlib
import json
from pathlib import Path
import tempfile
import unittest

from src.bundled_data import restore_bundled_data


class BundledDataTests(unittest.TestCase):
    def prepare(self, root):
        bundle = root / 'bundled'
        bundle.mkdir()
        (bundle / 'a.part').write_bytes(b'abc')
        (bundle / 'b.part').write_bytes(b'def')
        (bundle / 'index.json').write_text(json.dumps([{
            'path': 'table.xlsx', 'parts': ['a.part', 'b.part'],
            'sha256': hashlib.sha256(b'abcdef').hexdigest(),
        }]))

    def test_restore_and_preserve_updated_file(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            self.prepare(root)
            restore_bundled_data(root)
            self.assertEqual((root / 'table.xlsx').read_bytes(), b'abcdef')
            (root / 'table.xlsx').write_bytes(b'new snapshot')
            restore_bundled_data(root)
            self.assertEqual((root / 'table.xlsx').read_bytes(), b'new snapshot')

    def test_corrupt_bundle_leaves_no_partial_workbook(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            self.prepare(root)
            (root / 'bundled/a.part').write_bytes(b'bad')
            with self.assertRaises(ValueError):
                restore_bundled_data(root)
            self.assertEqual(list(root.iterdir()), [root / 'bundled'])

    def test_missing_part_leaves_no_partial_workbook(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            self.prepare(root)
            (root / 'bundled/b.part').unlink()
            with self.assertRaises(FileNotFoundError):
                restore_bundled_data(root)
            self.assertEqual(list(root.iterdir()), [root / 'bundled'])
