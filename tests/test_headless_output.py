import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile
from datetime import datetime
from hashlib import sha256
from uuid import UUID

from openpyxl import Workbook
from bom_compare_core import CompareError, ComparisonRequest, build_output_paths, compare_workbooks

ROOT = Path(__file__).resolve().parents[1]
class FrozenDateTime(datetime):
    @classmethod
    def now(cls, tz=None):
        return cls(2026, 1, 1, 9, 30, 0, tzinfo=tz)


def write_book(path, only_sku, shared_qty, sheet='BOM'):
    path.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    book.active.title = sheet
    book.active.append(('Part Number', 'Description', 'Qty'))
    book.active.append(('SHARED', 'Standard', shared_qty))
    book.active.append((only_sku, 'exclusive', 2))
    book.save(path)
    book.close()


def workbook_xml(path):
    with zipfile.ZipFile(path) as artifact:
        return artifact.read('xl/worksheets/sheet1.xml')


class HeadlessOutputTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.left = self.root / 'left' / 'assembly.xlsx'
        self.right = self.root / 'right' / 'assembly.xlsx'
        self.output = self.root / 'output'
        write_book(self.left, 'ONLY-ONE', 3)
        write_book(self.right, 'ONLY-TWO', 4)
        self.request = ComparisonRequest(self.left, self.right, self.output)

    def hashes(self):
        return [sha256(p.read_bytes()).hexdigest() for p in (self.left, self.right)]

    def test_same_basename_keeps_two_sides_and_repeated_runs(self):
        before = self.hashes()
        with patch('bom_compare_core.datetime', FrozenDateTime):
            first = compare_workbooks(self.request)
            second = compare_workbooks(self.request)
        outputs = (first.left_output_path, first.right_output_path, second.left_output_path, second.right_output_path)
        self.assertEqual(len(set(outputs)), 4)
        self.assertEqual(set(self.output.glob('*.xlsx')), set(outputs))
        for left, right in [(outputs[0], outputs[1]), (outputs[2], outputs[3])]:
            self.assertTrue(left.name.endswith('-left.xlsx'))
            self.assertTrue(right.name.endswith('-right.xlsx'))
            self.assertIn(b'ONLY-ONE', workbook_xml(left))
            self.assertNotIn(b'ONLY-TWO', workbook_xml(left))
            self.assertIn(b'ONLY-TWO', workbook_xml(right))
            self.assertNotIn(b'ONLY-ONE', workbook_xml(right))
        self.assertEqual(self.hashes(), before)

    def test_second_save_failure_cleans_only_the_new_pair(self):
        before = self.hashes()
        self.output.mkdir()
        prior = self.output / 'previous.xlsx'
        prior.write_bytes(b'previous output')
        original_save = Workbook.save
        calls = 0
        def fail_second(book, file):
            nonlocal calls
            calls += 1
            if calls == 2:
                raise OSError('fixture interrupted write')
            original_save(book, file)
        with patch.object(Workbook, 'save', fail_second):
            with self.assertRaisesRegex(OSError, 'fixture interrupted write'):
                compare_workbooks(self.request)
        self.assertEqual(calls, 2)
        self.assertEqual(list(self.output.iterdir()), [prior])
        self.assertEqual(prior.read_bytes(), b'previous output')
        self.assertEqual(self.hashes(), before)

    def test_existing_second_path_does_not_get_overwritten(self):
        self.output.mkdir()
        with patch('bom_compare_core.datetime', FrozenDateTime), patch('bom_compare_core.uuid4') as run_id:
            run_id.return_value = UUID(int=1)
            left, right = build_output_paths(left_path=self.left, right_path=self.right, output_dir=self.output)
            right.write_bytes(b'previous output')
            with self.assertRaisesRegex(CompareError, '结果路径已存在'):
                compare_workbooks(self.request)
        self.assertFalse(left.exists())
        self.assertEqual(right.read_bytes(), b'previous output')

    def test_cli_operates_with_gui_module_unavailable(self):
        guard = '''import importlib.abc, runpy, sys
class Headless(importlib.abc.MetaPathFinder):
    def find_spec(self,fullname,*args):
        if fullname in {"bom_compare_gui", "tkinter", "_tkinter"}:
            raise AssertionError("CLI requested a GUI module")
sys.meta_path.insert(0,Headless())
sys.argv=sys.argv[1:]
runpy.run_path(sys.argv[0],run_name="__main__")
'''
        run = subprocess.run([sys.executable,'-B','-c',guard,str(ROOT/'compare_bom.py'),str(self.left),str(self.right),'--output-dir',str(self.output)],
                             cwd=ROOT,env={**os.environ,'PYTHONUTF8':'1'},capture_output=True,text=True,encoding='utf-8',timeout=12)
        self.assertEqual(run.returncode,0,run.stderr)
        self.assertIn('左侧独有 1 行，右侧独有 1 行，逐字段差异 1 行',run.stdout)
        self.assertEqual(len(list(self.output.glob('*.xlsx'))),2)


if __name__ == '__main__':
    unittest.main()
