import contextlib
import io
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import BOS_PROJE_OLUSTUR as scaffold


class ScaffoldTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.target = self.root / 'bos proje'

    def test_every_manifest_file_is_created_with_zero_bytes(self):
        created, skipped = scaffold.create_empty_project(self.target)
        self.assertEqual((len(scaffold.FILES),0),(created,skipped))
        actual = {p.relative_to(self.target).as_posix() for p in self.target.rglob('*') if p.is_file()}
        self.assertEqual(set(scaffold.FILES),actual)
        self.assertTrue(all(p.stat().st_size == 0 for p in self.target.rglob('*') if p.is_file()))
        self.assertIn('banka_paneli/app.js',actual)
        self.assertIn('ing_bank/ing_excel_guncelle.py',actual)

    def test_existing_content_and_file_times_are_preserved(self):
        self.target.mkdir()
        existing = self.target/'panel.py'
        existing.write_bytes(b'user content\x00\xff')
        before = existing.stat().st_mtime_ns
        created, skipped = scaffold.create_empty_project(self.target)
        self.assertEqual((len(scaffold.FILES)-1,1),(created,skipped))
        self.assertEqual(b'user content\x00\xff',existing.read_bytes())
        self.assertEqual(before,existing.stat().st_mtime_ns)

    def test_second_execution_does_not_modify_existing_files(self):
        scaffold.create_empty_project(self.target)
        self.assertEqual((0,len(scaffold.FILES)),scaffold.create_empty_project(self.target))

    def test_file_in_place_of_directory_stops_before_creating_anything(self):
        self.target.mkdir()
        (self.target/'akbank').write_bytes(b'keep')
        with self.assertRaises(ValueError): scaffold.create_empty_project(self.target)
        self.assertEqual(['akbank'],[p.name for p in self.target.iterdir()])
        self.assertEqual(b'keep',(self.target/'akbank').read_bytes())

    def test_directory_in_place_of_file_stops_before_creating_anything(self):
        (self.target/'panel.py').mkdir(parents=True)
        with self.assertRaises(ValueError): scaffold.create_empty_project(self.target)
        self.assertEqual(['panel.py'],[p.name for p in self.target.iterdir()])

    def test_target_cannot_be_an_existing_file(self):
        self.target.write_bytes(b'keep')
        with self.assertRaises(ValueError): scaffold.create_empty_project(self.target)
        self.assertEqual(b'keep',self.target.read_bytes())

    def test_manifest_cannot_escape_target(self):
        for name in ('../outside.py','/outside.py','C:/outside.py','bank/../../outside.py',
                     'bank\\outside.py','bank//file.py','./file.py','',None):
            with self.subTest(name=name), self.assertRaises(ValueError):
                scaffold.create_empty_project(self.target,(name,))
        self.assertFalse(self.target.exists())

    def test_duplicate_case_and_file_directory_conflicts_are_rejected(self):
        for paths in (('panel.py','PANEL.PY'),('bank','bank/parser.py')):
            with self.subTest(paths=paths), self.assertRaises(ValueError):
                scaffold.create_empty_project(self.target,paths)
        self.assertFalse(self.target.exists())

    def test_links_are_rejected_before_creating_anything(self):
        with patch.object(scaffold,'_is_link',side_effect=lambda p:p.name=='akbank'), \
             self.assertRaises(ValueError):
            scaffold.create_empty_project(self.target)
        self.assertFalse(self.target.exists())

    def test_filesystem_root_is_not_a_target(self):
        with self.assertRaises(ValueError): scaffold.create_empty_project(Path(self.root.anchor))

    def test_list_mode_has_no_writes(self):
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            self.assertEqual(0,scaffold.main(['--hedef',str(self.target),'--listele']))
        self.assertIn('panel.py',output.getvalue())
        self.assertFalse(self.target.exists())

    def test_can_run_from_only_one_downloaded_file_without_project_or_packages(self):
        single = self.root/'BOS_PROJE_OLUSTUR.py'
        shutil.copy2(Path(scaffold.__file__),single)
        completed = subprocess.run([sys.executable,'-B',str(single),'--hedef',str(self.target)],
                                   cwd=self.root,capture_output=True,text=True,check=False)
        self.assertEqual(0,completed.returncode,completed.stderr+completed.stdout)
        self.assertEqual(len(scaffold.FILES),len([p for p in self.target.rglob('*') if p.is_file()]))
        self.assertEqual(Path(scaffold.__file__).read_bytes(),single.read_bytes())

    def test_default_target_is_next_to_downloaded_script(self):
        single = self.root/'BOS_PROJE_OLUSTUR.py'
        shutil.copy2(Path(scaffold.__file__),single)
        completed = subprocess.run([sys.executable,'-B',str(single)],cwd=self.root,
                                   capture_output=True,text=True,check=False)
        self.assertEqual(0,completed.returncode,completed.stderr)
        self.assertTrue((self.root/'peer-analysis-mini-bos'/'panel.py').is_file())
        self.assertGreater(single.stat().st_size,0)

    def test_manifest_contains_current_source_files_not_private_inputs(self):
        project = Path(scaffold.__file__).resolve().parent
        excluded = {'.git','.venv','venv','__pycache__','node_modules','data','outputs',
                    'kutuphaneler','.cache','.codex','.agents','.vscode','.idea'}
        source_suffixes = {'.py','.json','.md','.js','.html','.txt','.cmd'}
        expected = {path.relative_to(project).as_posix() for path in project.rglob('*')
                    if path.is_file() and not (set(path.relative_to(project).parts) & excluded)
                    and (path.suffix in source_suffixes or path.name in {'.gitignore','.gitattributes'})}
        self.assertEqual(expected,set(scaffold.FILES))
        self.assertFalse(any(name.endswith(('.exe','.xlsx','.pdf','.zip','.pyc')) for name in scaffold.FILES))


if __name__ == '__main__':
    unittest.main()
