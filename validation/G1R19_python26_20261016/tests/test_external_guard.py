"""Windows checkpoint sharing and exact process identity, without killing jobs."""
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest import TestCase, main, mock
import importlib.util
import json
import threading

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('external_guard',ROOT/'guard_batch.py')
guard=importlib.util.module_from_spec(spec);spec.loader.exec_module(guard)


class ExternalGuardTests(TestCase):
    def test_reader_closes_old_handle_and_writer_retries_a_sharing_conflict(self):
        directory=ROOT/'results';directory.mkdir(exist_ok=True)
        with TemporaryDirectory(prefix='guard-test-',dir=directory) as folder:
            path=Path(folder)/'checkpoint.json'
            self.assertTrue(path.resolve().is_relative_to(directory.resolve()))
            path.write_text(json.dumps({'generation':0}),encoding='utf8')
            read_file=guard.win32file.ReadFile
            results=[];threads=[]
            def replace_while_reading(handle,size):
                thread=threading.Thread(target=lambda:results.append(guard.atomic(path,{'generation':1})))
                threads.append(thread);thread.start()
                return read_file(handle,size)
            with mock.patch.object(guard.win32file,'ReadFile',replace_while_reading):
                self.assertEqual(guard.read(path,None),{'generation':0})
            for thread in threads:thread.join(timeout=2)
            self.assertEqual(results,[True])
            self.assertEqual(guard.read(path,None),{'generation':1})

    def test_process_birth_and_full_command_must_both_match(self):
        process=SimpleNamespace(create_time=lambda:123.,cmdline=lambda:['python.exe','case.py','--case','D02A'])
        record=dict(pid=999999,create_time=123.,command=process.cmdline())
        with mock.patch.object(guard.psutil,'Process',return_value=process):
            self.assertIs(guard.exact(record),process)
            self.assertIsNone(guard.exact(dict(record,create_time=124.)))
            self.assertIsNone(guard.exact(dict(record,command=['python.exe','case.py','--case','D01A'])))

    def test_venv_redirector_compatibility_preserves_all_solver_arguments(self):
        expected=['D:/venv/python.exe','case.py','--case','D02A']
        self.assertTrue(guard.compatible(['C:/runtime/python.exe',*expected[1:]],expected))
        self.assertFalse(guard.compatible(['C:/runtime/python.exe','case.py','--case','D01A'],expected))


if __name__=='__main__':main()
