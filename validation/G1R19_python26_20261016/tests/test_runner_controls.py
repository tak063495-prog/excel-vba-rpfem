"""Exercise the real child-process watchdog without starting a numerical job."""
from pathlib import Path
import importlib.util
import os
import subprocess
import sys
import unittest
import psutil

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('runner26',ROOT/'run_26.py')
runner=importlib.util.module_from_spec(spec);spec.loader.exec_module(runner)

class WatchdogTests(unittest.TestCase):
    def child(self):
        command=[sys.executable,'-c','import time;time.sleep(30)']
        process=subprocess.Popen(command,creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
        def cleanup():
            if process.poll() is None:process.kill()
            process.wait(timeout=2)
        self.addCleanup(cleanup)
        return {'process':process,'command':command,'create_time':psutil.Process(process.pid).create_time()}

    def test_owned_child_stops(self):
        active=self.child();runner.stop_owned(active)
        self.assertIsNotNone(active['process'].poll())

    def test_creation_time_mismatch_does_not_stop_child(self):
        active=self.child();active['create_time']+=100
        with self.assertRaisesRegex(RuntimeError,'identity changed'):runner.stop_owned(active)
        self.assertIsNone(active['process'].poll())

    def test_command_mismatch_does_not_stop_child(self):
        active=self.child();active['command']=active['command']+['unexpected']
        with self.assertRaisesRegex(RuntimeError,'identity changed'):runner.stop_owned(active)
        self.assertIsNone(active['process'].poll())

if __name__=='__main__':unittest.main()
