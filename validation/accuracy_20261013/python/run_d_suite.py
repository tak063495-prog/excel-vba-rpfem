"""Two-at-a-time D case experiments; every case has its own files and log."""
from pathlib import Path
import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
import subprocess
import sys
import time


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--parallel", type=int, default=2)
    parser.add_argument("--cases", nargs="+", default=["D01A", "D02A", "D03N", "D07A"])
    args = parser.parse_args()
    if not 1 <= args.parallel <= 2:
        raise ValueError("This precision-first D runner permits at most two simultaneous cases")
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    driver = Path(__file__).with_name("accuracy_driver.py")
    repo = driver.parents[3]
    pending = list(args.cases)
    active, records = {}, []

    def snapshot():
        path = output / "d_suite_progress.json"
        staged = path.with_suffix(".json.tmp")
        staged.write_text(json.dumps(records, indent=2) + "\n", encoding="utf-8")
        staged.replace(path)

    while pending or active:
        while pending and len(active) < args.parallel:
            case = pending.pop(0)
            log_path = output / (case + ".log")
            if log_path.exists() or (output / case).exists():
                raise FileExistsError("Refusing to overwrite existing experiment: " + case)
            command = [sys.executable, "-X", "utf8", str(driver), "--case", case,
                       "--output", str(output), "--cycles", "8", "--max-elements", "12000", "--fan-lower"]
            log = log_path.open("x", encoding="utf-8")
            source_hashes = {p.name: hashlib.sha256(p.read_bytes()).hexdigest()
                             for p in [driver, driver.with_name("experiment_core.py"), driver.with_name("reduction.py")]}
            process = subprocess.Popen(command, cwd=repo, stdout=log, stderr=subprocess.STDOUT,
                          creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)
            record = {"case": case, "state": "running", "pid": process.pid,
                      "start_utc": datetime.now(timezone.utc).isoformat(), "command": command,
                      "log": str(log_path), "source_sha256": source_hashes}
            records.append(record)
            active[case] = process, log, record, time.perf_counter()
            snapshot()
            print("D_START", case, process.pid, flush=True)
        finished = []
        for case, (process, log, record, started) in active.items():
            code = process.poll()
            if code is None:
                continue
            log.close()
            summary = output / case / "summary.json"
            result = json.loads(summary.read_text(encoding="utf-8")) if summary.exists() else {}
            record.update(state="finished", returncode=code, seconds=time.perf_counter() - started,
                          finish_utc=datetime.now(timezone.utc).isoformat(), summary=str(summary),
                          target_met=result.get("target_met", False),
                          width=result.get("relative_width_symmetric"),
                          lower=result.get("lower", {}).get("fs"), upper=result.get("upper", {}).get("fs"))
            snapshot()
            print("D_END", case, code, record["target_met"], record["width"], flush=True)
            finished.append(case)
        for case in finished:
            active.pop(case)
        if active and not finished:
            time.sleep(1.)
    print("D_SUITE_DONE", flush=True)


if __name__ == "__main__":
    main()
