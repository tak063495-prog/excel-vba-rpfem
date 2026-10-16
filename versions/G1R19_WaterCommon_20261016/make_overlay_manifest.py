"""Synchronize this isolated overlay and record exact source/import identities."""
from __future__ import annotations
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[1]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    entries = []
    for name in ("RPX_Mesh.bas", "RPX_Water.bas", "RPX_Guard.bas"):
        source = ROOT / "src/vba" / name
        text = source.read_text(encoding="utf-8").replace("\r\n", "\n").replace("\r", "\n")
        rawtext = text.replace("\n", "\r\n")
        source.write_bytes(rawtext.encode("utf-8"))
        imported = ROOT / "import" / name
        imported.write_bytes(rawtext.encode("cp932"))
        assert source.read_bytes().decode("utf-8") == imported.read_bytes().decode("cp932")
        baseline = REPO / "versions/G1R19/src/vba" / name
        entries.append(dict(module=name.removesuffix(".bas"), base=str(baseline.relative_to(REPO)),
                            base_sha256=sha(baseline), source=str(source.relative_to(ROOT)),
                            source_sha256=sha(source), imported=str(imported.relative_to(ROOT)),
                            import_sha256=sha(imported), decoded_text_equal=True))
    tracked = [ROOT / "tests/test_water_common_geometry.py", ROOT / "tests/test_water_common_source.py", ROOT / "tests/test_water_common_rollback.py",
               ROOT / "tests/WaterCommonSourceProbe.bas", ROOT / "tests/WaterCommonRollbackProbe.bas", ROOT / "README.md",
               REPO / "validation/G1R19_python26_water_20261016/python/waterline_conform.py",
               REPO / "validation/G1R19_python26_water_20261016/VBA_PROPOSAL.md",
               REPO / "docs/benchmarks/G1R17_native26_20261013/data/inputs/D08A.json",
               REPO / "versions/G1R17/results/geometry25/P02A_new_vba_mesh.bin"]
    data = dict(schema="rpfem.watercommon.overlay.v3", created_utc=datetime.now(timezone.utc).isoformat(),
                status="STATIC_ALGORITHM_TESTS_PASS; EXCEL_COMPILE_AND_RUNTIME_PENDING",
                scope="isolated prototype; no numerical bound or product 2% claim", modules=entries,
                references=[dict(path=str(p.relative_to(REPO)), sha256=sha(p)) for p in tracked],
                static_test_count=21, solver_runs=0, excel_runs=0,
                unchanged_policy=dict(q=8, raw_physical_tolerance=1e-7,
                                      own_root_relative_tolerance=1e-4, target_gap=0.02),
                results=dict(geometry_comparison="results/STATIC_GEOMETRY_COMPARISON.json",
                             fixture="results/D08A_expected_common_geometry.npz"))
    (ROOT / "SOURCE_MANIFEST.json").write_text(json.dumps(data, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")
    print(json.dumps({"status": data["status"], "source_import_pairs": len(entries),
                      "decoded_text_equal": all(x["decoded_text_equal"] for x in entries)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
