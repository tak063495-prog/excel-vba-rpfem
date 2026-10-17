"""Replay historical D05A physical fields and two full original dual vectors.

No optimizer or Excel runs.  PASS means that the archived pair and the expected
raw-dual acceptance AND rejection were reproduced.  It does not mean that both
dual vectors pass, nor that current VBA or a new cold driver completed.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import importlib.metadata
import json
import math
import os
from pathlib import Path, PurePosixPath
import platform
import struct
import sys
import time

sys.dont_write_bytecode = True
for _name in ("OPENBLAS_NUM_THREADS", "OMP_NUM_THREADS", "MKL_NUM_THREADS", "NUMEXPR_NUM_THREADS"):
    os.environ[_name] = "1"
ROOT = Path(__file__).resolve().parent


def require(condition, message):
    if not condition:
        raise ValueError(message)


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read(path):
    def pairs(items):
        result = {}
        for key, value in items:
            require(key not in result, "Duplicate JSON key: " + key)
            result[key] = value
        return result

    def reject(value):
        raise ValueError("Nonfinite JSON constant: " + value)
    return json.loads(Path(path).read_text(encoding="utf8"), object_pairs_hook=pairs, parse_constant=reject)


def inside(name):
    require(isinstance(name, str) and "\\" not in name, "A relative POSIX bundle path is required")
    pure = PurePosixPath(name)
    require(not pure.is_absolute() and ".." not in pure.parts, "Unsafe bundle path")
    path = (ROOT / name).resolve()
    require(path.is_relative_to(ROOT), "Bundle path escaped its directory")
    return path


def recorded_relative(name):
    name = name.replace("\\", "/")
    prefix = "D:/Work/excel-vba-rpfem/"
    if name.startswith(prefix):
        name = name[len(prefix):]
    pure = PurePosixPath(name)
    require(not pure.is_absolute() and ":" not in name and ".." not in pure.parts,
            "Unknown historical path prefix")
    return pure.as_posix()


def root_contract(metadata, mesh, mode, fs):
    identity = dict(case="D05A", mode=mode, q=8, mesh_id=mesh.mesh_id,
                    physical_model_id=mesh.physical_model_id, policy=mesh.provenance["psi_policy"])
    root = metadata.get("audit", {}).get("discrete_root_search", {})
    require(all(root.get(k) == v for k, v in identity.items()), "Recorded root identity mismatch")
    require(all(metadata.get(k) == v for k, v in identity.items() if k != "policy"),
            "Certificate identity mismatch")
    bracket = root.get("bracket")
    require(isinstance(bracket, list) and len(bracket) == 2 and
            all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) for v in bracket) and
            0 < bracket[0] < bracket[1], "Invalid recorded root bracket")
    width = (bracket[1] - bracket[0]) / max(1., sum(bracket) / 2)
    require(width <= 1e-4 and root.get("root_complete") is True and metadata.get("root_complete") is True and
            root.get("bracket_signs") == [1, -1] and root.get("endpoint") == fs ==
            bracket[0 if mode == "lower" else 1], "Recorded root signs/endpoint/width mismatch")
    return dict(bracket=bracket, normalized_width=width, recorded_contract_pass=True,
                opposite_sign_regenerated_from_full_vector=False)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True, help="A new JSON file outside historical/ and source files")
    args = parser.parse_args()
    output = args.output.resolve()
    require(not output.exists() and output.parent.is_dir(), "Use a new output with an existing parent")
    require(not output.is_relative_to(ROOT / "historical"), "Do not write into immutable historical evidence")
    require(output.suffix.lower() == ".json", "JSON output required")
    started = time.perf_counter()
    manifest = read(ROOT / "HISTORICAL_MANIFEST.json")
    require(manifest.get("schema") == "issue28_d05a_historical_v1", "Unexpected manifest schema")
    references = read(ROOT / "historical/program_refs/PROGRAM_REFERENCE_MANIFEST.json")
    require(references.get("schema") == "issue28_historical_program_references_v1" and
            references["original_historical_manifest_sha256"] == sha(ROOT / "HISTORICAL_MANIFEST.json"),
            "Original historical manifest differs from the portable audit binding")
    reference_before = {}
    for name, item in references["files"].items():
        path = inside(name)
        require(path.is_file() and path.stat().st_size == item["bytes"] and sha(path) == item["sha256"],
                "Portable audit/reference bytes differ: " + name)
        reference_before[name] = item["sha256"]

    def original_file(name):
        # The old manifest is immutable and binds the ORIGINAL verifier.  The
        # current portable adapter has its own, separate reference manifest.
        return inside("historical/program_refs/ORIGINAL_VERIFY_HISTORICAL.py" if name == "VERIFY_HISTORICAL.py" else name)

    before = {}
    for name, item in manifest["files"].items():
        path = original_file(name)
        require(path.is_file() and path.stat().st_size == item["bytes"] and sha(path) == item["sha256"],
                "Evidence bytes differ: " + name)
        before[name] = item["sha256"]
    manifest_sha = sha(ROOT / "HISTORICAL_MANIFEST.json")
    source_record = read(ROOT / "historical/pair/source_manifest.json")
    for name, digest in source_record["source_sha256"].items():
        relative = recorded_relative(name)
        require(relative in manifest["read_only_relocation"], "Historical source missing from focused bundle")
        path = inside(manifest["read_only_relocation"][relative])
        require(sha(path) == digest, "Original executed source differs: " + relative)

    sys.path.insert(0, str(ROOT / "historical/python_snapshot"))
    import experiment_core as core
    import benchmark_models as models
    from portable_program_audit import compare_program, program_digest, unpack_program
    np = core.np
    calls = []

    def forbidden(*unused, **ignored):
        calls.append("clarabel.DefaultSolver")
        raise RuntimeError("NEW_OPTIMIZATION_FORBIDDEN_IN_HISTORICAL_REPLAY")

    core.clarabel.DefaultSolver = forbidden
    models._MANIFEST = read(ROOT / "historical/inputs/INPUTS_26.json")
    models.install_model_builder()
    record = read(ROOT / "historical/inputs/D05A.json")
    models._validate_record(record)
    require(record["case"] == "D05A" and not record.get("water", {}).get("coupling"), "Unexpected canonical case")

    def saved_mesh(path):
        with np.load(path, allow_pickle=False) as data:
            return core.mesh_from_arrays(data["xy"].copy(), data["tri"].copy(), data["material_id"].copy(), record)

    pair = ROOT / "historical/pair"
    summary = read(pair / "summary.json")
    sides, meshes = {}, {}
    for mode in ("lower", "upper"):
        metadata = read(pair / ("best_" + mode + "_certificate.json"))
        require(metadata == summary[mode], "Summary and certificate differ")
        require(metadata["mesh_path"] == "best_" + mode + "_mesh.npz" and
                metadata["field_path"] == "best_" + mode + "_field.npz", "Unexpected pair artifact names")
        mesh = saved_mesh(pair / metadata["mesh_path"])
        with np.load(pair / metadata["field_path"], allow_pickle=False) as data:
            require(str(data["mode"]) == mode and int(data["q"]) == 8, "Field mode/q mismatch")
            fs, field = float(data["fs"]), data["x"].copy()
        require(math.isfinite(fs) and fs > 0 and metadata["fs"] == fs and
                metadata["elements"] == len(mesh.tri), "Field Fs/elements differ")
        root = root_contract(metadata, mesh, mode, fs)
        prepared = core.Prepared(mesh, fs, mode, 8, reduction="none")
        audit = core.evaluate(prepared, field)
        require(audit["physical_pass"] and audit["witness"], "Original physical witness rejected: " + mode)
        sides[mode] = dict(fs=fs, elements=len(mesh.tri), mesh_id=mesh.mesh_id,
                           physical_model_id=mesh.physical_model_id, q=8, audit=audit, root_record=root)
        meshes[mode] = mesh
    lower, upper = sides["lower"]["fs"], sides["upper"]["fs"]
    gap = 2 * (upper - lower) / (upper + lower)
    require(0 < lower <= upper and gap <= .02 and
            meshes["lower"].physical_model_id == meshes["upper"].physical_model_id, "Original two-percent pair rejected")
    raw = (ROOT / "historical/inputs/D05A_new_vba_mesh.bin").read_bytes()
    require(sha(ROOT / "historical/inputs/D05A_new_vba_mesh.bin") == models.case_entry("D05A")["initial_mesh"]["sha256"],
            "Initial geometry differs from its original manifest")
    nn, ne = struct.unpack_from("<2i", raw)
    xy = np.frombuffer(raw, dtype="<f8", count=2 * nn, offset=8).reshape(nn, 2).copy()
    cells = np.frombuffer(raw, dtype="<i4", count=4 * ne, offset=8 + 16 * nn).reshape(ne, 4)
    coarse = core.mesh_from_arrays(xy, cells[:, :3].copy(), cells[:, 3].copy(), record)

    vectors = []
    for number, expected_numerical in ((5, True), (65, False)):
        stem = ROOT / "historical/vectors" / ("upper_attempt_%04d" % number)
        metadata = read(stem.with_suffix(".json"))
        mesh = coarse if number == 5 else meshes["upper"]
        require(metadata["elements"] == len(mesh.tri), "Vector mesh element count differs")
        prepared = core.Prepared(mesh, metadata["fs"], "upper", 8, reduction="none")
        reference_entry = references["programs"][str(number)]
        with np.load(inside(reference_entry["path"]), allow_pickle=False) as arrays:
            reference_program = unpack_program(arrays)
        require(program_digest(reference_program) == reference_entry["program_digest"] == metadata["original_program_sha256"] and
                metadata.get("original_program_unchanged") is True, "Original LP digest differs")
        program_comparison = compare_program(prepared.original, reference_program)
        with np.load(stem.with_suffix(".npz"), allow_pickle=False) as data:
            require(float(data["fs"]) == metadata["fs"] and str(data["mode"]) == "upper" and int(data["q"]) == 8,
                    "Full vector Fs/mode/q differs")
            field, dual = data["x"].copy(), data["original_dual"].copy()
            coordinate = data["coordinate_dual"].copy()
            scale, rows = float(data["velocity_scale"]), data["cone_row_scales"].copy()
        require(math.isfinite(scale) and scale > 0 and np.isfinite(rows).all() and np.all(rows > 0),
                "Invalid historical coordinate/cone scale")
        nb = len(prepared.original[1])
        restored = np.r_[coordinate[:nb], rows * coordinate[nb:]] / scale
        require(restored.tobytes() == dual.tobytes(), "Full original dual inverse mapping differs")
        audit = core.evaluate(prepared, field, dual)
        # Preserve the raw physical/numerical gates on the freshly assembled
        # program.  Recorded residuals bind to the exact old-digest reference,
        # because a BLAS-dependent last bit is not a new physical coefficient.
        observed_program = prepared.original
        prepared.original = reference_program
        try:
            reference_audit = core.evaluate(prepared, field, dual)
        finally:
            prepared.original = observed_program
        require(reference_audit["dual_stationarity"] == metadata["dual_stationarity"] and
                reference_audit["dual_cone_error"] == metadata["dual_cone_error"],
                "Historical raw dual metadata differs from the exact old-digest program")
        require(audit["physical_pass"] and bool(audit["numerical_gate"]) == expected_numerical and
                bool(reference_audit["numerical_gate"]) == expected_numerical,
                "Raw numerical acceptance/rejection did not reproduce under unchanged gates")
        require(bool(audit["witness"]) == (number == 65), "Expected physical witness distinction did not reproduce")
        vectors.append(dict(attempt=number, fs=metadata["fs"], elements=len(mesh.tri),
                            mesh_id=mesh.mesh_id, original_program_sha256=program_digest(prepared.original),
                            observed_program_digest_exactly_equal=program_digest(prepared.original) == metadata["original_program_sha256"],
                            historical_reference_program_digest_equal=True, program_comparison=program_comparison,
                            inverse_dual_mapping_byte_equal=True, reference_program_audit=reference_audit,
                            historical_solver_status=metadata["status"], velocity_scale=scale,
                            expected_raw_numerical_gate=expected_numerical, audit=audit))
    completion = read(ROOT / "historical/provenance/COMPLETION.json")
    require(completion["cold_numerical_process_exitcode"] is None and completion["independent_audit_exitcode"] == 0,
            "Historical unknown exit status was changed")
    require(before == {name: sha(original_file(name)) for name in before} and manifest_sha == sha(ROOT / "HISTORICAL_MANIFEST.json") and
            reference_before == {name: sha(inside(name)) for name in reference_before},
            "Evidence changed during replay")
    require(not calls, "An optimizer was called")
    result = dict(status="PASS", scope="saved_D05A_original_q8_pair_and_two_original_dual_vector_replay",
                  verified_utc=datetime.now(timezone.utc).isoformat(), case="D05A", new_optimizer_calls=0,
                  new_cold_execution=False, current_VBA_success_proved=False,
                  historical_worker_exit_code=None, historical_wrapper_seconds=summary["case_budget"]["elapsed_seconds"],
                  physical_pair_within_two_percent=True, lower_fs=lower, upper_fs=upper, gap_percent=100 * gap,
                  recorded_both_root_contracts_pass=True, opposite_root_signs_regenerated=False,
                  lower_and_upper_original_physical_audits=sides, full_vector_replays=vectors,
                  raw_rejected_vector_preserved=True, raw_stationarity_threshold=1e-7,
                  historical_files_unchanged=True, manifest_sha256=manifest_sha, checked_files=len(before),
                  original_verifier_relocated_read_only=True,
                  current_auditor_distinct_from_historical_verifier=True,
                  reference_manifest_sha256=sha(ROOT / "historical/program_refs/PROGRAM_REFERENCE_MANIFEST.json"),
                  current_auditor_sha256=sha(Path(__file__)),
                  elapsed_seconds=time.perf_counter() - started,
                  environment=dict(python=platform.python_version(), dependencies={name: importlib.metadata.version(name)
                                      for name in ("numpy", "scipy", "clarabel")}))
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf8")
    print(json.dumps({k: result[k] for k in ("status", "case", "gap_percent", "new_optimizer_calls", "elapsed_seconds")}), flush=True)


if __name__ == "__main__":
    main()
