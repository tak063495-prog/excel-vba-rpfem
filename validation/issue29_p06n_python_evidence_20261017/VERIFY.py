"""Re-audit the archived P06N pair, without optimization or a cold run.

All original records and copied Python sources retain their exact bytes.  Only
the explicit manifest relocation table resolves their historical file names.
The lower opposite root sign is a retained record: its full vector was not
saved and this verifier does not pretend to recreate it.
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
import sys
import time

sys.dont_write_bytecode = True
for _key in ("OPENBLAS_NUM_THREADS", "OMP_NUM_THREADS", "MKL_NUM_THREADS", "NUMEXPR_NUM_THREADS"):
    os.environ[_key] = "1"
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

    def nonfinite(value):
        raise ValueError("Nonfinite JSON constant: " + value)

    return json.loads(Path(path).read_text(encoding="utf8"), object_pairs_hook=pairs,
                      parse_constant=nonfinite)


def bundle_path(name):
    require(isinstance(name, str) and "\\" not in name, "Relative POSIX bundle path required")
    pure = PurePosixPath(name)
    require(not pure.is_absolute() and ".." not in pure.parts, "Unsafe bundle path: " + name)
    path = (ROOT / name).resolve()
    require(path.is_relative_to(ROOT), "Bundle path escaped its directory")
    return path


def original_relative(name):
    """Translate names only; never edit or consult a historical absolute path."""
    require(isinstance(name, str), "A recorded file name must be a string")
    name = name.replace("\\", "/")
    prefix = "D:/Work/excel-vba-rpfem/"
    if name.startswith(prefix):
        name = name[len(prefix):]
    pure = PurePosixPath(name)
    require(not pure.is_absolute() and ":" not in name and ".." not in pure.parts,
            "Unknown historical path prefix: " + name)
    return pure.as_posix()


def check_root(metadata, mesh, mode, fs):
    identity = dict(case="P06N", mode=mode, q=8, mesh_id=mesh.mesh_id,
                    physical_model_id=mesh.physical_model_id,
                    policy=mesh.provenance["psi_policy"])
    root = metadata.get("audit", {}).get("discrete_root_search", {})
    require(all(root.get(k) == v for k, v in identity.items()), "Root identity mismatch: " + mode)
    require(all(metadata.get(k) == v for k, v in identity.items() if k != "policy"),
            "Certificate identity mismatch: " + mode)
    bracket = root.get("bracket")
    require(isinstance(bracket, list) and len(bracket) == 2 and
            all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)
                for v in bracket) and 0 < bracket[0] < bracket[1], "Invalid recorded root bracket")
    width = bracket[1] - bracket[0]
    require(width <= 1e-4 * max(1., sum(bracket) / 2), "Recorded root is too wide")
    require(root.get("root_complete") is True and metadata.get("root_complete") is True and
            root.get("bracket_signs") == [1, -1] and
            root.get("endpoint") == fs == bracket[0 if mode == "lower" else 1],
            "Recorded root completion/signs/endpoint mismatch")
    return dict(bracket=bracket, normalized_width=width / max(1., sum(bracket) / 2),
                recorded_identity_signs_and_width_pass=True,
                opposite_sign_recomputed_from_full_vector=mode == "upper")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True, help="A new JSON file outside the evidence directory")
    args = parser.parse_args()
    output = args.output.resolve()
    require(not output.exists(), "Refusing to overwrite an output")
    require(not output.is_relative_to(ROOT), "Write results outside the immutable evidence directory")
    require(output.parent.is_dir(), "The output parent directory must already exist")
    started = time.perf_counter()
    manifest = read(ROOT / "MANIFEST.json")
    require(manifest.get("schema") == "issue29_p06n_saved_pair_v1", "Unexpected manifest schema")
    expected = manifest["files"]
    before = {}
    for name, item in expected.items():
        path = bundle_path(name)
        require(path.is_file() and path.stat().st_size == item["bytes"] and sha(path) == item["sha256"],
                "Evidence bytes differ: " + name)
        before[name] = item["sha256"]
    manifest_sha = sha(ROOT / "MANIFEST.json")
    relocations = manifest["read_only_relocation"]

    def relocated(name, digest):
        relative = original_relative(name)
        require(relative in relocations, "Historical reference not included in this focused bundle: " + relative)
        bundled = relocations[relative]
        path = bundle_path(bundled)
        require(bundled in expected and sha(path) == digest, "Relocated historical reference differs: " + relative)
        return path

    sys.path.insert(0, str(ROOT / "python_snapshot"))
    import experiment_core as core
    import benchmark_models as models
    from original_dual import audit_original_dual
    np = core.np
    solver_calls = []

    def forbidden_solver(*unused, **ignored):
        solver_calls.append("clarabel.DefaultSolver")
        raise RuntimeError("NEW_OPTIMIZATION_FORBIDDEN_IN_SAVED_FIELD_VERIFIER")

    core.clarabel.DefaultSolver = forbidden_solver
    # The original model builder remains byte-identical.  Its input manifest
    # is supplied in memory so no historical absolute path is opened.
    models._MANIFEST = read(ROOT / "inputs/INPUTS_26.json")
    models.install_model_builder()
    record = read(ROOT / "inputs/P06N.json")
    models._validate_record(record)
    require(record["case"] == "P06N" and record["source_case"]["policy"] == "hold_initial" and
            not record.get("water", {}).get("coupling"), "Unexpected canonical model")
    pair = ROOT / "data/pair"
    summary = read(pair / "summary.json")
    component_manifest = read(pair / "COMPONENT_MANIFEST.json")
    for name, digest in component_manifest["collector_reaudit_source_sha256"].items():
        relocated(name, digest)
    repair_producer = read(ROOT / "data/upper_dual/summary.json")
    require(repair_producer["source_sha256_before"] == repair_producer["source_sha256_after"] and
            repair_producer.get("source_unchanged") is True, "Historical repair source freeze differs")
    for name, digest in repair_producer["source_sha256_before"].items():
        relocated(name, digest)
    scope = summary["execution_scope"]
    require(scope.get("kind") == "known_recipe_geometry_replay" and
            scope.get("historical_geometry_imported") is True and
            scope.get("historical_field_imported") is False and
            scope.get("historical_fs_imported") is False and scope.get("generic_cold_start") is False,
            "Historical execution classification differs")
    results, meshes = {}, {}
    for mode in ("lower", "upper"):
        metadata = read(pair / ("best_" + mode + "_certificate.json"))
        require(metadata == summary[mode], "Pair summary differs from the original certificate")
        require(metadata["mesh_path"] == "best_" + mode + "_mesh.npz" and
                metadata["field_path"] == "best_" + mode + "_field.npz", "Unexpected pair artifact names")
        for name, digest in metadata["source_hashes"].items():
            relocated(name, digest)
        with np.load(pair / metadata["mesh_path"], allow_pickle=False) as data:
            mesh = core.mesh_from_arrays(data["xy"].copy(), data["tri"].copy(),
                                         data["material_id"].copy(), record)
        with np.load(pair / metadata["field_path"], allow_pickle=False) as data:
            require(str(data["mode"]) == mode and int(data["q"]) == 8, "Field mode/q mismatch")
            fs, field = float(data["fs"]), data["x"].copy()
        require(math.isfinite(fs) and fs > 0 and np.isfinite(field).all(), "Nonfinite/nonpositive witness")
        require(metadata["fs"] == fs and metadata["elements"] == len(mesh.tri), "Fs/element count mismatch")
        root = check_root(metadata, mesh, mode, fs)
        prepared = core.Prepared(mesh, fs, mode, 8, reduction="none")
        audit = core.evaluate(prepared, field)
        require(audit["physical_pass"] is True and audit["witness"] is True,
                "Original q8 physical or witness gate rejected: " + mode)
        require(audit["audit"]["pass"] is True, "Independent physical audit rejected: " + mode)
        meshes[mode] = mesh
        results[mode] = dict(fs=fs, q=8, elements=len(mesh.tri), mesh_id=mesh.mesh_id,
                             physical_model_id=mesh.physical_model_id, policy=mesh.provenance["psi_policy"],
                             original_q8_physical_and_witness_pass=True, audit=audit,
                             root_record=root)
        del prepared, field
    require(meshes["lower"].physical_model_id == meshes["upper"].physical_model_id,
            "The physical models differ")
    lower, upper = results["lower"]["fs"], results["upper"]["fs"]
    gap = 2 * (upper - lower) / (upper + lower)
    require(0 < lower <= upper and gap <= .02, "The unchanged symmetric two-percent gate rejected the pair")

    dual_results = {}
    vectors = {}
    identity = dict(case="P06N", mode="upper", q=8, mesh_id=meshes["upper"].mesh_id,
                    physical_model_id=meshes["upper"].physical_model_id, policy="hold_initial")
    upper_root = results["upper"]["root_record"]["bracket"]
    prepared = core.Prepared(meshes["upper"], upper_root[0], "upper", 8, reduction="none")
    before_record = read(ROOT / "data/upper_dual/terminal_original_dual_audit.json")
    require(all(before_record.get(k) == v for k, v in identity.items() if k != "policy") and
            before_record.get("fs") == upper_root[0], "Companion original dual context/Fs mismatch")
    for label, name in (("before_repair", "terminal_original_dual.npz"),
                        ("after_repair", "repaired_original_dual.npz")):
        with np.load(ROOT / "data/upper_dual" / name, allow_pickle=False) as data:
            # The original helper NPZ stored mode/q/Fs only.  Its full model
            # identity is bound by the unchanged companion audit and SHA map.
            keys = identity if label == "after_repair" else ("mode", "q")
            observed = {k: (int(data[k]) if k == "q" else str(data[k])) for k in keys}
            require(observed == {k: identity[k] for k in keys} and float(data["fs"]) == upper_root[0],
                    "Original dual context/Fs mismatch")
            vectors[label] = data["dual"].copy()
        dual_results[label] = audit_original_dual(prepared, vectors[label])
    require(dual_results["before_repair"]["accepted"] is False and
            dual_results["after_repair"]["accepted"] is True and
            dual_results["after_repair"]["search_sign"] == 1,
            "Original dual rejection/acceptance did not reproduce")
    ny = len(prepared.original[1])
    require(vectors["before_repair"][ny:].tobytes() == vectors["after_repair"][ny:].tobytes(),
            "Original cone variable z changed during the historical repair")
    after = {name: sha(bundle_path(name)) for name in expected}
    require(before == after and manifest_sha == sha(ROOT / "MANIFEST.json"), "Evidence changed during the audit")
    require(not solver_calls, "An optimizer was invoked")
    result = dict(schema="issue29_p06n_saved_pair_verification_v1", status="PASS",
                  verified_utc=datetime.now(timezone.utc).isoformat(), case="P06N",
                  evidence_scope="saved_pair_original_q8_reaudit_and_full_upper_original_dual_reaudit",
                  new_optimization=False, optimizer_calls=0, new_cold_execution=False,
                  generic_cold_two_percent_achievement_proved=False,
                  source_and_input_bytes_unchanged=True, manifest_sha256=manifest_sha,
                  verified_file_count=len(before), lower_fs=lower, upper_fs=upper,
                  relative_gap=gap, gap_percent=100 * gap, target_gap=.02,
                  physical_pair_within_two_percent=True, recorded_both_root_contracts_pass=True,
                  opposite_root_signs_recomputed_from_full_vectors=dict(lower=False, upper=True),
                  lower_opposite_sign_limit="The historical trial/identity/width record is retained; no full opposite-endpoint lower vector was saved. Its sign is not independently regenerated here.",
                  sides=results, upper_original_dual=dual_results,
                  historical_repair_cone_z_bitwise_unchanged=True,
                  elapsed_seconds=time.perf_counter() - started,
                  environment=dict(python=platform.python_version(), platform=platform.platform(),
                                   dependencies={name: importlib.metadata.version(name)
                                                 for name in ("numpy", "scipy", "clarabel")}))
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf8")
    print(json.dumps({k: result[k] for k in ("status", "case", "gap_percent", "optimizer_calls", "elapsed_seconds")}), flush=True)


if __name__ == "__main__":
    main()
