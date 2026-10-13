Python primal fields can be independently checked by the existing VBA original
program and physical audit. The test loader is isolated from production and
does not import solver status, dual vectors, brackets, root completion, or
accuracy achievement. It never installs a bound or certificate.

The exporter reconstructs the explicit benchmark model, checks case / side /
Fs / q8 / mesh identity / physical model identity against the field and source
certificate, and runs a fresh unreduced Python primal and physical audit. Its
RP25M001 payload contains only geometry, material, body and boundary data,
primal field, Fs and q. Source root/status metadata are either recorded as
ignored provenance or recoverable from the hashed source certificate.

VBA reads the original benchmark definition from its input sheets. It checks
c/phi/psi/gamma and stress scale exactly, positive triangles, material regions
and their areas, original total area, reference gravity, and every original
boundary restraint and load. It reconstructs topology and definition segments
without moving a node or changing a triangle. The test supports the dry,
soil-only P03/P06 benchmarks with partitioned regions and reference-total gravity.
Other models require a separate loader and are rejected by this test.

Lower stresses are copied without modification. Upper velocities are copied
without modification and VBA rebuilds any auxiliary epigraph prices from its
original cones. The original literal equations/cones and full physical gate
must both pass; the lower raw physical errors remain limited to 1e-7. Upper
value must satisfy the existing 0.999999 + 1e-8 threshold, and lower value must
be at least 1.000001. The loader rejects q4, NaN coordinates and changed
material values before any audit. Tests verify unchanged nonzero bound/status
sentinels and absence of a newly created native root.

Completed real Excel checks are recorded separately below. A source root flag
describes Python provenance only; all native import root flags remain absent.

| Source / retained test directory | Elements | Fs | Source root | Native physical value | Result |
| --- | ---: | ---: | --- | ---: | --- |
| n_lower_anglefan5594_close/P06N / python_field_excel_1791903229682201900 | 5594 lower | 0.6767836604585639 | true | 1.00000728583069 | PASS |
| n_flow6106_jump1_close/P06N / python_field_excel_1791903712492988200 | 6106 upper | 0.6914407501171215 | false | 0.999353332335558 | PASS |
| a_jump_flow5659/P06A / python_field_excel_1791903997288551900 | 5659 upper | 0.861 | false | 0.991857115939689 | PASS |
| n_flow5_c1_close/P06N / python_field_excel_1791904224800777700 | 6106 upper | 0.6914407501171215 | true | 0.999998574829121 | PASS |
| a_lower20792_close/P06A / python_field_excel_1791905614894344800 | 20792 lower | 0.8458963353848571 | false | 1.00085859385119 | PASS |
| frozen progress6 P03A / python_field_excel_1791908825521159200 | 13325 upper | 1.17069 | true | 0.999951229262553 | PASS |
| frozen progress6 P03N / python_field_excel_1791910355270646900 | 13363 upper | 1.167378 | true | 0.999997991247848 | PASS |

The two N upper meshes are different. The fresh jump mesh is
ac54daa613bf3492792ef1592fbe534ebb04a15e3ff0edbc10730b4c49eafdd5,
with field NPZ SHA256
4b92e3d4205e98f8c826129d1672d9f9340c96fc17d7c109ffafb5ac09c9bd88.
The completed original-dual recovery mesh is
87cdea024b90db6080d2fbdac23f74272ed2bf238abcb3f87cce5b7bd8ff6684,
with field NPZ SHA256
975349c8b8058da899bca875a012be7c3e8cd8df40e6effc989a5aac9d139832.
Both are independently valid primal witnesses, but only the second source has
completed Python root metadata. Native import does not inherit that metadata.

Each test retains source hashes, original input and DLL hashes, actual loaded
VBA hashes, owned Excel PID, independent exact binary roundtrip, a fresh Python
original audit of the native field, the literal VBA original program and
field, and an independent streamed fsum/hypot replay of that program. The
first lower test also has a separate independent_vba_matrix_replay.json.
P03 inputs are read directly from the immutable benchmark catalogue when no
separate experiment input JSON exists. The source_case must exactly match the
Python reconstruction, and coupled pore cases are rejected before export.
Temporary workbooks and DLL copies are test outputs, not deliverables. The DLL
has not changed.

Reproduce from the repository root using the configured Python environment:

```powershell
python -X utf8 versions/G1R18/tests/test_python_witness_excel.py --case P06N --mode upper --folder validation/accuracy_20261013/results/n_flow5_c1_close/P06N
```

The test creates a fresh hidden owned Excel process, records all existing PIDs,
and never attaches to an existing Excel process or changes macro security.
There is no HSD/KKT solve in this import test. Passing a field import does not
establish a native root or native 1% achievement.
