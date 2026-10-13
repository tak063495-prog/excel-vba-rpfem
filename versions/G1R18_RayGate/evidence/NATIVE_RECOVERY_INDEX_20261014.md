The two archives contain completed owned-Excel experiments, original literal
matrices, primal fields, original dual vectors, exact model context and full
audit/source/PID hashes. Every member has a SHA256 in its adjacent manifest.
Curated failures remain failures. No workbook, DLL, live P06A production output
or duplicate tamper payload is included.

| Archive | Bytes | SHA256 | Scope |
| --- | ---: | --- | --- |
| native_recovery_primal_evidence_20261014.zip | 43067560 | 233c690d7f9eb5e92e462aa63c985dd816e7dc56a97ee1334b67642c3d26576c | Five P06 primal imports, rejected AHO3/near-root AF attempts, fixed-load AF rollback, passive upper ray/wedge, recovered lower primal plus separate original free-load dual root |
| native_p03_primal_evidence_20261014.zip | 37548838 | 91abb9d68aed5ae6f1a5267b927d773591899c3c76adca040b2894c2f6cfab4f | Frozen progress6 P03A/P03N upper physical imports; root/status/bounds are not imported |

The final recovery/root fixture is adapt_excel_1791910801411367300. Its original
free-load positive primal at Fs=1.37329120594012 and separate original negative
dual at Fs=1.373382541241769 pass the unchanged original raw gates and independent
matrix/vector replay. Metadata and nonfinite working-cost rejection tests pass.
Both saved lower/upper fields and bounds remain unchanged.

P03A/P03N original field imports pass at Fs=1.17069 / 1.167378 with native
physical values 0.999951229262553 / 0.999997991247848. These demonstrate original
model/field parity. They do not install a native root or infer a precision goal.

The ordinary G1R18 P06A baseline remains a separate live experiment. The optional
LOWER_FIXED_RECOVERY setting is disabled by default and has not been adopted
for production. Actual P06 numerical recovery still needs validation. The DLL
and ABI are unchanged; no physical/numerical gate was relaxed.

Rebuild and verify every archive member from retained completed local folders:

```powershell
python versions/G1R18_RayGate/tools/curate_recovery_evidence.py
python versions/G1R18_RayGate/tools/curate_recovery_evidence.py --only python_field_excel_1791908825521159200 --only python_field_excel_1791910355270646900 --name native_p03_primal_evidence_20261014
```
