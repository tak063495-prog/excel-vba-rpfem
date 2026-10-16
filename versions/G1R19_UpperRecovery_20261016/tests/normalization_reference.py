"""Independent small-array reference for the homogeneous work identity; no solver."""
from dataclasses import dataclass
import math


@dataclass
class Program:
    rows: list
    rhs: list
    equalities: int
    mode: str = 'upper'
    q: int = 8
    phase: str = 'reference_total'


def normalize(program, candidate):
    """Reject an inhomogeneous cone/equality context before positive rescaling."""
    x = list(candidate)
    if (program.mode != 'upper' or program.q != 8 or program.phase != 'reference_total'
            or not all(math.isfinite(v) for v in program.rhs + x)
            or any(v != 0 for v in program.rhs[program.equalities:])):
        return x, None, False
    rows = [i for i, v in enumerate(program.rhs[:program.equalities]) if v != 0]
    if len(rows) != 1:
        return x, None, False
    row = rows[0]
    ratio = math.fsum(a*b for a, b in zip(program.rows[row], x)) / program.rhs[row]
    if not math.isfinite(ratio) or ratio <= 0:
        return x, ratio, False
    y = [v / ratio for v in x]
    return (y, ratio, True) if all(math.isfinite(v) for v in y) else (x, ratio, False)


def fixture():
    return Program([[.5, .5], [1., -1.], [-1., 0.]], [1., 0., 0.], 2)
