//! Bounded original-equality correction search. No return value certifies physics.
//!
//! Real, undamped LSMR translated from scipy.sparse.linalg._isolve.lsmr
//! (SciPy 1.17.0), including its x0 and stopping rules. See LICENSE_LSMR.txt.
//! Algorithm: D. C.-L. Fong and M. A. Saunders, SIAM J. Sci. Comput. 33 (2011).
use super::*;

#[repr(C)]
#[derive(Clone, Copy)]
pub struct RpaEqualityProblem {
    pub struct_size: u32, pub abi_version: u32,
    pub n: u32, pub m: u32, pub nnz: u32, pub reserved: u32,
    pub row_ptr: *const i32, pub col_index: *const i32,
    pub a_values: *const f64, pub b: *const f64, pub x0: *const f64,
}

#[repr(C)]
#[derive(Clone, Copy, Default)]
pub struct RpaEqualityReport {
    pub struct_size: u32, pub abi_version: u32, pub istop: u32, pub iterations: u32,
    pub normr: f64, pub normar: f64, pub norma: f64, pub conda: f64, pub normx: f64,
}

struct Csr { n: usize, m: usize, rows: Vec<i32>, cols: Vec<i32>, values: Vec<f64> }
impl Csr {
    fn mul(&self, x: &[f64]) -> Vec<f64> {
        let mut out = vec![0.; self.m];
        for (r, y) in out.iter_mut().enumerate() {
            for j in self.rows[r] as usize..self.rows[r + 1] as usize {
                *y += self.values[j] * x[self.cols[j] as usize];
            }
        }
        out
    }
    fn tmul(&self, x: &[f64]) -> Vec<f64> {
        let mut out = vec![0.; self.n];
        for (r, &y) in x.iter().enumerate() {
            for j in self.rows[r] as usize..self.rows[r + 1] as usize {
                out[self.cols[j] as usize] += self.values[j] * y;
            }
        }
        out
    }
}

fn norm(x: &[f64]) -> f64 { x.iter().fold(0_f64, |v, &x| v.hypot(x)) }
fn scale(x: &mut [f64], factor: f64) { for v in x { *v *= factor; } }
fn sym_ortho(a: f64, b: f64) -> (f64, f64, f64) {
    let sign = |v: f64| if v == 0. { 0. } else { v.signum() };
    if b == 0. { (sign(a), 0., a.abs()) }
    else if a == 0. { (0., sign(b), b.abs()) }
    else if b.abs() > a.abs() {
        let tau = a / b; let s = sign(b) / (1. + tau * tau).sqrt();
        (s * tau, s, b / s)
    } else {
        let tau = b / a; let c = sign(a) / (1. + tau * tau).sqrt();
        (c, c * tau, a / c)
    }
}

fn lsmr(a: &Csr, b: &[f64], initial: &[f64], maxiter: u32,
        atol: f64, btol: f64, conlim: f64) -> Result<(Vec<f64>, RpaEqualityReport), i32> {
    let mut x = initial.to_vec();
    let ax = a.mul(&x);
    let mut u: Vec<f64> = b.iter().zip(ax).map(|(&b, ax)| b - ax).collect();
    let normb = norm(b); let mut beta = norm(&u);
    let mut v = vec![0.; a.n]; let mut alpha = 0.;
    if beta > 0. { scale(&mut u, 1. / beta); v = a.tmul(&u); alpha = norm(&v); }
    if alpha > 0. { scale(&mut v, 1. / alpha); }
    let mut zetabar = alpha * beta; let mut alphabar = alpha;
    let (mut rho, mut rhobar, mut cbar, mut sbar) = (1., 1., 1., 0.);
    let mut h = v.clone(); let mut hbar = vec![0.; a.n];
    let (mut betadd, mut betad, mut rhodold) = (beta, 0., 1.);
    let (mut tautildeold, mut thetatilde, mut zeta) = (0., 0., 0.);
    let mut norm_a2 = alpha * alpha;
    let (mut maxrbar, mut minrbar) = (0_f64, 1e100_f64);
    let mut report = RpaEqualityReport { struct_size: size_of::<RpaEqualityReport>() as u32,
        abi_version: ABI, normr: beta, normar: alpha * beta,
        norma: norm_a2.sqrt(), conda: 1., ..Default::default() };
    // SciPy early exits preserve x0 when A^T r is zero; normx is its estimate.
    if report.normar == 0. { return Ok((x, report)); }
    if normb == 0. { x.fill(0.); return Ok((x, report)); }
    let ctol = if conlim > 0. { 1. / conlim } else { 0. };
    for itn in 1..=maxiter {
        let av = a.mul(&v);
        for (u, av) in u.iter_mut().zip(av) { *u = -alpha * *u + av; }
        beta = norm(&u);
        if beta > 0. {
            scale(&mut u, 1. / beta);
            let atu = a.tmul(&u);
            for (v, atu) in v.iter_mut().zip(atu) { *v = -beta * *v + atu; }
            alpha = norm(&v);
            if alpha > 0. { scale(&mut v, 1. / alpha); }
        }
        // Undamped LSMR: Qhat still matters for signs of alphabar.
        let (chat, shat, alphahat) = sym_ortho(alphabar, 0.);
        let rhoold = rho;
        let (c, s, newrho) = sym_ortho(alphahat, beta); rho = newrho;
        let thetanew = s * alpha; alphabar = c * alpha;
        let rhobarold = rhobar; let zetaold = zeta;
        let thetabar = sbar * rho; let rhotemp = cbar * rho;
        let rotation = sym_ortho(cbar * rho, thetanew);
        cbar = rotation.0; sbar = rotation.1; rhobar = rotation.2;
        zeta = cbar * zetabar; zetabar *= -sbar;
        let hb_scale = -(thetabar * rho / (rhoold * rhobarold));
        let x_scale = zeta / (rho * rhobar); let h_scale = -(thetanew / rho);
        for j in 0..a.n {
            hbar[j] = hb_scale * hbar[j] + h[j];
            x[j] += x_scale * hbar[j]; h[j] = h_scale * h[j] + v[j];
        }
        let betaacute = chat * betadd; let betacheck = -shat * betadd;
        let betahat = c * betaacute; betadd = -s * betaacute;
        let thetatildeold = thetatilde;
        let (ctildeold, stildeold, rhotildeold) = sym_ortho(rhodold, thetabar);
        thetatilde = stildeold * rhobar; rhodold = ctildeold * rhobar;
        betad = -stildeold * betad + ctildeold * betahat;
        tautildeold = (zetaold - thetatildeold * tautildeold) / rhotildeold;
        let taud = (zeta - thetatilde * tautildeold) / rhodold;
        // betacheck is identically zero without damping.
        report.normr = (betacheck * betacheck + (betad - taud).powi(2)
            + betadd * betadd).sqrt();
        norm_a2 += beta * beta; report.norma = norm_a2.sqrt(); norm_a2 += alpha * alpha;
        maxrbar = maxrbar.max(rhobarold);
        if itn > 1 { minrbar = minrbar.min(rhobarold); }
        report.conda = maxrbar.max(rhotemp) / minrbar.min(rhotemp);
        report.normar = zetabar.abs(); report.normx = norm(&x); report.iterations = itn;
        let test1 = report.normr / normb;
        let test2 = if report.norma * report.normr != 0. {
            report.normar / (report.norma * report.normr)
        } else { f64::INFINITY };
        let test3 = 1. / report.conda;
        let t1 = test1 / (1. + report.norma * report.normx / normb);
        let rtol = btol + atol * report.norma * report.normx / normb;
        let mut istop = 0;
        if itn >= maxiter { istop = 7; }
        if 1. + test3 <= 1. { istop = 6; }
        if 1. + test2 <= 1. { istop = 5; }
        if 1. + t1 <= 1. { istop = 4; }
        if test3 <= ctol { istop = 3; }
        if test2 <= atol { istop = 2; }
        if test1 <= rtol { istop = 1; }
        report.istop = istop;
        if !finite(&x) || !finite(&[report.normr, report.normar, report.norma,
                report.conda, report.normx]) { return Err(SOLVER); }
        if istop != 0 { break; }
    }
    Ok((x, report))
}

/// Solve a real undamped CSR correction system with a copied initial correction.
/// Caller retains its original physical field and audits source + output fully.
#[no_mangle]
pub unsafe extern "C" fn rpfem_accuracy_equalities(p: *const RpaEqualityProblem,
    maxiter: u32, atol: f64, btol: f64, conlim: f64,
    out: *mut f64, out_len: u32, report: *mut RpaEqualityReport) -> i32 {
    ffi(|| {
        header(p)?; header(report)?;
        if maxiter == 0 || maxiter > 1000 || !finite(&[atol, btol, conlim])
            || atol < 0. || btol < 0. || conlim < 0. { return Err(INVALID); }
        let input = ptr::read(p);
        let (n, m, nz) = (input.n as usize, input.m as usize, input.nnz as usize);
        if n == 0 || n > MAX_N || m > MAX_M || nz > MAX_NNZ { return Err(LIMIT); }
        if input.reserved != 0 { return Err(INVALID); }
        if out_len as usize != n { return Err(BUFFER); }
        checked_ptr(out, n)?;
        // Output buffers may not mutate any original input or one another.
        if overlap(out, n, report, 1) || overlap(out, n, p, 1)
            || overlap(out, n, input.row_ptr, m + 1) || overlap(out, n, input.col_index, nz)
            || overlap(out, n, input.a_values, nz) || overlap(out, n, input.b, m)
            || overlap(out, n, input.x0, n) || overlap(report, 1, p, 1)
            || overlap(report, 1, input.row_ptr, m + 1) || overlap(report, 1, input.col_index, nz)
            || overlap(report, 1, input.a_values, nz) || overlap(report, 1, input.b, m)
            || overlap(report, 1, input.x0, n) { return Err(INVALID); }
        let rows = copied(input.row_ptr, m + 1)?; let cols = copied(input.col_index, nz)?;
        let values = copied(input.a_values, nz)?; let b = copied(input.b, m)?;
        let initial = copied(input.x0, n)?;
        if !finite(&values) || !finite(&b) || !finite(&initial)
            || rows[0] != 0 || rows[m] != nz as i32 { return Err(INVALID); }
        for r in 0..m {
            let (start, end) = (rows[r], rows[r + 1]);
            if start < 0 || end < start || end as usize > nz { return Err(INVALID); }
            let mut previous = None;
            for &column in &cols[start as usize..end as usize] {
                if column < 0 || column as usize >= n
                    || previous.is_some_and(|v| v >= column) { return Err(INVALID); }
                previous = Some(column);
            }
        }
        let matrix = Csr { n, m, rows, cols, values };
        let (x, diagnostics) = lsmr(&matrix, &b, &initial, maxiter, atol, btol, conlim)?;
        if !finite(&x) || !finite(&[diagnostics.normr, diagnostics.normar, diagnostics.norma,
            diagnostics.conda, diagnostics.normx]) { return Err(SOLVER); }
        ptr::copy_nonoverlapping(x.as_ptr(), out, n); ptr::write(report, diagnostics);
        Ok(())
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn equality_abi_sizes() {
        assert_eq!(size_of::<RpaEqualityProblem>(), 64);
        assert_eq!(size_of::<RpaEqualityReport>(), 56);
    }
    #[test]
    fn inconsistent_rank_deficient_correction() {
        let a = Csr { n: 2, m: 3, rows: vec![0, 2, 4, 6],
            cols: vec![0, 1, 0, 1, 0, 1], values: vec![1., 1., 2., 2., 3., 3.] };
        let (x, report) = lsmr(&a, &[1., 2., 4.], &[0., 0.], 100, 1e-15, 1e-15, 1e8).unwrap();
        assert!((x[0] - 17. / 28.).abs() < 1e-12);
        assert!((x[1] - 17. / 28.).abs() < 1e-12);
        assert_eq!(report.istop, 2);
    }
}
