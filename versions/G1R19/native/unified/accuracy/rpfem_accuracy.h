#ifndef RPFEM_ACCURACY_H
#define RPFEM_ACCURACY_H
#include <stdint.h>
#ifdef __cplusplus
extern "C" {
#endif
/* Windows x64 C ABI. Structs contain explicit pairs of uint32 then doubles
   or pointers, so all 8-byte members start at 8-byte-aligned offsets. */
#define RPFEM_ACCURACY_ABI 1u
enum { RPA_OK=0, RPA_INVALID_ARGUMENT=-1, RPA_LIMIT=-2,
       RPA_INVALID_HANDLE=-3, RPA_STATE=-4, RPA_BUFFER=-5,
       RPA_PANIC=-6, RPA_SOLVER_ERROR=-7, RPA_BUSY=-8 };
enum { RPA_CREATED=0, RPA_RUNNING=1, RPA_DONE=2,
       RPA_CANCELLED=3, RPA_ERROR=4 };
enum { RPA_ZERO_CONE=0, RPA_NONNEGATIVE_CONE=1, RPA_SOC_CONE=3 };
typedef uint64_t rpa_handle;
/* size=88; row_ptr offset24, col_index32, A_values40,b48,Pdiag56,q64,
   cone_kind72, cone_dim80. CSR is zero-based and sorted within each row;
   duplicate columns are rejected. P is diagonal nonnegative, all data finite.
   Each pointer is borrowed ONLY for create; its declared allocation must be
   readable and correctly aligned. All arrays are copied before create returns. */
typedef struct {
    uint32_t struct_size, abi_version;
    uint32_t n, m, nnz, cone_count;
    const int32_t *row_ptr, *col_index;
    const double *A_values, *b, *Pdiag, *q;
    const int32_t *cone_kind, *cone_dim;
} rpa_problem;
/* size=128; doubles begin at16..72, integer flags at80/84, static eps88,
   dynamic eps96, dynamic delta104, flags112/116, method120/reserved124.
   direct_method:0 auto,1 qdldl,2 faer. Explicit solver tolerances NEVER replace
   the caller's unscaled original physical/primal/dual audit. */
typedef struct {
    uint32_t struct_size, abi_version, max_iter, max_threads;
    double time_limit, tol_gap_abs, tol_gap_rel, tol_feas;
    double tol_infeas_abs, tol_infeas_rel;
    double iterative_refinement_reltol, iterative_refinement_abstol;
    uint32_t iterative_refinement_max_iter, equilibrate;
    double static_regularization_constant;
    double dynamic_regularization_eps, dynamic_regularization_delta;
    uint32_t static_regularization_enable, dynamic_regularization_enable;
    uint32_t direct_method, reserved;
} rpa_settings;
/* size=80; state/status offset8/12, iter/cancel16/20,n/m24/28;
   diagnostics at32..72. Status is the Clarabel enum integer, metadata only.
   diagnostics_valid=1 means each reported scalar is finite. */
typedef struct {
    uint32_t struct_size, abi_version, state, solver_status;
    uint32_t iterations, cancel_requested, n, m;
    double solve_time, primal_objective, dual_objective;
    double primal_residual, dual_residual, diagnostics_valid;
} rpa_progress;

uint32_t rpfem_accuracy_abi_version(void);
int32_t rpfem_accuracy_default_settings(rpa_settings *settings);
int32_t rpfem_accuracy_create(const rpa_problem *problem,
                            const rpa_settings *settings, rpa_handle *handle);
int32_t rpfem_accuracy_start(rpa_handle handle);
int32_t rpfem_accuracy_poll(rpa_handle handle, rpa_progress *progress);
int32_t rpfem_accuracy_cancel(rpa_handle handle);
/* Exact lengths n,m are mandatory. z is the dual of THIS generic supplied
   problem. A QP/cap/floor/helper constraint dual is NOT an original guide.
   result returns only finite x/z after worker completion; status is no proof. */
int32_t rpfem_accuracy_result(rpa_handle handle, double *x, uint32_t x_length,
                            double *z, uint32_t z_length, rpa_progress *progress);
/* A running worker is cancelled and BUSY returned. Poll with Excel DoEvents
   until completed, then destroy; destruction never blocks a running solve. */
int32_t rpfem_accuracy_destroy(rpa_handle handle);

/* size=64; pointers row_ptr24,col_index32,A_values40,b48,x0=56.
   This is an ordinary bounded LSMR correction problem A*correction=b.
   For staged weighted equilibrium recovery, A is row_norm*original_E,
   b is the immutable row_norm*(original_b-original_E*source), and x0 is
   the preceding correction (initially a full-length allocated zero array).
   Return is the updated correction, not source+correction. All pointers
   must be readable/aligned for their declared lengths and are copied. */
typedef struct {
    uint32_t struct_size, abi_version, n, m, nnz, reserved;
    const int32_t *row_ptr, *col_index;
    const double *A_values, *b, *x0;
} rpa_equality_problem;
/* size=56; doubles normr16,normar24,norma32,conda40,normx48.
   istop/iteration/norms are numerical diagnostics, never physical proof. */
typedef struct {
    uint32_t struct_size, abi_version, istop, iterations;
    double normr, normar, norma, conda, normx;
} rpa_equality_report;
/* Synchronous, bounded to 1..1000 iterations. Exact n output required.
   Caller must audit source+correction against the whole unscaled original
   physical program before adoption, and preserve source on rejection. */
int32_t rpfem_accuracy_equalities(const rpa_equality_problem *problem,
    uint32_t maxiter, double atol, double btol, double conlim,
    double *correction, uint32_t correction_length,
    rpa_equality_report *report);
#ifdef __cplusplus
}
#endif
#endif
