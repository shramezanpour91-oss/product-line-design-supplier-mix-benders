# ============================================================
# MANUAL BENDERS (BD2) — CLOSED-FORM SUBPROBLEM VERSION
#
# CLEANED VERSION WITHOUT A CPLEX LP SUBPROBLEM:
#   - beta1 / beta2 / beta3 / beta4 / beta5 are computed directly by formula.
#   - no dual LP model is built, solved, or used inside callbacks.
#   - lazy Benders cuts are generated from the computed beta values.
#   - fractional user cuts are disabled because the formulas are incumbent-based.
#   - beta1 component rows are collected silently for Excel export.
#   - z has a safe finite upper bound in the master.
# ============================================================

import pandas as pd
import cplex
from cplex.callbacks import LazyConstraintCallback
import sys
import bisect
import os
import time
import math
import itertools
from datetime import datetime


# ============================================================
# USER SETTINGS FOR BATCH RUNS
# ============================================================

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
address = BASE_DIR / "data" / "Main Dataset.xlsx"

N_values = [10]
P_values = [1000]
K_values = [200]
V_values = [3]
S_values = [4]

'''
N_values = [10, 20, 40]
P_values = [1000, 2000, 4000, 8000, 16000]
K_values = [200, 400]
V_values = [3, 5]
S_values = [2, 4]
'''

ARTIFICIAL_SUPPLIER_COST = 1000.0
ARTIFICIAL_SUPPLIER_CAP = 50000.0

TIME_LIMIT = 3600
THREADS = 1

# Console printing controls for final results.
# Keep PRINT_FULL_FINAL_TABLES = False for large instances.
PRINT_FINAL_RESULTS_TO_CONSOLE = False
PRINT_FULL_FINAL_TABLES = False
FINAL_PREVIEW_ROWS = 20


def format_excel_worksheets(writer):
    workbook = writer.book
    one_decimal = workbook.add_format({"num_format": "0.#"})
    text_format = workbook.add_format({"text_wrap": True})

    for sheet_name, worksheet in writer.sheets.items():
        worksheet.set_column(0, 200, 16, one_decimal)

        if sheet_name == "summary":
            worksheet.set_column(0, 0, 35, text_format)
            worksheet.set_column(1, 1, 22, one_decimal)

        if sheet_name == "summary_all_runs":
            worksheet.set_column(0, 200, 18, one_decimal)
            worksheet.set_column(20, 35, 45, text_format)


# ----------------------------
# Utilities: pivotal supplier + cost f(A)
# ----------------------------

def build_supplier_prefix(nP, nS, r_arr, Cap_arr):
    cap_prefix = [None] * nP
    capr_prefix = [None] * nP
    r_sorted = [None] * nP
    total_cap = [0.0] * nP
    suppliers_sorted = [None] * nP

    for p in range(nP):
        suppliers = sorted(range(nS), key=lambda s: float(r_arr[p][s]))
        suppliers_sorted[p] = suppliers

        caps = []
        caprs = []
        rs = []

        run_c = 0.0
        run_cr = 0.0
        for s in suppliers:
            c = float(Cap_arr[p][s])
            rr = float(r_arr[p][s])
            run_c += c
            run_cr += c * rr
            caps.append(run_c)
            caprs.append(run_cr)
            rs.append(rr)

        cap_prefix[p] = caps
        capr_prefix[p] = caprs
        r_sorted[p] = rs
        total_cap[p] = run_c

    return cap_prefix, capr_prefix, r_sorted, total_cap, suppliers_sorted


def build_delta_a(nIp, nP, a_IpP):
    delta_a = [[[0.0] * nP for _ in range(nIp)] for _ in range(nIp)]
    for i in range(nIp):
        ai = a_IpP[i]
        for j in range(nIp):
            aj = a_IpP[j]
            row = delta_a[i][j]
            for p in range(nP):
                row[p] = float(aj[p]) - float(ai[p])
    return delta_a


def build_sigma_order_sets(nIp, nK, sigma_IpK):
    sigma_order_by_k = []
    smaller_than = [[None] * nK for _ in range(nIp)]
    greater_than = [[None] * nK for _ in range(nIp)]

    for k in range(nK):
        order = sorted(range(nIp), key=lambda i: float(sigma_IpK[i][k]))
        sigma_order_by_k.append(order)
        pos = {i: t for t, i in enumerate(order)}
        for i in range(nIp):
            t = pos[i]
            smaller_than[i][k] = order[:t]
            greater_than[i][k] = order[t + 1:]

    return sigma_order_by_k, smaller_than, greater_than


def pivotal_supplier_pos_from_prefix(p_pos, A_p, cap_prefix):
    caps = cap_prefix[p_pos]
    if not caps:
        return None

    target = float(A_p)
    t = bisect.bisect_left(caps, target)

    if t >= len(caps):
        return None

    return t


def f_component_from_prefix(p_pos, A_p, cap_prefix, capr_prefix, r_sorted):
    A = float(A_p)
    if A <= 1e-12:
        return 0.0

    caps = cap_prefix[p_pos]
    caprs = capr_prefix[p_pos]
    rs = r_sorted[p_pos]

    t = bisect.bisect_left(caps, A)
    if t >= len(caps):
        return float("inf")

    prev_cap = 0.0 if t == 0 else float(caps[t - 1])
    prev_capr = 0.0 if t == 0 else float(caprs[t - 1])

    return prev_capr + float(rs[t]) * (A - prev_cap)


def f_total_from_prefix(Ap_list, nP, cap_prefix, capr_prefix, r_sorted):
    tot = 0.0
    for p in range(nP):
        c = f_component_from_prefix(p, Ap_list[p], cap_prefix, capr_prefix, r_sorted)
        if c == float("inf"):
            return float("inf")
        tot += float(c)
    return float(tot)


# ----------------------------
# beta3 alpha-breakpoint helpers
# ----------------------------

def alpha_candidates_for_beta3(i, jstar, k, nP, lam, a_IpP, Abar, cap_prefix, tol=1e-10):
    alphas = {0.0, 1.0}
    lk = float(lam[k])

    for p in range(nP):
        d = lk * (float(a_IpP[jstar][p]) - float(a_IpP[i][p]))
        if abs(d) <= tol:
            continue

        A0 = float(Abar[p])
        for B in cap_prefix[p]:
            alpha = 1.0 - (float(B) - A0) / d
            if -tol <= alpha <= 1.0 + tol:
                alpha = min(1.0, max(0.0, float(alpha)))
                alphas.add(round(alpha, 12))

    return sorted(alphas)


def shifted_cost_for_beta3_alpha(
    i, jstar, k, alpha, nP, lam, a_IpP, Abar, cap_prefix, capr_prefix, r_sorted
):
    shifted_A = [0.0] * nP
    lk = float(lam[k])

    for p in range(nP):
        shifted_A[p] = (
            float(Abar[p])
            + lk * (float(a_IpP[jstar][p]) - float(a_IpP[i][p])) * (1.0 - float(alpha))
        )

    return f_total_from_prefix(shifted_A, nP, cap_prefix, capr_prefix, r_sorted)


def compute_beta3_star_value(
    i, k, jstar, nP, lam, Pi_Ip, a_IpP, Abar, cap_prefix, capr_prefix, r_sorted
):
    candidates = alpha_candidates_for_beta3(
        i=i,
        jstar=jstar,
        k=k,
        nP=nP,
        lam=lam,
        a_IpP=a_IpP,
        Abar=Abar,
        cap_prefix=cap_prefix,
    )

    base_cost = f_total_from_prefix(Abar, nP, cap_prefix, capr_prefix, r_sorted)
    if base_cost == float("inf"):
        return 0.0

    best = float("-inf")
    lk = float(lam[k])
    pi_i = float(Pi_Ip[i])
    pi_j = float(Pi_Ip[jstar])

    for alpha in candidates:
        shifted_cost = shifted_cost_for_beta3_alpha(
            i=i,
            jstar=jstar,
            k=k,
            alpha=alpha,
            nP=nP,
            lam=lam,
            a_IpP=a_IpP,
            Abar=Abar,
            cap_prefix=cap_prefix,
            capr_prefix=capr_prefix,
            r_sorted=r_sorted,
        )
        if shifted_cost == float("inf"):
            continue

        revenue_part = lk * ((pi_i - pi_j) * float(alpha) + pi_j)
        value = float(revenue_part) - float(shifted_cost)
        if value > best:
            best = value

    if best == float("-inf"):
        return 0.0

    return max(0.0, float(best) + float(base_cost)-lk*pi_i)
    # -lk*pi_i

# ----------------------------
# Exact primal evaluation for incumbent lazy-check
# ----------------------------

def evaluate_selected_set_exact(
    x_inc,
    nIp, nK, nP,
    lam, Pi_Ip, a_IpP, sigma_IpK,
    cap_prefix, capr_prefix, r_sorted
):
    selected = [i for i in range(nIp) if x_inc[i] > 0.5]

    if len(selected) == 0:
        return {
            "selected": selected,
            "winner": [-1] * nK,
            "Abar": [0.0] * nP,
            "revenue": 0.0,
            "cost": 0.0,
            "value": 0.0,
        }

    winner = [-1] * nK
    for k in range(nK):
        best_i = -1
        best_sig = float("inf")
        for i in selected:
            s = float(sigma_IpK[i][k])
            if s < best_sig:
                best_sig = s
                best_i = i
        winner[k] = best_i

    revenue = 0.0
    for k in range(nK):
        revenue += float(lam[k]) * float(Pi_Ip[winner[k]])

    Abar = [0.0] * nP
    for p in range(nP):
        acc = 0.0
        for k in range(nK):
            acc += float(lam[k]) * float(a_IpP[winner[k]][p])
        Abar[p] = float(acc)

    cost = f_total_from_prefix(Abar, nP, cap_prefix, capr_prefix, r_sorted)
    value = float("-inf") if cost == float("inf") else float(revenue - cost)

    return {
        "selected": selected,
        "winner": winner,
        "Abar": Abar,
        "revenue": float(revenue),
        "cost": float(cost),
        "value": float(value),
    }


# ----------------------------
# Forward-Addition Greedy lower bound / MIP start
# ----------------------------

def forward_addition_greedy_mip_start(
    nIp, nK, nP, V, lam, Pi_Ip, a_arr, sigma_arr,
    cap_prefix, capr_prefix, r_sorted
):
    def evaluate(Spos):
        if not Spos:
            return 0.0, [-1] * nK, [0.0] * nP, 0.0

        winner = [-1] * nK
        for k in range(nK):
            best_i = -1
            best_sig = float("inf")
            for i in Spos:
                s = float(sigma_arr[i][k])
                if s < best_sig:
                    best_sig = s
                    best_i = i
            winner[k] = best_i

        rev = 0.0
        for k in range(nK):
            rev += float(lam[k]) * float(Pi_Ip[winner[k]])

        Abar = [0.0] * nP
        for p in range(nP):
            Abar[p] = sum(float(lam[k]) * float(a_arr[winner[k]][p]) for k in range(nK))

        cost = f_total_from_prefix(Abar, nP, cap_prefix, capr_prefix, r_sorted)
        if cost == float("inf"):
            return float("-inf"), winner, Abar, cost

        return float(rev - cost), winner, Abar, cost

    current_set = []
    current_val, current_winner, current_Abar, current_cost = evaluate(current_set)

    improved = True
    while improved and len(current_set) < min(V, nIp):
        improved = False
        best_add = None
        best_add_val = current_val
        best_add_winner = current_winner
        best_add_Abar = current_Abar
        best_add_cost = current_cost

        remaining = [j for j in range(nIp) if j not in current_set]

        for j in remaining:
            cand_set = current_set + [j]
            cand_val, cand_winner, cand_Abar, cand_cost = evaluate(cand_set)

            if cand_val > best_add_val + 1e-12:
                best_add = j
                best_add_val = cand_val
                best_add_winner = cand_winner
                best_add_Abar = cand_Abar
                best_add_cost = cand_cost

        if best_add is not None:
            current_set.append(best_add)
            current_val = best_add_val
            current_winner = best_add_winner
            current_Abar = best_add_Abar
            current_cost = best_add_cost
            improved = True

    selected_set = set(current_set)
    x_start = [1.0 if i in selected_set else 0.0 for i in range(nIp)]
    z_start = max(0.0, float(current_val))

    return current_set, x_start, z_start, current_winner, current_Abar, current_cost


# ----------------------------
# Closed-form beta computation for callbacks and final export
# ----------------------------

def compute_assignment_Abar_rhat(x_inc, nIp, nK, nP, lam, a_IpP, sigma_IpK, cap_prefix, r_sorted):
    """Return selected products, ybar, Abar, pivotal supplier positions, and rhat."""
    selected = [i for i in range(nIp) if float(x_inc[i]) > 0.5]

    ybar = [[0.0] * nK for _ in range(nIp)]
    if selected:
        for k in range(nK):
            best_i = -1
            best_sig = float("inf")
            for i in selected:
                sig = float(sigma_IpK[i][k])
                if sig < best_sig:
                    best_sig = sig
                    best_i = i
            if best_i >= 0:
                ybar[best_i][k] = 1.0

    Abar = [0.0] * nP
    for p in range(nP):
        acc = 0.0
        for k in range(nK):
            for i in range(nIp):
                if ybar[i][k] > 0.5:
                    acc += float(lam[k]) * float(a_IpP[i][p])
        Abar[p] = float(acc)

    shat = [None] * nP
    rhat = [0.0] * nP
    for p in range(nP):
        shat[p] = pivotal_supplier_pos_from_prefix(
            p_pos=p,
            A_p=Abar[p],
            cap_prefix=cap_prefix,
        )
        if shat[p] is None:
            rhat[p] = float("inf")
        else:
            rhat[p] = float(r_sorted[p][shat[p]])

    return selected, ybar, Abar, shat, rhat


def compute_closed_form_betas(
    x_inc,
    nIp, nK, nP, nS,
    lam, Pi_Ip, a_IpP, sigma_IpK, r_PS,
    cap_prefix, capr_prefix, r_sorted,
    smaller_than,
):
    """
    Compute all beta values directly from the closed-form expressions.

    No CPLEX LP subproblem is created or solved.

    beta3 follows the stated closed-form cases exactly:
      - beta3[i][k] = 0 if x_i = 0,
      - beta3[i][k] = 0 if x_i = 1 and y_i^k = 0,
      - beta3[i][k] is computed from the alpha-breakpoint formula only when
        y_i^k = 1.

    For the y_i^k = 1 case, jstar is the next-best selected product for customer k
    after i in sigma order. If no such selected product exists, beta3[i][k] is 0.
    """
    selected, ybar, Abar, shat, rhat = compute_assignment_Abar_rhat(
        x_inc=x_inc,
        nIp=nIp,
        nK=nK,
        nP=nP,
        lam=lam,
        a_IpP=a_IpP,
        sigma_IpK=sigma_IpK,
        cap_prefix=cap_prefix,
        r_sorted=r_sorted,
    )

    beta2 = [[0.0] * nK for _ in range(nIp)]
    beta3 = [[0.0] * nK for _ in range(nIp)]
    beta4 = [0.0] * nP
    beta5 = [[0.0] * nS for _ in range(nP)]
    beta1 = [0.0] * nK

    if len(selected) == 0:
        for i in range(nIp):
            for k in range(nK):
                beta2[i][k] = float(lam[k]) * float(Pi_Ip[i])
        return beta1, beta2, beta3, beta4, beta5, 0.0, [0.0] * nIp, [0.0] * nIp

    # beta4 and beta5 from the pivotal supplier formula.
    for p in range(nP):
        beta4[p] = float(rhat[p])
        for s in range(nS):
            beta5[p][s] = max(0.0, float(rhat[p]) - float(r_PS[p][s]))

    # beta2 formula.
    for i in range(nIp):
        xi = 1.0 if float(x_inc[i]) > 0.5 else 0.0
        for k in range(nK):
            beta2[i][k] = 0.0 if xi > 0.5 else float(lam[k]) * float(Pi_Ip[i])

    # beta3 formula.
    # Correct cases:
    #   1) beta3[i][k] = 0 if x_i = 0.
    #   2) beta3[i][k] = 0 if x_i = 1 and y_i^k = 0.
    #   3) beta3[i][k] is computed by the alpha-breakpoint formula only if y_i^k = 1.
    for i in range(nIp):
        xi = 1.0 if float(x_inc[i]) > 0.5 else 0.0

        for k in range(nK):
            if xi <= 0.5:
                beta3[i][k] = 0.0
                continue

            if ybar[i][k] <= 0.5:
                beta3[i][k] = 0.0
                continue

            # Since y_i^k = 1, product i is the current winner for customer k.
            # jstar is the next-best selected product if i is removed, i.e., the
            # selected product with the smallest sigma greater than sigma_i^k.
            sig_i = float(sigma_IpK[i][k])
            jstar = None
            best_sig_above = float("inf")

            for j in selected:
                if j == i:
                    continue
                sig_j = float(sigma_IpK[j][k])
                if sig_j > sig_i + 1e-12 and sig_j < best_sig_above:
                    best_sig_above = sig_j
                    jstar = j

            if jstar is None:
                beta3[i][k] = 0.0
                continue

            beta3[i][k] = compute_beta3_star_value(
                i=i,
                k=k,
                jstar=jstar,
                nP=nP,
                lam=lam,
                Pi_Ip=Pi_Ip,
                a_IpP=a_IpP,
                Abar=Abar,
                cap_prefix=cap_prefix,
                capr_prefix=capr_prefix,
                r_sorted=r_sorted,
            )

    # beta1 formula using the formula-computed beta3 values.
    for k in range(nK):
        best_k = float("-inf")

        for i in selected:
            term_cost = 0.0
            for p in range(nP):
                term_cost += float(a_IpP[i][p]) * float(rhat[p])

            first_term = float(lam[k]) * (float(Pi_Ip[i]) - float(term_cost))

            sum_beta3 = 0.0
            for j in smaller_than[i][k]:
                if float(x_inc[j]) > 0.5:
                    sum_beta3 += float(beta3[j][k])
                    
            cand = first_term - sum_beta3
            if cand > best_k:
                best_k = cand

        if best_k == float("-inf"):
            best_k = 0.0
        beta1[k] = max(0.0, float(best_k))

    coefLHS_opt = [0.0] * nIp
    coefLHS_feas = [0.0] * nIp
    for i in range(nIp):
        s2 = sum(beta2[i][k] for k in range(nK))
        s3 = sum(beta3[i][k] for k in range(nK))
        coefLHS_opt[i] = float(s3 - s2)
        coefLHS_feas[i] = float(-s3)

    return beta1, beta2, beta3, beta4, beta5, 0.0, coefLHS_opt, coefLHS_feas


def compute_closed_form_betas_and_cut(
    x_inc,
    nIp, nK, nP, nS,
    lam, Pi_Ip, a_IpP, sigma_IpK, r_PS, Cap_PS,
    cap_prefix, capr_prefix, r_sorted,
    smaller_than,
):
    beta1, beta2, beta3, beta4, beta5, _, coefLHS_opt, coefLHS_feas = compute_closed_form_betas(
        x_inc=x_inc,
        nIp=nIp,
        nK=nK,
        nP=nP,
        nS=nS,
        lam=lam,
        Pi_Ip=Pi_Ip,
        a_IpP=a_IpP,
        sigma_IpK=sigma_IpK,
        r_PS=r_PS,
        cap_prefix=cap_prefix,
        capr_prefix=capr_prefix,
        r_sorted=r_sorted,
        smaller_than=smaller_than,
    )

    RHS0 = 0.0
    RHS0 += sum(beta1[k] for k in range(nK))
    RHS0 += sum(beta3[i][k] for i in range(nIp) for k in range(nK))
    RHS0 += sum(float(Cap_PS[p][s]) * beta5[p][s] for p in range(nP) for s in range(nS))

    return beta1, beta2, beta3, beta4, beta5, float(RHS0), coefLHS_opt, coefLHS_feas


# ----------------------------
# Silent beta1 component collector for Excel export
# ----------------------------

def collect_beta1_values_with_components(
    beta1_val,
    beta3_val,
    x_inc,
    nIp,
    nK,
    nP,
    lam,
    Pi_Ip,
    a_IpP,
    sigma_IpK,
    cap_prefix,
    capr_prefix,
    r_sorted,
    smaller_than,
    Ip,
    K,
    iter_no,
    source="beta1 closed-form values with components",
    zero_tol=1e-12,
):
    """
    Silent collector for Excel export.

    This is the closed-form replacement for the previous
    collect_beta1_lp_values_with_components() helper.  It does not solve any LP;
    it only decomposes the already-computed beta1 values into the same reporting
    components used by the beta1 formula.
    """
    if beta1_val is None:
        return []

    selected = [i for i in range(nIp) if float(x_inc[i]) > 0.5]
    if len(selected) == 0:
        return []

    selected, ybar, Abar, shat, rhat = compute_assignment_Abar_rhat(
        x_inc=x_inc,
        nIp=nIp,
        nK=nK,
        nP=nP,
        lam=lam,
        a_IpP=a_IpP,
        sigma_IpK=sigma_IpK,
        cap_prefix=cap_prefix,
        r_sorted=r_sorted,
    )

    rows = []
    for k in range(nK):
        beta1_k = float(beta1_val[k])

        for i in selected:
            term_cost = 0.0
            for p in range(nP):
                term_cost += float(a_IpP[i][p]) * float(rhat[p])

            lambda_term = float(lam[k]) * (float(Pi_Ip[i]) - float(term_cost))

            sum_beta3 = 0.0
            if beta3_val is not None:
                for j in smaller_than[i][k]:
                    if float(x_inc[j]) > 0.5:
                        sum_beta3 += float(beta3_val[j][k])

            candidate = lambda_term - sum_beta3

            if (
                abs(beta1_k) > zero_tol
                or abs(lambda_term) > zero_tol
                or abs(sum_beta3) > zero_tol
                or abs(candidate) > zero_tol
            ):
                rows.append({
                    "iteration": iter_no,
                    "source": source,
                    "k": K[k],
                    "i": Ip[i],
                    "beta1_value": beta1_k,
                    "lambda_term": lambda_term,
                    "sum_beta3_before_i": sum_beta3,
                    "lambda_term_minus_sum_beta3": candidate,
                    "selected_products": str([Ip[t] for t in selected]),
                    "Abar": str([round(v, 10) for v in Abar]),
                    "shat_sorted_0based": str(shat),
                    "rhat": str([round(v, 10) if v != float("inf") else "inf" for v in rhat]),
                })

    return rows


# Backward-compatible alias for any remaining old export/callback references.
# The closed-form version does not solve an LP; this alias prevents NameError
# if an older section still calls the previous helper name.
collect_beta1_lp_values_with_components = collect_beta1_values_with_components


# ----------------------------
# Lazy callback
# ----------------------------

class BD2LazyCallback_ClosedForm(LazyConstraintCallback):
    def __call__(self):
        x_vals = self.get_values(self.master_x_idx_list)
        x_inc = [1.0 if x_vals[i] > 0.5 else 0.0 for i in range(self.nIp)]
        iter_no = self.iters + 1

        if self.force_at_least_one and sum(x_inc) < 1.0:
            ind = list(self.master_x_idx_list)
            val = [1.0] * len(ind)
            self.add(cplex.SparsePair(ind=ind, val=val), "G", 1.0)
            return

        key = tuple(int(v) for v in x_inc)

        if not hasattr(self, "_incumbent_cut_cache"):
            self._incumbent_cut_cache = {}

        if not hasattr(self, "_incumbent_primal_cache"):
            self._incumbent_primal_cache = {}

        if key in self._incumbent_primal_cache:
            primal_eval = self._incumbent_primal_cache[key]
        else:
            primal_eval = evaluate_selected_set_exact(
                x_inc=x_inc,
                nIp=self.nIp,
                nK=self.nK,
                nP=self.nP,
                lam=self.lam,
                Pi_Ip=self.Pi_Ip,
                a_IpP=self.a_IpP,
                sigma_IpK=self.sigma_IpK,
                cap_prefix=self.cap_prefix,
                capr_prefix=self.capr_prefix,
                r_sorted=self.r_sorted,
            )
            self._incumbent_primal_cache[key] = primal_eval

        if key in self._incumbent_cut_cache:
            beta1_val, beta2_val, beta3_val, beta4_val, beta5_val, RHS0, coefLHS_opt, coefLHS_feas = self._incumbent_cut_cache[key]
        else:
            beta1_val, beta2_val, beta3_val, beta4_val, beta5_val, RHS0, coefLHS_opt, coefLHS_feas = compute_closed_form_betas_and_cut(
                x_inc=x_inc,
                nIp=self.nIp,
                nK=self.nK,
                nP=self.nP,
                nS=self.nS,
                lam=self.lam,
                Pi_Ip=self.Pi_Ip,
                a_IpP=self.a_IpP,
                sigma_IpK=self.sigma_IpK,
                r_PS=self.r_PS,
                Cap_PS=self.Cap_PS,
                cap_prefix=self.cap_prefix,
                capr_prefix=self.capr_prefix,
                r_sorted=self.r_sorted,
                smaller_than=self.smaller_than,
            )
            self._incumbent_cut_cache[key] = (
                beta1_val, beta2_val, beta3_val, beta4_val, beta5_val,
                RHS0, coefLHS_opt, coefLHS_feas,
            )

        rows_beta1 = collect_beta1_values_with_components(
            beta1_val=beta1_val,
            beta3_val=beta3_val,
            x_inc=x_inc,
            nIp=self.nIp,
            nK=self.nK,
            nP=self.nP,
            lam=self.lam,
            Pi_Ip=self.Pi_Ip,
            a_IpP=self.a_IpP,
            sigma_IpK=self.sigma_IpK,
            cap_prefix=self.cap_prefix,
            capr_prefix=self.capr_prefix,
            r_sorted=self.r_sorted,
            smaller_than=self.smaller_than,
            Ip=self.Ip,
            K=self.K,
            iter_no=iter_no,
            source="BD2-LAZY beta1 closed-form",
        )
        if hasattr(self, "beta1_component_rows"):
            self.beta1_component_rows.extend(rows_beta1)

        best_ub = self.get_best_objective_value()
        if best_ub < self.UB:
            self.UB = best_ub

        feas_val = RHS0 + sum(coefLHS_feas[i] * float(x_inc[i]) for i in range(self.nIp))
        if feas_val < -self.cut_eps:
            ind = []
            val = []
            for i in range(self.nIp):
                c = coefLHS_feas[i]
                if abs(c) > self.zero_tol:
                    ind.append(self.master_x_idx_list[i])
                    val.append(c)
            self.add(cplex.SparsePair(ind=ind, val=val), "G", -RHS0)

            self.iters += 1
            self.feascuts_added += 1
            if self.print_every > 0 and self.iters % self.print_every == 0:
                print(
                    f"[BD2-LAZY-closed-form(feas-cut)] it={self.iters}  "
                    f"LB={self.LB:.6f}  UB={self.UB:.6f}  gap={self.UB-self.LB:.6f}"
                )
            return

        z_val = float(self.get_values(self.master_z_idx))
        primal_val = float(primal_eval["value"])

        if primal_val > self.LB:
            self.LB = primal_val

        if z_val <= primal_val + self.cut_eps:
            self.iters += 1
            if self.print_every > 0 and self.iters % self.print_every == 0:
                print(
                    f"[BD2-LAZY-closed-form(skip)] it={self.iters}  "
                    f"LB={self.LB:.6f}  UB={self.UB:.6f}  gap={self.UB-self.LB:.6f}"
                )
            return

        lhs = z_val + sum(coefLHS_opt[i] * float(x_inc[i]) for i in range(self.nIp))

        if lhs > RHS0 + self.cut_eps:
            ind = [self.master_z_idx]
            val = [1.0]
            for i in range(self.nIp):
                c = coefLHS_opt[i]
                if abs(c) > self.zero_tol:
                    ind.append(self.master_x_idx_list[i])
                    val.append(c)
            self.add(cplex.SparsePair(ind=ind, val=val), "L", RHS0)
            self.optcuts_added += 1

        self.iters += 1
        if self.print_every > 0 and self.iters % self.print_every == 0:
            print(
                f"[BD2-LAZY-closed-form(opt-cut)] it={self.iters}  "
                f"LB={self.LB:.6f}  UB={self.UB:.6f}  gap={self.UB-self.LB:.6f}"
            )



# ----------------------------
# Console final-result printing helpers
# ----------------------------

def _safe_float(v):
    try:
        return float(v)
    except Exception:
        return v


def _print_vector_preview(name, labels, values, max_rows=20, full=False, digits=6):
    """Print a labeled vector without flooding the console on large instances."""
    n = len(values)
    print(f"\n{name} (count={n}):")
    limit = n if full else min(n, max_rows)
    for idx in range(limit):
        val = values[idx]
        if isinstance(val, float):
            val = round(val, digits)
        print(f"  {labels[idx]}: {val}")
    if not full and n > limit:
        print(f"  ... {n - limit} more rows not printed. Set PRINT_FULL_FINAL_TABLES=True to print all.")


def _print_matrix_preview(name, row_labels, col_labels, matrix, max_rows=20, full=False, digits=6):
    """Print nonzero entries of a matrix, with a preview limit for large instances."""
    entries = []
    for i, row in enumerate(matrix):
        for j, val in enumerate(row):
            fv = _safe_float(val)
            if isinstance(fv, float) and abs(fv) <= 1e-12:
                continue
            entries.append((row_labels[i], col_labels[j], fv))

    print(f"\n{name} nonzero entries (count={len(entries)}):")
    limit = len(entries) if full else min(len(entries), max_rows)
    for idx in range(limit):
        r, c, val = entries[idx]
        if isinstance(val, float):
            val = round(val, digits)
        print(f"  ({r}, {c}): {val}")
    if len(entries) == 0:
        print("  all entries are zero")
    elif not full and len(entries) > limit:
        print(f"  ... {len(entries) - limit} more nonzero rows not printed. Set PRINT_FULL_FINAL_TABLES=True to print all.")


def print_final_results_to_console(
    *,
    sample_id,
    n_value,
    P_value,
    K_value,
    V_value,
    S_value,
    status,
    objective,
    z_value,
    best_bound,
    mip_gap_percent,
    selected_final_labels,
    primal_eval_final,
    RHS0_final,
    coefLHS_final,
    coefLHS_feas_final,
    beta1_val,
    beta2_val,
    beta3_val,
    beta4_val,
    beta5_val,
    x_inc_final,
    Abar_final,
    rhat_sp_final,
    shat_final,
    Ip,
    K,
    P,
    S,
    total_time_sec,
    greedy_time_sec,
    optcuts_added,
    feascuts_added,
):
    """Print the final solution and beta values computed by closed-form formulas."""
    if not PRINT_FINAL_RESULTS_TO_CONSOLE:
        return

    full = bool(PRINT_FULL_FINAL_TABLES)
    max_rows = int(FINAL_PREVIEW_ROWS)

    print("\n====================================================")
    print("FINAL CLOSED-FORM BD2 RESULT")
    print("====================================================")
    print(f"sample_id: {sample_id}")
    print(f"n={n_value}, P={P_value}, K={K_value}, V={V_value}, S={S_value}")
    print(f"Status: {status}")
    print(f"Objective: {objective}")
    print(f"z: {z_value}")
    print(f"Best bound: {best_bound}")
    print(f"MIP gap (%): {mip_gap_percent}")
    print(f"RHS0: {RHS0_final}")
    print(f"Selected products: {selected_final_labels}")
    print(f"Greedy time (sec): {greedy_time_sec}")
    print(f"Total time (sec): {total_time_sec}")
    print(f"Lazy optimality cuts: {optcuts_added}")
    print(f"Lazy feasibility cuts: {feascuts_added}")

    if primal_eval_final is not None:
        print("\nFinal primal evaluation:")
        print(f"  revenue: {primal_eval_final.get('revenue')}")
        print(f"  cost: {primal_eval_final.get('cost')}")
        print(f"  value = revenue - cost: {primal_eval_final.get('value')}")
        winners = primal_eval_final.get("winner", [])
        if winners:
            winner_labels = [None if w < 0 else Ip[w] for w in winners]
            _print_vector_preview("Winner product by customer k", K, winner_labels, max_rows=max_rows, full=full, digits=6)

    _print_vector_preview("x_final", Ip, x_inc_final, max_rows=max_rows, full=full, digits=6)
    _print_vector_preview("Abar_final", P, Abar_final, max_rows=max_rows, full=full, digits=6)
    _print_vector_preview("rhat_sp_final", P, rhat_sp_final, max_rows=max_rows, full=full, digits=6)
    _print_vector_preview("shat_final sorted-position 0-based", P, shat_final, max_rows=max_rows, full=full, digits=6)

    _print_vector_preview("beta1", K, beta1_val, max_rows=max_rows, full=full, digits=6)
    _print_matrix_preview("beta2", Ip, K, beta2_val, max_rows=max_rows, full=full, digits=6)
    _print_matrix_preview("beta3", Ip, K, beta3_val, max_rows=max_rows, full=full, digits=6)
    _print_vector_preview("beta4", P, beta4_val, max_rows=max_rows, full=full, digits=6)
    _print_matrix_preview("beta5", P, S, beta5_val, max_rows=max_rows, full=full, digits=6)
    _print_vector_preview("coefLHS_opt", Ip, coefLHS_final, max_rows=max_rows, full=full, digits=6)
    _print_vector_preview("coefLHS_feas", Ip, coefLHS_feas_final, max_rows=max_rows, full=full, digits=6)

    print("====================================================\n")

# ----------------------------
# Main solver
# ----------------------------

def solve_BD2_manual_closed_form(address, n_value, P_value, K_value, V_value, S_value, sample_id):
    total_start_time = time.perf_counter()
    start_datetime = datetime.now()

    print("\n====================================================")
    print(f"Solving BD2 closed-form sample {sample_id}")
    print(f"n={n_value}, P={P_value}, K={K_value}, V={V_value}, S={S_value}")
    print(f"Input file: {address}")
    print(f"Time limit for this run = {TIME_LIMIT} seconds")
    print("====================================================\n")

    lam_data   = pd.read_excel(address, sheet_name='Lambda^k')
    Pi_data    = pd.read_excel(address, sheet_name='Pi_i')
    a_data     = pd.read_excel(address, sheet_name='a_ip')
    sigma_data = pd.read_excel(address, sheet_name='Sigma_i^k')
    r_data     = pd.read_excel(address, sheet_name='r_p^s')
    Cap_data   = pd.read_excel(address, sheet_name='Cap_p^s')

    I  = list(range(0, n_value + 1))
    Ip = list(range(1, n_value + 1))
    K  = list(range(1, K_value + 1))
    P  = list(range(1, P_value + 1))

    S_original = list(range(1, S_value + 1))
    artificial_supplier = S_value + 1
    S = S_original + [artificial_supplier]

    V = V_value

    nIp = len(Ip)
    nK = len(K)
    nP = len(P)
    nS = len(S)

    lam_df   = pd.DataFrame(lam_data)
    Pi_df    = pd.DataFrame(Pi_data)
    a_df     = pd.DataFrame(a_data)
    sigma_df = pd.DataFrame(sigma_data)
    r_df     = pd.DataFrame(r_data)
    Cap_df   = pd.DataFrame(Cap_data)

    # ========================================================
    # AUTOMATIC SCALING RULES
    # ========================================================
    # Pi scaling:
    # P = 1000  -> Pi unchanged
    # P = 2000  -> Pi multiplied by 2
    # P = 4000  -> Pi multiplied by 4
    # P = 8000  -> Pi multiplied by 8
    # P = 16000 -> Pi multiplied by 16
    Pi_multiplier = 2 ** math.log2(P_value / 1000)
    
    # Capacity scaling for number of suppliers:
    # S = 2 -> supplier multiplier = 1
    # S = 4 -> supplier multiplier = 1 / 1.5
    supplier_multiplier = 1 / (1.5 ** math.log2(S_value / 2))
    
    # K = 200 -> K multiplier = 1
    # K = 400 -> K multiplier = 2
    K_multiplier = K_value / 200
    
    
    # Final capacity multiplier combines both effects
    Cap_multiplier = supplier_multiplier * K_multiplier

    lam = [float(lam_df.loc[0, k]) for k in K]
    Pi_I = [float(Pi_df.loc[0, i]) * Pi_multiplier for i in I]
    Pi_Ip = [float(Pi_df.loc[0, i]) * Pi_multiplier for i in Ip]
    a_IpP = [[float(a_df.loc[i, p]) for p in P] for i in Ip]
    sigma_IpK = [[float(sigma_df.loc[i, k]) for k in K] for i in Ip]

    # Real suppliers are read from Excel; one artificial last supplier is appended.
    r_PS = [
        [float(r_df.loc[p - 1, s]) for s in S_original] + [ARTIFICIAL_SUPPLIER_COST]
        for p in P
    ]
    Cap_PS = [
        [float(Cap_df.loc[p - 1, s]) * Cap_multiplier for s in S_original] + [ARTIFICIAL_SUPPLIER_CAP]
        for p in P
    ]

    cap_prefix, capr_prefix, r_sorted, total_cap, suppliers_sorted = build_supplier_prefix(
        nP, nS, r_PS, Cap_PS
    )

    delta_a = build_delta_a(nIp, nP, a_IpP)
    sigma_order_by_k, smaller_than, greater_than = build_sigma_order_sets(nIp, nK, sigma_IpK)

    master = cplex.Cplex()
    master.objective.set_sense(master.objective.sense.maximize)
    master.set_log_stream(sys.stdout)
    master.set_results_stream(sys.stdout)
    master.set_error_stream(None)
    master.set_warning_stream(None)

    master.parameters.mip.display.set(4)
    master.parameters.threads.set(1)
    master.parameters.timelimit.set(TIME_LIMIT)
    master.parameters.threads.set(THREADS)
    master.parameters.mip.strategy.heuristicfreq.set(-1)
    master.parameters.preprocessing.presolve.set(1)
    master.parameters.mip.cuts.mircut.set(-1)
    master.parameters.mip.cuts.implied.set(-1)
    master.parameters.mip.cuts.gomory.set(-1)

    z_ub = sum(float(lam[k]) * max(Pi_Ip) for k in range(nK)) if nIp > 0 else 0.0

    z_idx = master.variables.get_num()
    master.variables.add(
        names=["z"],
        lb=[0.0],
        ub=[float(z_ub)],
        obj=[1.0],
    )

    x_idx_list = []
    for i_label in Ip:
        x_idx_list.append(master.variables.get_num())
        master.variables.add(names=[f"x_{i_label}"], types=["B"], obj=[0.0])

    master.linear_constraints.add(
        lin_expr=[cplex.SparsePair(ind=x_idx_list, val=[1.0] * nIp)],
        senses=["L"], rhs=[float(V)]
    )

    greedy_start_time = time.perf_counter()
    S_greedy_pos, x_start, z_start, win_g, Abar_g, cost_g = forward_addition_greedy_mip_start(
        nIp=nIp, nK=nK, nP=nP, V=V,
        lam=lam, Pi_Ip=Pi_Ip, a_arr=a_IpP, sigma_arr=sigma_IpK,
        cap_prefix=cap_prefix, capr_prefix=capr_prefix, r_sorted=r_sorted,
    )
    greedy_end_time = time.perf_counter()
    greedy_time_sec = greedy_end_time - greedy_start_time

    S_greedy_labels = [Ip[i] for i in S_greedy_pos]

    start_ind = [z_idx] + x_idx_list
    start_val = [float(z_start)] + [float(x_start[i]) for i in range(nIp)]
    try:
        master.MIP_starts.add(
            cplex.SparsePair(ind=start_ind, val=start_val),
            master.MIP_starts.effort_level.solve_MIP,
        )
        print("\n[Forward-Addition Greedy MIP start added]")
        print("Greedy MIP start computed and added. Details are exported to Excel.")
    except Exception as e:
        print("WARNING: Could not add MIP start:", repr(e))

    # No CPLEX LP subproblem is built in this version.
    # All beta values are computed directly in the lazy callback from formulas.

    lcb = master.register_callback(BD2LazyCallback_ClosedForm)
    lcb.nIp = nIp
    lcb.nK = nK
    lcb.nP = nP
    lcb.nS = nS
    lcb.master_z_idx = z_idx
    lcb.master_x_idx_list = x_idx_list
    lcb.cut_eps = 1e-6
    lcb.zero_tol = 1e-12
    lcb.force_at_least_one = False

    lcb.lam = lam
    lcb.Pi_Ip = Pi_Ip
    lcb.a_IpP = a_IpP
    lcb.sigma_IpK = sigma_IpK
    lcb.r_PS = r_PS
    lcb.Cap_PS = Cap_PS

    lcb.cap_prefix = cap_prefix
    lcb.capr_prefix = capr_prefix
    lcb.r_sorted = r_sorted
    lcb.suppliers_sorted = suppliers_sorted

    lcb.sigma_order_by_k = sigma_order_by_k
    lcb.smaller_than = smaller_than
    lcb.greater_than = greater_than

    lcb.Ip = Ip
    lcb.K = K


    lcb.LB = float(z_start)
    lcb.UB = 1e100
    lcb.iters = 0
    lcb.print_every = 10
    lcb.optcuts_added = 0
    lcb.feascuts_added = 0
    lcb.beta1_component_rows = []

    master.solve()

    total_end_time = time.perf_counter()
    total_time_sec = total_end_time - total_start_time

    print("\nSolve finished.")
    print("Status:", master.solution.get_status_string())
    print("Detailed final values are exported to Excel only.")
    print("Total solving time (sec):", total_time_sec)

    best_bound = None
    mip_gap_percent = None
    try:
        best_bound = master.solution.MIP.get_best_objective()
    except Exception:
        best_bound = None
    try:
        mip_gap_percent = 100.0 * master.solution.MIP.get_mip_relative_gap()
    except Exception:
        mip_gap_percent = None

    x_final_vals = master.solution.get_values(x_idx_list)
    x_inc_final = [1.0 if v > 0.5 else 0.0 for v in x_final_vals]
    selected_final_labels = [Ip[i] for i, v in enumerate(x_inc_final) if v > 0.5]

    out_final = compute_closed_form_betas_and_cut(
        x_inc=x_inc_final,
        nIp=nIp,
        nK=nK,
        nP=nP,
        nS=nS,
        lam=lam,
        Pi_Ip=Pi_Ip,
        a_IpP=a_IpP,
        sigma_IpK=sigma_IpK,
        r_PS=r_PS,
        Cap_PS=Cap_PS,
        cap_prefix=cap_prefix,
        capr_prefix=capr_prefix,
        r_sorted=r_sorted,
        smaller_than=smaller_than,
    )

    if out_final is not None:
        beta1_val, beta2_val, beta3_val, beta4_val, beta5_val, RHS0_final, coefLHS_final, coefLHS_feas_final = out_final

        final_beta1_component_rows = collect_beta1_values_with_components(
            beta1_val=beta1_val,
            beta3_val=beta3_val,
            x_inc=x_inc_final,
            nIp=nIp,
            nK=nK,
            nP=nP,
            lam=lam,
            Pi_Ip=Pi_Ip,
            a_IpP=a_IpP,
            sigma_IpK=sigma_IpK,
            cap_prefix=cap_prefix,
            capr_prefix=capr_prefix,
            r_sorted=r_sorted,
            smaller_than=smaller_than,
            Ip=Ip,
            K=K,
            iter_no="final",
            source="BD2-FINAL beta1 closed-form",
        )

        selected_final = [i for i in range(nIp) if x_inc_final[i] > 0.5]
        ybar_final_only_Ip = [[0.0] * nK for _ in range(nIp)]
        for k in range(nK):
            best_i = -1
            best_sig = float("inf")
            for i in selected_final:
                s = float(sigma_IpK[i][k])
                if s < best_sig:
                    best_sig = s
                    best_i = i
            if best_i >= 0:
                ybar_final_only_Ip[best_i][k] = 1.0

        y0k_final = [0.0] * nK
        for k in range(nK):
            y0k_final[k] = 1.0 - sum(ybar_final_only_Ip[i][k] for i in range(nIp))

        Abar_final = [0.0] * nP
        for p in range(nP):
            acc = 0.0
            for k in range(nK):
                for i in range(nIp):
                    if ybar_final_only_Ip[i][k] > 0.5:
                        acc += float(lam[k]) * float(a_IpP[i][p])
            Abar_final[p] = float(acc)

        shat_final = [None] * nP
        rhat_sp_final = [0.0] * nP

        for p in range(nP):
            shat_final[p] = pivotal_supplier_pos_from_prefix(
                p_pos=p,
                A_p=Abar_final[p],
                cap_prefix=cap_prefix,
            )
            if shat_final[p] is None:
                rhat_sp_final[p] = float("inf")
            else:
                rhat_sp_final[p] = float(r_sorted[p][shat_final[p]])

        primal_eval_final = evaluate_selected_set_exact(
            x_inc=x_inc_final,
            nIp=nIp,
            nK=nK,
            nP=nP,
            lam=lam,
            Pi_Ip=Pi_Ip,
            a_IpP=a_IpP,
            sigma_IpK=sigma_IpK,
            cap_prefix=cap_prefix,
            capr_prefix=capr_prefix,
            r_sorted=r_sorted,
        )

        print_final_results_to_console(
            sample_id=sample_id,
            n_value=n_value,
            P_value=P_value,
            K_value=K_value,
            V_value=V_value,
            S_value=S_value,
            status=master.solution.get_status_string(),
            objective=master.solution.get_objective_value(),
            z_value=master.solution.get_values("z"),
            best_bound=best_bound,
            mip_gap_percent=mip_gap_percent,
            selected_final_labels=selected_final_labels,
            primal_eval_final=primal_eval_final,
            RHS0_final=RHS0_final,
            coefLHS_final=coefLHS_final,
            coefLHS_feas_final=coefLHS_feas_final,
            beta1_val=beta1_val,
            beta2_val=beta2_val,
            beta3_val=beta3_val,
            beta4_val=beta4_val,
            beta5_val=beta5_val,
            x_inc_final=x_inc_final,
            Abar_final=Abar_final,
            rhat_sp_final=rhat_sp_final,
            shat_final=shat_final,
            Ip=Ip,
            K=K,
            P=P,
            S=S,
            total_time_sec=total_time_sec,
            greedy_time_sec=greedy_time_sec,
            optcuts_added=lcb.optcuts_added,
            feascuts_added=lcb.feascuts_added,
        )

        df_x = pd.DataFrame({
            "i": [Ip[i] for i in range(nIp)],
            "x_value": x_inc_final,
        })

        y_rows_all = []
        for k in range(nK):
            y_rows_all.append({"i": I[0], "k": K[k], "y_ik": y0k_final[k]})
        for i in range(nIp):
            for k in range(nK):
                y_rows_all.append({"i": Ip[i], "k": K[k], "y_ik": ybar_final_only_Ip[i][k]})

        df_yik = pd.DataFrame(y_rows_all)
        df_ybar_final = pd.DataFrame(y_rows_all)

        df_Abar = pd.DataFrame({"p": [P[p] for p in range(nP)], "Abar": Abar_final})

        df_rhat_sp = pd.DataFrame({
            "p": [P[p] for p in range(nP)],
            "s_hat_pos_sorted_0based": shat_final,
            "s_hat_pos_sorted_1based": [None if shat_final[p] is None else shat_final[p] + 1 for p in range(nP)],
            "rhat_sp": rhat_sp_final,
        })

        # Build beta1 summary with the two formula components used for each beta1_k.
        # For each k, beta1_k is associated with the selected product i that maximizes:
        #   lambda^k * (pi_i - sum_p a_ip rhat_p) - sum_{j before i} beta3_jk.
        df_beta1_components_tmp = pd.DataFrame(final_beta1_component_rows)
        beta1_summary_rows = []
        for k_idx in range(nK):
            k_label = K[k_idx]
            beta1_k = float(beta1_val[k_idx])

            if not df_beta1_components_tmp.empty:
                rows_k = df_beta1_components_tmp[df_beta1_components_tmp["k"] == k_label]
            else:
                rows_k = pd.DataFrame()

            if not rows_k.empty:
                best_row = rows_k.loc[rows_k["lambda_term_minus_sum_beta3"].idxmax()]
                beta1_summary_rows.append({
                    "k": k_label,
                    "beta1": beta1_k,
                    "argmax_i": best_row["i"],
                    "lambda_term": best_row["lambda_term"],
                    "sum_beta3": best_row["sum_beta3_before_i"],
                    "lambda_term_minus_sum_beta3": best_row["lambda_term_minus_sum_beta3"],
                })
            else:
                beta1_summary_rows.append({
                    "k": k_label,
                    "beta1": beta1_k,
                    "argmax_i": None,
                    "lambda_term": None,
                    "sum_beta3": None,
                    "lambda_term_minus_sum_beta3": None,
                })

        df_beta1 = pd.DataFrame(beta1_summary_rows)

        lazy_beta1_rows = []
        if hasattr(lcb, "beta1_component_rows"):
            lazy_beta1_rows = list(lcb.beta1_component_rows)

        all_beta1_component_rows = lazy_beta1_rows + final_beta1_component_rows
        df_beta1_components = pd.DataFrame(all_beta1_component_rows)

        if df_beta1_components.empty:
            df_beta1_components = pd.DataFrame(columns=[
                "iteration",
                "source",
                "k",
                "i",
                "beta1_value",
                "lambda_term",
                "sum_beta3_before_i",
                "lambda_term_minus_sum_beta3",
                "selected_products",
                "Abar",
                "shat_sorted_0based",
                "rhat",
            ])

        beta2_rows = []
        for i in range(nIp):
            for k in range(nK):
                beta2_rows.append({"i": Ip[i], "k": K[k], "beta2": beta2_val[i][k]})
        df_beta2 = pd.DataFrame(beta2_rows)

        beta3_rows = []
        beta3_case_violations = 0
        for i in range(nIp):
            for k in range(nK):
                yik = float(ybar_final_only_Ip[i][k])
                xik = float(x_inc_final[i])
                must_be_zero = (xik <= 0.5) or (xik > 0.5 and yik <= 0.5)
                beta3_value = float(beta3_val[i][k])
                violation = bool(must_be_zero and abs(beta3_value) > 1e-9)
                if violation:
                    beta3_case_violations += 1
                beta3_rows.append({
                    "i": Ip[i],
                    "k": K[k],
                    "x_i": xik,
                    "y_ik": yik,
                    "beta3": beta3_value,
                    "must_be_zero_by_case": must_be_zero,
                    "case_violation": violation,
                })
        df_beta3 = pd.DataFrame(beta3_rows)

        df_beta4 = pd.DataFrame({"p": [P[p] for p in range(nP)], "beta4": beta4_val})

        beta5_rows = []
        for p in range(nP):
            for s in range(nS):
                beta5_rows.append({"p": P[p], "s": S[s], "beta5": beta5_val[p][s]})
        df_beta5 = pd.DataFrame(beta5_rows)

        # ----------------------------------------------------
        # Complete beta export tables
        # ----------------------------------------------------
        # Long format: every beta value in one sheet.
        # This is useful for filtering/pivoting all beta variables together.
        beta_all_rows = []

        for k in range(nK):
            beta_all_rows.append({
                "beta": "beta1",
                "index_1_name": "k",
                "index_1": K[k],
                "index_2_name": None,
                "index_2": None,
                "value": beta1_val[k],
            })

        for i in range(nIp):
            for k in range(nK):
                beta_all_rows.append({
                    "beta": "beta2",
                    "index_1_name": "i",
                    "index_1": Ip[i],
                    "index_2_name": "k",
                    "index_2": K[k],
                    "value": beta2_val[i][k],
                })

        for i in range(nIp):
            for k in range(nK):
                beta_all_rows.append({
                    "beta": "beta3",
                    "index_1_name": "i",
                    "index_1": Ip[i],
                    "index_2_name": "k",
                    "index_2": K[k],
                    "value": beta3_val[i][k],
                })

        for p in range(nP):
            beta_all_rows.append({
                "beta": "beta4",
                "index_1_name": "p",
                "index_1": P[p],
                "index_2_name": None,
                "index_2": None,
                "value": beta4_val[p],
            })

        for p in range(nP):
            for s in range(nS):
                beta_all_rows.append({
                    "beta": "beta5",
                    "index_1_name": "p",
                    "index_1": P[p],
                    "index_2_name": "s",
                    "index_2": S[s],
                    "value": beta5_val[p][s],
                })

        df_beta_all_long = pd.DataFrame(beta_all_rows)

        # Wide matrix format: easier to inspect beta2/beta3/beta5 visually.
        df_beta2_matrix = pd.DataFrame(beta2_val, index=Ip, columns=K).reset_index()
        df_beta2_matrix = df_beta2_matrix.rename(columns={"index": "i"})

        df_beta3_matrix = pd.DataFrame(beta3_val, index=Ip, columns=K).reset_index()
        df_beta3_matrix = df_beta3_matrix.rename(columns={"index": "i"})

        df_beta5_matrix = pd.DataFrame(beta5_val, index=P, columns=S).reset_index()
        df_beta5_matrix = df_beta5_matrix.rename(columns={"index": "p"})

        df_beta1_full = pd.DataFrame({"k": K, "beta1": beta1_val})
        df_beta4_full = pd.DataFrame({"p": P, "beta4": beta4_val})

        df_summary = pd.DataFrame({
            "metric": ["objective", "z", "RHS0", "greedy_time_sec", "total_time_sec", "greedy_z_start"],
            "value": [
                master.solution.get_objective_value(),
                master.solution.get_values("z"),
                RHS0_final,
                greedy_time_sec,
                total_time_sec,
                z_start,
            ],
        })

        output_dir = os.path.dirname(address)
        output_file = os.path.join(
            output_dir,
            f"BD2_closed_form_beta_values_n{n_value}_P{P_value}_K{K_value}_V{V_value}_S{S_value}_sample{sample_id}.xlsx"
        )

        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            df_summary.to_excel(writer, sheet_name="summary", index=False)
            df_x.to_excel(writer, sheet_name="x_final", index=False)
            df_yik.to_excel(writer, sheet_name="y_ik", index=False)
            df_ybar_final.to_excel(writer, sheet_name="ybar_final", index=False)
            df_Abar.to_excel(writer, sheet_name="Abar_final", index=False)
            df_rhat_sp.to_excel(writer, sheet_name="rhat_sp", index=False)
            # Individual complete beta tables.
            df_beta1.to_excel(writer, sheet_name="beta1", index=False)
            df_beta1_full.to_excel(writer, sheet_name="beta1_all", index=False)
            df_beta1_components.to_excel(writer, sheet_name="beta1_components", index=False)
            df_beta2.to_excel(writer, sheet_name="beta2", index=False)
            df_beta3.to_excel(writer, sheet_name="beta3", index=False)
            df_beta4.to_excel(writer, sheet_name="beta4", index=False)
            df_beta4_full.to_excel(writer, sheet_name="beta4_all", index=False)
            df_beta5.to_excel(writer, sheet_name="beta5", index=False)

            # Complete all-in-one and matrix beta sheets.
            df_beta_all_long.to_excel(writer, sheet_name="all_betas_long", index=False)
            df_beta2_matrix.to_excel(writer, sheet_name="beta2_matrix", index=False)
            df_beta3_matrix.to_excel(writer, sheet_name="beta3_matrix", index=False)
            df_beta5_matrix.to_excel(writer, sheet_name="beta5_matrix", index=False)

        print(f"Final results and all beta values exported to: {output_file}")
        print("Console variable-value printing is disabled; inspect the Excel sheets for x, y, Abar, rhat, beta1-beta5, and cut coefficients.")
    else:
        print("WARNING: Could not compute final closed-form beta values, so beta values were not exported.")
        RHS0_final = None
        output_file = None

    end_datetime = datetime.now()

    return {
        "sample_id": sample_id,
        "n": n_value,
        "P": P_value,
        "K": K_value,
        "V": V_value,
        "S": S_value,
        "artificial_supplier": artificial_supplier,
        "artificial_supplier_cost": ARTIFICIAL_SUPPLIER_COST,
        "artificial_supplier_cap": ARTIFICIAL_SUPPLIER_CAP,
        "status": master.solution.get_status_string(),
        "objective": master.solution.get_objective_value(),
        "z": master.solution.get_values("z"),
        "best_bound": best_bound,
        "mip_gap_percent": mip_gap_percent,
        "RHS0": RHS0_final,
        "greedy_z_start": z_start,
        "greedy_time_sec": greedy_time_sec,
        "total_time_sec": total_time_sec,
        "user_cuts_added": 0,
        "user_optimality_cuts": 0,
        "user_feasibility_cuts": 0,
        "lazy_optimality_cuts": lcb.optcuts_added,
        "lazy_feasibility_cuts": lcb.feascuts_added,
        "selected_products": str(selected_final_labels),
        "start_datetime": start_datetime.strftime("%Y-%m-%d %H:%M:%S"),
        "end_datetime": end_datetime.strftime("%Y-%m-%d %H:%M:%S"),
        "detail_output_file": output_file,
    }


# ============================================================
# BATCH RUNNER
# ============================================================

if __name__ == "__main__":
    batch_start_time = time.perf_counter()

    experiments = list(itertools.product(N_values, P_values, K_values, V_values, S_values))
    all_summary_rows = []

    for sample_id, (n_value, P_value, K_value, V_value, S_value) in enumerate(experiments, start=1):
        try:
            row = solve_BD2_manual_closed_form(
                address=address,
                n_value=n_value,
                P_value=P_value,
                K_value=K_value,
                V_value=V_value,
                S_value=S_value,
                sample_id=sample_id,
            )
        except Exception as e:
            row = {
                "sample_id": sample_id,
                "n": n_value,
                "P": P_value,
                "K": K_value,
                "V": V_value,
                "S": S_value,
                "status": "ERROR",
                "error_message": repr(e),
            }
            print(f"ERROR in sample {sample_id}: {repr(e)}")

        all_summary_rows.append(row)

    batch_end_time = time.perf_counter()

    summary_output = os.path.join(os.path.dirname(address), "BD2_closed_form_summary_all_runs.xlsx")
    df_summary_all = pd.DataFrame(all_summary_rows)

    with pd.ExcelWriter(summary_output, engine="xlsxwriter") as writer:
        df_summary_all.to_excel(writer, sheet_name="summary_all_runs", index=False)
        format_excel_worksheets(writer)

    print("\n====================================================")
    print("BD2 BATCH RUN FINISHED")
    print(f"Number of runs attempted: {len(experiments)}")
    print(f"Batch total wall-clock time: {batch_end_time - batch_start_time:.2f} seconds")
    print(f"Each CPLEX run had its own time limit: {TIME_LIMIT} seconds")
    print(f"Master summary exported to: {summary_output}")
    print("====================================================")
