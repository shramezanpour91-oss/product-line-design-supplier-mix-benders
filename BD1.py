# ============================================================
# Batch-from-one-Excel Benders Decomposition using CPLEX Lazy Cuts
# with automatic dataset scaling and Excel exports
# ============================================================

import os
import sys
import time
import math
import itertools
from datetime import datetime
from itertools import combinations
import pandas as pd
import cplex
from cplex.callbacks import LazyConstraintCallback
from cplex.exceptions import CplexError


# ============================================================
# USER SETTINGS
# ============================================================

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
address = BASE_DIR / "data" / "Main Dataset.xlsx"

N_values = [10]
K_values = [200]
V_values = [3]
P_values = [1000]
S_values = [2]

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

# General numerical tolerance
TOL = 1e-6

# Cut-screening tolerances
CUT_ABS_TOL = 1e-4
CUT_REL_TOL = 1e-6


# ============================================================
# EXCEL FORMATTER
# ============================================================

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


# ============================================================
# MAIN SOLVER
# ============================================================

def solve_one_sample(address, n_value, P_value, K_value, V_value, S_value, sample_id):

    total_start_time = time.perf_counter()
    start_datetime = datetime.now()

    print("\n====================================================")
    print(f"Solving Benders sample {sample_id}")
    print(f"n={n_value}, P={P_value}, K={K_value}, V={V_value}, S={S_value}")
    print(f"Input file: {address}")
    print(f"Time limit for this run = {TIME_LIMIT} seconds")
    print("====================================================\n")

    # ========================================================
    # READ DATA FROM MAIN EXCEL FILE
    # ========================================================

    lam_df = pd.DataFrame(pd.read_excel(address, sheet_name="Lambda^k"))
    Pi_df = pd.DataFrame(pd.read_excel(address, sheet_name="Pi_i"))
    a_df = pd.DataFrame(pd.read_excel(address, sheet_name="a_ip"))
    sigma_df = pd.DataFrame(pd.read_excel(address, sheet_name="Sigma_i^k"))
    r_df = pd.DataFrame(pd.read_excel(address, sheet_name="r_p^s"))
    Cap_df = pd.DataFrame(pd.read_excel(address, sheet_name="Cap_p^s"))

    # ========================================================
    # SETS BASED ON CURRENT SAMPLE
    # ========================================================

    I = list(range(0, n_value + 1))
    Ip = list(range(1, n_value + 1))
    K = list(range(1, K_value + 1))
    P = list(range(1, P_value + 1))

    S_original = list(range(1, S_value + 1))
    artificial_supplier = S_value + 1
    S = S_original + [artificial_supplier]

    V = V_value

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


    # ========================================================
    # PARAMETERS
    # ========================================================

    lam = {k: float(lam_df.loc[0, k]) for k in K}

    Pi = {
        i: float(Pi_df.loc[0, i]) * Pi_multiplier
        for i in I
    }

    a = {
        (i, p): float(a_df.loc[i, p])
        for i in Ip
        for p in P
    }

    sigma = {
        (i, k): float(sigma_df.loc[i, k])
        for i in I
        for k in K
    }

    r = {}
    Cap = {}

    for p in P:
        for s in S_original:
            r[(p, s)] = float(r_df.loc[p - 1, s])
            Cap[(p, s)] = float(Cap_df.loc[p - 1, s]) * Cap_multiplier

        # Artificial supplier used internally only
        r[(p, artificial_supplier)] = ARTIFICIAL_SUPPLIER_COST
        Cap[(p, artificial_supplier)] = ARTIFICIAL_SUPPLIER_CAP

    # Pool for avoiding duplicate or nearly duplicate cuts
    CUT_POOL = set()

    # ========================================================
    # VARIABLE NAME PRECOMPUTE
    # ========================================================

    def x_name(i):
        return f"x_{i}"

    def y_name(i, k):
        return f"y_{i}_{k}"

    Z_NAME = "z"

    X_names = {i: x_name(i) for i in Ip}
    Y_names = {(i, k): y_name(i, k) for i in I for k in K}

    Y_ip_pairs = [(i, k) for i in Ip for k in K]
    Y_ip_names = [Y_names[(i, k)] for (i, k) in Y_ip_pairs]
    Y_ip_index = {(i, k): idx for idx, (i, k) in enumerate(Y_ip_pairs)}

    Z_IDX = None
    Y_ip_var_indices = None

    # ========================================================
    # PREPROCESS SUPPLIERS SORTED BY UNIT COST
    # ========================================================

    supplier_order = {}
    total_cap = {}
    prefix_caps = {}
    prefix_rcap = {}
    sorted_supplier_id = {}
    sorted_r = {}
    sorted_cap = {}

    for p in P:
        ordered = sorted(
            [(s, r[(p, s)], Cap[(p, s)]) for s in S],
            key=lambda x: (x[1], x[0])
        )

        supplier_order[p] = ordered
        prefix_caps[p] = [0.0]
        prefix_rcap[p] = [0.0]
        sorted_supplier_id[p] = [None]
        sorted_r[p] = [None]
        sorted_cap[p] = [None]

        for s_id, cost_ps, cap_ps in ordered:
            sorted_supplier_id[p].append(s_id)
            sorted_r[p].append(cost_ps)
            sorted_cap[p].append(cap_ps)
            prefix_caps[p].append(prefix_caps[p][-1] + cap_ps)
            prefix_rcap[p].append(prefix_rcap[p][-1] + cost_ps * cap_ps)

        total_cap[p] = prefix_caps[p][-1]

    # ========================================================
    # PRECOMPUTE A_p(y) TERMS FOR CUT GENERATION
    # ========================================================

    Ap_terms = {}

    for p in P:
        terms = []

        for i in Ip:
            a_ip = a[(i, p)]

            if abs(a_ip) <= 1e-12:
                continue

            for k in K:
                coeff = lam[k] * a_ip

                if abs(coeff) > 1e-12:
                    var_nm = Y_names[(i, k)]
                    y_pos = Y_ip_index[(i, k)]
                    terms.append((var_nm, coeff, i, k, y_pos))

        Ap_terms[p] = terms

    Ap_ypos = {}
    Ap_coeff = {}

    for p in P:
        Ap_ypos[p] = [term[4] for term in Ap_terms[p]]
        Ap_coeff[p] = [float(term[1]) for term in Ap_terms[p]]

    # ========================================================
    # CUSTOMER RANK PRECOMPUTE
    # ========================================================

    rank_pos = {}
    sorted_products_for_k = {}

    for k in K:
        order_k = sorted(I, key=lambda i: (sigma[(i, k)], i))
        sorted_products_for_k[k] = order_k
        rank_pos[k] = {i: pos for pos, i in enumerate(order_k)}

    # ========================================================
    # HELPER STRUCTURES
    # ========================================================

    better_set = {}

    for i in Ip:
        for k in K:
            better_set[(i, k)] = [
                j for j in Ip
                if sigma[(j, k)] > sigma[(i, k)]
            ]

    # ========================================================
    # HELPER FUNCTIONS
    # ========================================================

    def compute_A_p(y_values, p):
        return sum(
            coeff * y_values[(i, k)]
            for _, coeff, i, k, _ in Ap_terms[p]
        )

    def pivotal_supplier_position(p, Abar):
        if Abar <= TOL:
            return 1

        ordered_len = len(supplier_order[p])

        for t in range(1, ordered_len + 1):
            if prefix_caps[p][t] >= Abar - TOL:
                return t

        raise RuntimeError(
            f"No supplier capacity covers demand Abar={Abar} for component p={p}. "
            f"Please increase ARTIFICIAL_SUPPLIER_CAP."
        )

    def pivotal_supplier_data(p, Abar):
        t_hat = pivotal_supplier_position(p, Abar)
        s_hat = sorted_supplier_id[p][t_hat]
        r_hat = sorted_r[p][t_hat]

        return t_hat, s_hat, r_hat

    def recourse_value_G_p_formula8(p, Abar):
        t_hat, _, r_hat = pivotal_supplier_data(p, Abar)
        first_term = r_hat * prefix_caps[p][t_hat] - prefix_rcap[p][t_hat]

        return first_term - Abar * r_hat

    def recourse_value_G_p(p, Abar):
        return recourse_value_G_p_formula8(p, Abar)

    def compute_total_recourse(y_values):
        total_g = 0.0

        for p in P:
            Abar_p = compute_A_p(y_values, p)
            total_g += recourse_value_G_p(p, Abar_p)

        return total_g

    def compute_revenue(y_values):
        return sum(
            lam[k] * Pi[i] * y_values[(i, k)]
            for i in Ip
            for k in K
        )

    def build_optimality_cut_from_formula8(y_val):
        rhs_cut = 0.0
        coeff_y = {(i, k): 0.0 for i in Ip for k in K}
        G_bar = 0.0

        for p in P:
            Abar_p = compute_A_p(y_val, p)
            t_hat, _, alpha_star = pivotal_supplier_data(p, Abar_p)

            rhs_p = alpha_star * prefix_caps[p][t_hat] - prefix_rcap[p][t_hat]
            rhs_cut += rhs_p

            for _, coeff, i, k, _ in Ap_terms[p]:
                coeff_y[(i, k)] += alpha_star * coeff

            G_bar += rhs_p - alpha_star * Abar_p

        ind = [Z_NAME]
        val = [1.0]

        for i in Ip:
            for k in K:
                c = coeff_y[(i, k)]

                if abs(c) > 1e-12:
                    ind.append(Y_names[(i, k)])
                    val.append(float(c))

        return ind, val, float(rhs_cut), float(G_bar)

    def build_optimality_cut_from_formula8_array(y_ip_vals):
        rhs_cut = 0.0
        coeff_y = [0.0] * len(Y_ip_pairs)
        G_bar = 0.0

        for p in P:
            coeffs = Ap_coeff[p]
            positions = Ap_ypos[p]

            Abar_p = 0.0

            for c, pos in zip(coeffs, positions):
                Abar_p += c * y_ip_vals[pos]

            t_hat = pivotal_supplier_position(p, Abar_p)
            alpha_star = sorted_r[p][t_hat]

            rhs_p = alpha_star * prefix_caps[p][t_hat] - prefix_rcap[p][t_hat]
            rhs_cut += rhs_p

            for c, pos in zip(coeffs, positions):
                coeff_y[pos] += alpha_star * c

            G_bar += rhs_p - alpha_star * Abar_p

        ind = [Z_IDX]
        val = [1.0]

        for pos, c in enumerate(coeff_y):
            if abs(c) > 1e-12:
                ind.append(Y_ip_var_indices[pos])
                val.append(float(c))

        return ind, val, float(rhs_cut), float(G_bar)

    def make_cut_signature(ind, val, rhs_cut, decimals=8):
        return (
            tuple(ind),
            tuple(round(v, decimals) for v in val),
            round(rhs_cut, decimals)
        )

    # ========================================================
    # GREEDY HEURISTIC FOR LOWER BOUND
    # ========================================================

    def assign_choice_given_selected(selected_products):
        selected_set = set(selected_products)
        choice = {}

        for k in K:
            for i in sorted_products_for_k[k]:
                if i == 0 or i in selected_set:
                    choice[k] = i
                    break

        return choice

    def build_y_from_choice(choice):
        y_vals = {(i, k): 0.0 for i in I for k in K}

        for k in K:
            y_vals[(choice[k], k)] = 1.0

        return y_vals

    def build_compact_state_from_choice(selected_products, choice):
        revenue_val = 0.0
        A_vals = {p: 0.0 for p in P}

        for k in K:
            chosen_i = choice[k]

            if chosen_i == 0:
                continue

            revenue_val += lam[k] * Pi[chosen_i]

            for p in P:
                A_vals[p] += lam[k] * a[(chosen_i, p)]

        G_val = 0.0

        for p in P:
            G_val += recourse_value_G_p(p, A_vals[p])

        return {
            "feasible": True,
            "selected": list(selected_products),
            "choice": choice,
            "revenue": revenue_val,
            "A": A_vals,
            "G": G_val,
            "obj": revenue_val + G_val
        }

    def evaluate_selected_set(selected_products):
        choice = assign_choice_given_selected(selected_products)
        return build_compact_state_from_choice(selected_products, choice)

    def evaluate_add_candidate_incremental(current_eval, candidate_i):
        new_choice = dict(current_eval["choice"])
        new_A = dict(current_eval["A"])
        new_revenue = float(current_eval["revenue"])

        switched_any = False

        for k in K:
            old_i = current_eval["choice"][k]

            if rank_pos[k][candidate_i] < rank_pos[k][old_i]:
                switched_any = True
                new_choice[k] = candidate_i

                old_rev = 0.0 if old_i == 0 else lam[k] * Pi[old_i]
                new_rev = lam[k] * Pi[candidate_i]
                new_revenue += new_rev - old_rev

                for p in P:
                    old_a = 0.0 if old_i == 0 else lam[k] * a[(old_i, p)]
                    new_a = lam[k] * a[(candidate_i, p)]
                    new_A[p] += new_a - old_a

        if not switched_any:
            return {
                "feasible": True,
                "selected": current_eval["selected"] + [candidate_i],
                "choice": new_choice,
                "revenue": new_revenue,
                "A": new_A,
                "G": current_eval["G"],
                "obj": new_revenue + current_eval["G"]
            }

        new_G = 0.0

        for p in P:
            new_G += recourse_value_G_p(p, new_A[p])

        return {
            "feasible": True,
            "selected": current_eval["selected"] + [candidate_i],
            "choice": new_choice,
            "revenue": new_revenue,
            "A": new_A,
            "G": new_G,
            "obj": new_revenue + new_G
        }

    def evaluate_add_candidates_incremental(current_eval, candidate_list):
        trial_eval = current_eval

        for candidate_i in candidate_list:
            trial_eval = evaluate_add_candidate_incremental(trial_eval, candidate_i)

        return trial_eval

    def greedy_initial_solution():
        greedy_start_time = time.perf_counter()

        current_selected = []
        current_eval = evaluate_selected_set(current_selected)
        current_obj = current_eval["obj"]

        greedy_steps = []

        print("\n================ PAIR-AWARE GREEDY FORWARD ADDITION ================")
        print(f"Initial selected products = {current_selected}")
        print(f"Initial revenue           = {current_eval['revenue']:.6f}")
        print(f"Initial recourse G(y)     = {current_eval['G']:.6f}")
        print(f"Initial objective         = {current_obj:.6f}")
        print("--------------------------------------------------------------------")

        step = 0

        while len(current_selected) < V:
            step += 1

            best_single_eval = None
            best_single_product = None

            remaining_products = [
                i for i in Ip
                if i not in current_selected
            ]

            # Single-product addition
            for i in remaining_products:
                trial_eval = evaluate_add_candidate_incremental(current_eval, i)

                if trial_eval["obj"] > current_eval["obj"] + TOL:
                    if (
                        best_single_eval is None
                        or trial_eval["obj"] > best_single_eval["obj"] + TOL
                    ):
                        best_single_eval = trial_eval
                        best_single_product = i

            if best_single_eval is not None:
                old_obj = current_eval["obj"]

                current_selected = best_single_eval["selected"]
                current_eval = best_single_eval

                improvement = current_eval["obj"] - old_obj

                greedy_steps.append({
                    "step": step,
                    "addition_type": "single",
                    "added_products": [best_single_product],
                    "objective_after_addition": current_eval["obj"],
                    "improvement": improvement,
                    "revenue_after_addition": current_eval["revenue"],
                    "recourse_after_addition": current_eval["G"],
                    "selected_products": current_selected.copy(),
                })

                print(
                    f"Step {step}: added product {best_single_product} | "
                    f"objective = {current_eval['obj']:.6f} | "
                    f"improvement = {improvement:.6f} | "
                    f"revenue = {current_eval['revenue']:.6f} | "
                    f"G(y) = {current_eval['G']:.6f} | "
                    f"selected = {current_selected}"
                )

                continue

            # Pair-aware fallback
            best_pair_eval = None
            best_pair_products = None

            if len(current_selected) + 2 <= V:
                for i, j in combinations(remaining_products, 2):
                    trial_eval = evaluate_add_candidates_incremental(
                        current_eval,
                        [i, j]
                    )

                    if trial_eval["obj"] > current_eval["obj"] + TOL:
                        if (
                            best_pair_eval is None
                            or trial_eval["obj"] > best_pair_eval["obj"] + TOL
                        ):
                            best_pair_eval = trial_eval
                            best_pair_products = [i, j]

            if best_pair_eval is not None:
                old_obj = current_eval["obj"]

                current_selected = best_pair_eval["selected"]
                current_eval = best_pair_eval

                improvement = current_eval["obj"] - old_obj

                greedy_steps.append({
                    "step": step,
                    "addition_type": "pair",
                    "added_products": best_pair_products.copy(),
                    "objective_after_addition": current_eval["obj"],
                    "improvement": improvement,
                    "revenue_after_addition": current_eval["revenue"],
                    "recourse_after_addition": current_eval["G"],
                    "selected_products": current_selected.copy(),
                })

                print(
                    f"Step {step}: added pair {best_pair_products} | "
                    f"objective = {current_eval['obj']:.6f} | "
                    f"improvement = {improvement:.6f} | "
                    f"revenue = {current_eval['revenue']:.6f} | "
                    f"G(y) = {current_eval['G']:.6f} | "
                    f"selected = {current_selected}"
                )

                continue

            print(
                f"Step {step}: no improving single product or pair found. "
                f"Greedy stops."
            )
            break

        greedy_time = time.perf_counter() - greedy_start_time

        current_eval["greedy_steps"] = greedy_steps
        current_eval["greedy_time_seconds"] = greedy_time

        print("--------------------------------------------------------------------")
        print(f"Greedy selected products: {current_eval['selected']}")
        print(f"Greedy revenue:           {current_eval['revenue']:.6f}")
        print(f"Greedy recourse G(y):     {current_eval['G']:.6f}")
        print(f"Greedy lower bound:       {current_eval['obj']:.6f}")
        print(f"Greedy time:              {greedy_time:.4f} seconds")
        print("====================================================================\n")

        return current_eval

    # ========================================================
    # BENDERS LAZY CALLBACK
    # ========================================================

    class BendersLazyCallback(LazyConstraintCallback):
        def __call__(self):
            z_val = self.get_values(Z_IDX)
            y_ip_vals = self.get_values(Y_ip_var_indices)

            ind, val, rhs_cut, G_bar = build_optimality_cut_from_formula8_array(
                y_ip_vals
            )

            violation = z_val - G_bar
            scale = max(1.0, abs(z_val), abs(G_bar))
            violation_threshold = max(CUT_ABS_TOL, CUT_REL_TOL * scale)

            if violation <= violation_threshold:
                return

            signature = make_cut_signature(ind, val, rhs_cut, decimals=8)

            if signature in CUT_POOL:
                return

            CUT_POOL.add(signature)

            self.add(
                constraint=cplex.SparsePair(ind=ind, val=val),
                sense="L",
                rhs=rhs_cut
            )

    # ========================================================
    # BUILD MASTER MODEL
    # ========================================================

    build_start_time = time.perf_counter()

    master = cplex.Cplex()
    master.objective.set_sense(master.objective.sense.maximize)

    master.set_log_stream(sys.stdout)
    master.set_error_stream(sys.stderr)
    master.set_warning_stream(sys.stderr)
    master.set_results_stream(sys.stdout)

    master.parameters.timelimit.set(TIME_LIMIT)
    master.parameters.threads.set(THREADS)

    master.parameters.mip.tolerances.mipgap.set(1e-4)
    master.parameters.mip.tolerances.absmipgap.set(1e-6)

    master.parameters.preprocessing.presolve.set(1)
    master.parameters.emphasis.mip.set(2)

    # ========================================================
    # VARIABLES
    # ========================================================

    var_names = []
    obj = []
    lb = []
    ub = []
    types = []

    # x variables
    for i in Ip:
        var_names.append(X_names[i])
        obj.append(0.0)
        lb.append(0.0)
        ub.append(1.0)
        types.append(master.variables.type.binary)

    # y variables
    for i in I:
        for k in K:
            var_names.append(Y_names[(i, k)])

            coef = 0.0 if i == 0 else lam[k] * Pi[i]
            obj.append(float(coef))

            lb.append(0.0)
            ub.append(1.0)
            types.append(master.variables.type.continuous)

    # z variable
    var_names.append(Z_NAME)
    obj.append(1.0)
    lb.append(-cplex.infinity)
    ub.append(0.0)
    types.append(master.variables.type.continuous)

    master.variables.add(
        obj=obj,
        lb=lb,
        ub=ub,
        types="".join(types),
        names=var_names
    )

    Z_IDX = master.variables.get_indices(Z_NAME)
    Y_ip_var_indices = master.variables.get_indices(Y_ip_names)

    # ========================================================
    # CONSTRAINTS
    # ========================================================

    rows = []
    senses = []
    rhs = []
    constraint_names = []

    # Assignment constraints: sum_i y[i,k] = 1 for all k
    for k in K:
        ind = [Y_names[(i, k)] for i in I]
        val = [1.0] * len(ind)

        rows.append(cplex.SparsePair(ind=ind, val=val))
        senses.append("E")
        rhs.append(1.0)
        constraint_names.append(f"assignment_{k}")

    # Cardinality constraint: sum_i x[i] <= V
    rows.append(
        cplex.SparsePair(
            ind=[X_names[i] for i in Ip],
            val=[1.0] * len(Ip)
        )
    )
    senses.append("L")
    rhs.append(float(V))
    constraint_names.append("cardinality")

    # Linking constraints: y[i,k] <= x[i]
    for i in Ip:
        for k in K:
            rows.append(
                cplex.SparsePair(
                    ind=[Y_names[(i, k)], X_names[i]],
                    val=[1.0, -1.0]
                )
            )
            senses.append("L")
            rhs.append(0.0)
            constraint_names.append(f"link_{i}_{k}")

    # Preference constraints:
    # x[i] + sum_{j: sigma[j,k] > sigma[i,k]} y[j,k] <= 1
    for i in Ip:
        for k in K:
            ind = [X_names[i]]
            val = [1.0]

            for j in better_set[(i, k)]:
                ind.append(Y_names[(j, k)])
                val.append(1.0)

            rows.append(cplex.SparsePair(ind=ind, val=val))
            senses.append("L")
            rhs.append(1.0)
            constraint_names.append(f"preference_{i}_{k}")

    master.linear_constraints.add(
        lin_expr=rows,
        senses=senses,
        rhs=rhs,
        names=constraint_names
    )

    # ========================================================
    # GREEDY INITIAL LOWER BOUND
    # ========================================================

    greedy_sol = greedy_initial_solution()
    greedy_time = greedy_sol["greedy_time_seconds"]

    greedy_y = build_y_from_choice(greedy_sol["choice"])
    INITIAL_LB = greedy_sol["obj"]

    print("\n================ PAIR-AWARE GREEDY INITIAL SOLUTION SUMMARY =================")
    print("Selected products (greedy):", greedy_sol["selected"])
    print("Greedy revenue            =", greedy_sol["revenue"])
    print("Greedy recourse G(y)      =", greedy_sol["G"])
    print("Initial LB (greedy)       =", INITIAL_LB)
    print("Greedy time (s)           =", greedy_time)

    print("\nGreedy addition steps:")
    for row in greedy_sol["greedy_steps"]:
        print(
            f"Step {row['step']} | "
            f"type = {row['addition_type']} | "
            f"added = {row['added_products']} | "
            f"objective = {row['objective_after_addition']:.6f} | "
            f"improvement = {row['improvement']:.6f} | "
            f"revenue = {row['revenue_after_addition']:.6f} | "
            f"G(y) = {row['recourse_after_addition']:.6f} | "
            f"selected = {row['selected_products']}"
        )

    print("============================================================================\n")

    # Add initial Benders optimality cut from greedy solution
    ind, val, rhs_cut, G_bar = build_optimality_cut_from_formula8(greedy_y)

    master.linear_constraints.add(
        lin_expr=[cplex.SparsePair(ind=ind, val=val)],
        senses=["L"],
        rhs=[rhs_cut],
        names=["initial_benders_cut"]
    )

    CUT_POOL.add(make_cut_signature(ind, val, rhs_cut, decimals=8))

    # Build MIP start from greedy solution
    start_names = []
    start_values = []

    selected_set = set(greedy_sol["selected"])

    for i in Ip:
        start_names.append(X_names[i])
        start_values.append(1.0 if i in selected_set else 0.0)

    for i in I:
        for k in K:
            start_names.append(Y_names[(i, k)])
            start_values.append(greedy_y[(i, k)])

    start_names.append(Z_NAME)
    start_values.append(greedy_sol["G"])

    try:
        master.MIP_starts.add(
            cplex.SparsePair(ind=start_names, val=start_values),
            master.MIP_starts.effort_level.auto,
            "greedy_start"
        )

        print("================ MIP START INFORMATION =================")
        print("Greedy MIP start added successfully.")
        print("Initial LB from greedy =", INITIAL_LB)
        print("Initial z value        =", greedy_sol["G"])
        print("=========================================================\n")

    except CplexError as e:
        print("\nWARNING: Could not add greedy MIP start.")
        print("CPLEX error:", e)

    master.register_callback(BendersLazyCallback)

    build_time = time.perf_counter() - build_start_time

    # ========================================================
    # SOLVE
    # ========================================================

    solve_start_time = time.perf_counter()

    try:
        master.solve()
    except CplexError as e:
        raise RuntimeError(f"CPLEX failed while solving the Benders master: {e}")

    solve_time = time.perf_counter() - solve_start_time
    total_time = time.perf_counter() - total_start_time
    end_datetime = datetime.now()

    # ========================================================
    # SOLUTION OUTPUT
    # ========================================================

    status_code = master.solution.get_status()
    status_string = master.solution.get_status_string(status_code)
    time_limit_reached = "time limit" in status_string.lower()

    if not master.solution.is_primal_feasible():
        return {
            "sample_id": sample_id,
            "n": n_value,
            "P": P_value,
            "K": K_value,
            "V": V_value,
            "S": S_value,
            "objective_value": None,
            "best_bound": None,
            "mip_gap_percent": None,
            "z_value": None,
            "initial_LB_greedy": INITIAL_LB,
            "greedy_revenue": greedy_sol["revenue"],
            "greedy_recourse_G": greedy_sol["G"],
            "number_of_benders_cuts": len(CUT_POOL),
            "status_code": status_code,
            "status": status_string,
            "time_limit_reached": time_limit_reached,
            "build_time_seconds": build_time,
            "solver_time_seconds": solve_time,
            "total_time_seconds": total_time,
            "greedy_time_seconds": greedy_time,
            "selected_products": "",
            "greedy_selected_products": ", ".join(map(str, greedy_sol["selected"])),
        }

    obj_value = master.solution.get_objective_value()

    try:
        best_bound = master.solution.MIP.get_best_objective()
    except CplexError:
        best_bound = None

    try:
        mip_gap = master.solution.MIP.get_mip_relative_gap()
        mip_gap_percent = 100.0 * mip_gap
    except CplexError:
        mip_gap = None
        mip_gap_percent = None

    x_values = {
        i: master.solution.get_values(X_names[i])
        for i in Ip
    }

    offered_products = [
        i for i in Ip
        if x_values[i] > 0.5
    ]

    z_value = master.solution.get_values(Z_NAME)

    print("\n================ FINAL SOLUTION =================")
    print("Final objective value     =", obj_value)
    print("Best bound                =", best_bound)
    print("Final MIP gap percent     =", mip_gap_percent)
    print("Solution status code      =", status_code)
    print("Solution status           =", status_string)
    print("z value                   =", z_value)
    print("Initial LB (greedy)       =", INITIAL_LB)
    print("Greedy selected products  =", greedy_sol["selected"])
    print("Final offered products    =", offered_products)
    print("Number of Benders cuts    =", len(CUT_POOL))
    print("Build time (s)            =", build_time)
    print("Solver time (s)           =", solve_time)
    print("Total solution time (s)   =", total_time)
    print("Start time                =", start_datetime)
    print("End time                  =", end_datetime)

    # ========================================================
    # EXTRACT SOLUTION
    # ========================================================

    x_rows = []
    for i in Ip:
        x_rows.append({
            "i": i,
            "x_value": x_values[i],
            "greedy_x_value": 1.0 if i in selected_set else 0.0
        })

    df_x = pd.DataFrame(x_rows)

    y_rows = []
    for i in I:
        for k in K:
            y_val = master.solution.get_values(Y_names[(i, k)])
            y_rows.append({
                "i": i,
                "k": k,
                "y_value": y_val,
                "greedy_y_value": greedy_y[(i, k)]
            })

    df_y = pd.DataFrame(y_rows)

    df_summary = pd.DataFrame([
        {"metric": "sample_id", "value": sample_id},
        {"metric": "n", "value": n_value},
        {"metric": "P", "value": P_value},
        {"metric": "K", "value": K_value},
        {"metric": "V", "value": V_value},
        {"metric": "S", "value": S_value},
        {"metric": "final_objective_value", "value": obj_value},
        {"metric": "best_bound", "value": best_bound},
        {"metric": "mip_gap_percent", "value": mip_gap_percent},
        {"metric": "z_value", "value": z_value},
        {"metric": "initial_LB_greedy", "value": INITIAL_LB},
        {"metric": "greedy_revenue", "value": greedy_sol["revenue"]},
        {"metric": "greedy_recourse_G", "value": greedy_sol["G"]},
        {"metric": "greedy_time_seconds", "value": greedy_time},
        {"metric": "number_of_benders_cuts", "value": len(CUT_POOL)},
        {"metric": "status_code", "value": status_code},
        {"metric": "status_string", "value": status_string},
        {"metric": "time_limit_reached", "value": time_limit_reached},
        {"metric": "build_time_seconds", "value": build_time},
        {"metric": "solver_time_seconds", "value": solve_time},
        {"metric": "total_time_seconds", "value": total_time},
        {"metric": "time_limit_seconds_per_run", "value": TIME_LIMIT},
        {"metric": "selected_products", "value": ", ".join(map(str, offered_products))},
        {"metric": "greedy_selected_products", "value": ", ".join(map(str, greedy_sol["selected"]))},
    ])

    df_final_offered_products = pd.DataFrame({
        "offered_product": offered_products
    })

    df_greedy_offered_products = pd.DataFrame({
        "greedy_offered_product": greedy_sol["selected"]
    })

    df_greedy_steps = pd.DataFrame(greedy_sol["greedy_steps"])

    df_greedy_choice = pd.DataFrame([
        {
            "k": k,
            "greedy_chosen_product": greedy_sol["choice"][k]
        }
        for k in K
    ])

    df_greedy_A = pd.DataFrame([
        {
            "p": p,
            "greedy_A_p": greedy_sol["A"][p]
        }
        for p in P
    ])

    df_x_nonzero = df_x[df_x["x_value"].abs() > TOL].copy()
    df_y_nonzero = df_y[df_y["y_value"].abs() > TOL].copy()

    df_greedy_x_nonzero = df_x[df_x["greedy_x_value"].abs() > TOL].copy()
    df_greedy_y_nonzero = df_y[df_y["greedy_y_value"].abs() > TOL].copy()

    # ========================================================
    # EXPORT TO EXCEL
    # ========================================================

    output_dir = os.path.dirname(address)
    output_subdir = os.path.join(output_dir, "Benders_subset_results")
    os.makedirs(output_subdir, exist_ok=True)

    detailed_output_file = os.path.join(
        output_subdir,
        f"Benders_solution_sample_{sample_id}_n={n_value}_P={P_value}_K={K_value}_V={V_value}_S={S_value}.xlsx"
    )

    with pd.ExcelWriter(detailed_output_file, engine="xlsxwriter") as writer:
        df_summary.to_excel(writer, sheet_name="summary", index=False)

        df_final_offered_products.to_excel(writer, sheet_name="final_offered_products", index=False)
        df_greedy_offered_products.to_excel(writer, sheet_name="greedy_offered_products", index=False)

        df_greedy_steps.to_excel(writer, sheet_name="greedy_steps", index=False)
        df_greedy_choice.to_excel(writer, sheet_name="greedy_choice_by_k", index=False)
        df_greedy_A.to_excel(writer, sheet_name="greedy_A_by_p", index=False)

        df_x.to_excel(writer, sheet_name="x_all", index=False)
        df_y.to_excel(writer, sheet_name="y_all", index=False)

        df_x_nonzero.to_excel(writer, sheet_name="x_nonzero", index=False)
        df_y_nonzero.to_excel(writer, sheet_name="y_nonzero", index=False)

        df_greedy_x_nonzero.to_excel(writer, sheet_name="greedy_x_nonzero", index=False)
        df_greedy_y_nonzero.to_excel(writer, sheet_name="greedy_y_nonzero", index=False)

        format_excel_worksheets(writer)

    print(f"\nDetailed solution exported to: {detailed_output_file}")

    return {
        "sample_id": sample_id,
        "n": n_value,
        "P": P_value,
        "K": K_value,
        "V": V_value,
        "S": S_value,
        "objective_value": obj_value,
        "best_bound": best_bound,
        "mip_gap_percent": mip_gap_percent,
        "z_value": z_value,
        "initial_LB_greedy": INITIAL_LB,
        "greedy_revenue": greedy_sol["revenue"],
        "greedy_recourse_G": greedy_sol["G"],
        "number_of_benders_cuts": len(CUT_POOL),
        "status_code": status_code,
        "status": status_string,
        "time_limit_reached": time_limit_reached,
        "build_time_seconds": build_time,
        "solver_time_seconds": solve_time,
        "total_time_seconds": total_time,
        "greedy_time_seconds": greedy_time,
        "selected_products": ", ".join(map(str, offered_products)),
        "greedy_selected_products": ", ".join(map(str, greedy_sol["selected"])),
    }


# ============================================================
# BATCH LOOP
# ============================================================

print("Python used:")
print(sys.executable)
print("CPLEX used:")
print(cplex.__file__)
print("CPLEX version:")
print(cplex.Cplex().get_version())

all_results = []

experiments = list(itertools.product(
    N_values,
    P_values,
    K_values,
    V_values,
    S_values
))

batch_start_time = time.perf_counter()

for sample_id, (n_value, P_value,  K_value, V_value, S_value) in enumerate(experiments, start=1):

    try:
        result_row = solve_one_sample(
            address=address,
            n_value=n_value,
            P_value=P_value,
            K_value=K_value,
            V_value=V_value,
            S_value=S_value,
            sample_id=sample_id
        )

        all_results.append(result_row)

    except Exception as e:
        print(f"\nERROR in sample {sample_id}: {e}")

        Pi_multiplier = P_value / 1000
        supplier_multiplier = 2 / S_value
        K_multiplier = K_value / 200
        Cap_multiplier = supplier_multiplier * K_multiplier

        all_results.append({
            "sample_id": sample_id,
            "n": n_value,
            "P": P_value,
            "K": K_value,
            "V": V_value,
            "S": S_value,
            "objective_value": None,
            "best_bound": None,
            "mip_gap_percent": None,
            "z_value": None,
            "initial_LB_greedy": None,
            "greedy_revenue": None,
            "greedy_recourse_G": None,
            "number_of_benders_cuts": None,
            "status_code": None,
            "status": f"FAILED: {e}",
            "time_limit_reached": None,
            "build_time_seconds": None,
            "solver_time_seconds": None,
            "total_time_seconds": None,
            "greedy_time_seconds": None,
            "selected_products": "",
            "greedy_selected_products": "",
        })

batch_end_time = time.perf_counter()


# ============================================================
# MASTER SUMMARY EXPORT
# ============================================================

summary_df = pd.DataFrame(all_results)

output_dir = os.path.dirname(address)
output_subdir = os.path.join(output_dir, "Benders_subset_results")
os.makedirs(output_subdir, exist_ok=True)

summary_output = os.path.join(output_subdir, "Benders_all_subset_runs_summary.xlsx")

with pd.ExcelWriter(summary_output, engine="xlsxwriter") as writer:
    summary_df.to_excel(writer, sheet_name="summary_all_runs", index=False)
    format_excel_worksheets(writer)

print("\n====================================================")
print("BENDERS BATCH RUN FINISHED")
print(f"Number of runs attempted: {len(experiments)}")
print(f"Batch total wall-clock time: {batch_end_time - batch_start_time:.2f} seconds")
print(f"Each CPLEX run had its own time limit: {TIME_LIMIT} seconds")
print(f"Master summary exported to: {summary_output}")
print("====================================================")
