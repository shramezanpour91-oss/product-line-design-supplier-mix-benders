# ============================================================
# Batch-from-one-Excel Direct MILP using CPLEX API
# ============================================================

import os
import sys
import time
import math
import itertools
import pandas as pd
import cplex
from cplex.exceptions import CplexError
from cplex.callbacks import MIPInfoCallback


# ============================================================
# USER SETTINGS
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
ARTIFICIAL_SUPPLIER_CAP =50000.0

TIME_LIMIT = 3600
PROGRESS_INTERVAL_SECONDS = 60
eps = 1e-9


# ============================================================
# EXCEL FORMATTER: all numbers shown as 0.1 decimal
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
            worksheet.set_column(20, 30, 45, text_format)


# ============================================================
# MAIN SOLVER
# ============================================================

def solve_one_sample(address, n_value, P_value, K_value, V_value, S_value, sample_id):

    total_start_time = time.perf_counter()

    print("\n====================================================")
    print(f"Solving sample {sample_id}")
    print(f"n={n_value}, P={P_value}, K={K_value}, V={V_value}, S={S_value}")
    print(f"Input file: {address}")
    print(f"Time limit for this run = {TIME_LIMIT} seconds")
    print("====================================================\n")

    lam_df = pd.DataFrame(pd.read_excel(address, sheet_name="Lambda^k"))
    Pi_df = pd.DataFrame(pd.read_excel(address, sheet_name="Pi_i"))
    a_df = pd.DataFrame(pd.read_excel(address, sheet_name="a_ip"))
    sigma_df = pd.DataFrame(pd.read_excel(address, sheet_name="Sigma_i^k"))
    r_df = pd.DataFrame(pd.read_excel(address, sheet_name="r_p^s"))
    Cap_df = pd.DataFrame(pd.read_excel(address, sheet_name="Cap_p^s"))

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

    def x_name(i):
        return f"X_{i}"

    def y_name(i, k):
        return f"Y_{i}_{k}"

    def q_name(p, s):
        return f"Q_{p}_{s}"

    X_names = {i: x_name(i) for i in Ip}
    Y_names = {(i, k): y_name(i, k) for i in I for k in K}
    Q_names = {(p, s): q_name(p, s) for p in P for s in S}

    # ========================================================
    # GREEDY HEURISTIC
    # ========================================================

    def allocate_q_for_demand(demand_p):
        q_values = {(p, s): 0.0 for p in P for s in S}
        total_cost = 0.0

        for p in P:
            remaining = demand_p.get(p, 0.0)

            if remaining <= eps:
                continue

            suppliers_sorted = sorted(S, key=lambda s: r[p, s])

            for s in suppliers_sorted:
                if remaining <= eps:
                    break

                amount = min(remaining, Cap[p, s])
                q_values[p, s] = amount
                total_cost += r[p, s] * amount
                remaining -= amount

            if remaining > 1e-6:
                return False, q_values, math.inf

        return True, q_values, total_cost

    def product_is_selectable_under_preference_rule(i, k, selected_products):
        if i not in selected_products:
            return False

        min_sigma = min(sigma[h, k] for h in selected_products)
        return sigma[i, k] <= min_sigma + eps

    def evaluate_selected_products(selected_products):
        selected_products = list(selected_products)

        x_values = {i: 0.0 for i in Ip}
        y_values = {(i, k): 0.0 for i in I for k in K}
        q_values = {(p, s): 0.0 for p in P for s in S}

        for i in selected_products:
            x_values[i] = 1.0

        revenue = 0.0
        demand_p = {p: 0.0 for p in P}
        chosen_product_by_k = {}

        for k in K:
            chosen_i = 0
            best_value = 0.0

            if selected_products:
                feasible_products = [
                    i for i in selected_products
                    if product_is_selectable_under_preference_rule(i, k, selected_products)
                ]

                for i in feasible_products:
                    value = lam[k] * Pi[i]

                    if value > best_value + eps:
                        best_value = value
                        chosen_i = i

            y_values[(chosen_i, k)] = 1.0
            chosen_product_by_k[k] = chosen_i

            if chosen_i != 0:
                revenue += lam[k] * Pi[chosen_i]

                for p in P:
                    demand_p[p] += lam[k] * a[(chosen_i, p)]

        feasible, q_values, purchase_cost = allocate_q_for_demand(demand_p)

        if not feasible:
            return {
                "feasible": False,
                "objective": -math.inf,
                "revenue": revenue,
                "purchase_cost": math.inf,
                "x_values": x_values,
                "y_values": y_values,
                "q_values": q_values,
                "chosen_product_by_k": chosen_product_by_k,
                "demand_p": demand_p,
            }

        return {
            "feasible": True,
            "objective": revenue - purchase_cost,
            "revenue": revenue,
            "purchase_cost": purchase_cost,
            "x_values": x_values,
            "y_values": y_values,
            "q_values": q_values,
            "chosen_product_by_k": chosen_product_by_k,
            "demand_p": demand_p,
        }

    def greedy_forward_addition():
        greedy_start_time = time.perf_counter()

        selected_products = []
        remaining_products = list(Ip)

        current_result = evaluate_selected_products(selected_products)
        current_obj = current_result["objective"]

        greedy_steps = []

        print("\n================ GREEDY FORWARD ADDITION ================")
        print(f"Initial objective with no products = {current_obj:.6f}")

        for step in range(1, V + 1):
            best_candidate = None
            best_candidate_result = None
            best_candidate_obj = current_obj

            for candidate in remaining_products:
                trial_products = selected_products + [candidate]
                trial_result = evaluate_selected_products(trial_products)

                if not trial_result["feasible"]:
                    continue

                trial_obj = trial_result["objective"]

                if trial_obj > best_candidate_obj + eps:
                    best_candidate = candidate
                    best_candidate_result = trial_result
                    best_candidate_obj = trial_obj

            if best_candidate is None:
                print(f"Step {step}: no improving product found. Greedy stops.")
                break

            selected_products.append(best_candidate)
            remaining_products.remove(best_candidate)

            improvement = best_candidate_obj - current_obj
            current_obj = best_candidate_obj
            current_result = best_candidate_result

            greedy_steps.append({
                "step": step,
                "added_product": best_candidate,
                "objective_after_addition": current_obj,
                "improvement": improvement,
                "selected_products": selected_products.copy(),
            })

            print(
                f"Step {step}: added product {best_candidate} | "
                f"objective = {current_obj:.6f} | "
                f"improvement = {improvement:.6f} | "
                f"selected = {selected_products}"
            )

        greedy_end_time = time.perf_counter()

        current_result["greedy_time_seconds"] = greedy_end_time - greedy_start_time
        current_result["selected_products"] = selected_products.copy()

        print("----------------------------------------------------------")
        print(f"Greedy selected products: {selected_products}")
        print(f"Greedy lower bound objective: {current_result['objective']:.6f}")
        print(f"Greedy time: {current_result['greedy_time_seconds']:.4f} seconds")
        print("==========================================================\n")

        return current_result, greedy_steps

    greedy_result, greedy_steps = greedy_forward_addition()

    GREEDY_LOWER_BOUND = greedy_result["objective"]
    greedy_selected_products = greedy_result["selected_products"]
    greedy_x_values = greedy_result["x_values"]
    greedy_y_values = greedy_result["y_values"]
    greedy_q_values = greedy_result["q_values"]

    # ========================================================
    # CPLEX CALLBACK
    # ========================================================

    class ProgressCallback(MIPInfoCallback):
        def __call__(self):
            try:
                current_time = self.get_time()

                if not hasattr(self, "start_time"):
                    self.start_time = current_time

                if not hasattr(self, "last_print_time"):
                    self.last_print_time = self.start_time

                elapsed = current_time - self.start_time

                if current_time - self.last_print_time >= PROGRESS_INTERVAL_SECONDS:
                    self.last_print_time = current_time

                    if self.has_incumbent():
                        incumbent = self.get_incumbent_objective_value()
                        best_bound_cb = self.get_best_objective_value()
                        mip_gap_cb = self.get_MIP_relative_gap()

                        print(
                            f"[CPLEX progress] "
                            f"time = {elapsed:.1f} sec | "
                            f"incumbent = {incumbent:.6f} | "
                            f"best bound = {best_bound_cb:.6f} | "
                            f"gap = {100.0 * mip_gap_cb:.4f}% | "
                            f"nodes = {self.get_num_nodes()}"
                        )
                    else:
                        best_bound_cb = self.get_best_objective_value()

                        print(
                            f"[CPLEX progress] "
                            f"time = {elapsed:.1f} sec | "
                            f"no incumbent yet | "
                            f"best bound = {best_bound_cb:.6f} | "
                            f"nodes = {self.get_num_nodes()}"
                        )
            except Exception:
                pass

    # ========================================================
    # BUILD MODEL
    # ========================================================

    build_start_time = time.perf_counter()

    model = cplex.Cplex()
    model.objective.set_sense(model.objective.sense.maximize)

    model.set_log_stream(sys.stdout)
    model.set_error_stream(sys.stderr)
    model.set_warning_stream(sys.stderr)
    model.set_results_stream(sys.stdout)

    model.parameters.timelimit.set(TIME_LIMIT)
    model.parameters.mip.tolerances.mipgap.set(1e-4)
    model.parameters.mip.tolerances.absmipgap.set(1e-6)
    model.parameters.mip.display.set(4)

    model.register_callback(ProgressCallback)

    var_names = []
    obj = []
    lb = []
    ub = []
    types = []

    for i in Ip:
        var_names.append(X_names[i])
        obj.append(0.0)
        lb.append(0.0)
        ub.append(1.0)
        types.append(model.variables.type.binary)

    for i in I:
        for k in K:
            var_names.append(Y_names[(i, k)])
            obj.append(0.0 if i == 0 else float(lam[k] * Pi[i]))
            lb.append(0.0)
            ub.append(cplex.infinity)
            types.append(model.variables.type.continuous)

    for p in P:
        for s in S:
            var_names.append(Q_names[(p, s)])
            obj.append(float(-r[(p, s)]))
            lb.append(0.0)
            ub.append(cplex.infinity)
            types.append(model.variables.type.continuous)

    model.variables.add(
        obj=obj,
        lb=lb,
        ub=ub,
        types="".join(types),
        names=var_names
    )

    rows = []
    senses = []
    rhs = []
    constraint_names = []

    for k in K:
        rows.append(cplex.SparsePair(
            ind=[Y_names[(i, k)] for i in I],
            val=[1.0] * len(I)
        ))
        senses.append("E")
        rhs.append(1.0)
        constraint_names.append(f"choice_{k}")

    rows.append(cplex.SparsePair(
        ind=[X_names[i] for i in Ip],
        val=[1.0] * len(Ip)
    ))
    senses.append("L")
    rhs.append(float(V))
    constraint_names.append("max_number_products")

    for i in Ip:
        for k in K:
            rows.append(cplex.SparsePair(
                ind=[Y_names[(i, k)], X_names[i]],
                val=[1.0, -1.0]
            ))
            senses.append("L")
            rhs.append(0.0)
            constraint_names.append(f"offer_link_{i}_{k}")

    for i in Ip:
        for k in K:
            ind = [X_names[i]]
            val = [1.0]

            for j in Ip:
                if sigma[(j, k)] > sigma[(i, k)]:
                    ind.append(Y_names[(j, k)])
                    val.append(1.0)

            rows.append(cplex.SparsePair(ind=ind, val=val))
            senses.append("L")
            rhs.append(1.0)
            constraint_names.append(f"preference_{i}_{k}")

    for p in P:
        ind = []
        val = []

        for i in Ip:
            for k in K:
                coeff = lam[k] * a[(i, p)]

                if abs(coeff) > eps:
                    ind.append(Y_names[(i, k)])
                    val.append(float(coeff))

        for s in S:
            ind.append(Q_names[(p, s)])
            val.append(-1.0)

        rows.append(cplex.SparsePair(ind=ind, val=val))
        senses.append("L")
        rhs.append(0.0)
        constraint_names.append(f"demand_cover_{p}")

    for p in P:
        for s in S:
            rows.append(cplex.SparsePair(
                ind=[Q_names[(p, s)]],
                val=[1.0]
            ))
            senses.append("L")
            rhs.append(float(Cap[(p, s)]))
            constraint_names.append(f"capacity_{p}_{s}")

    model.linear_constraints.add(
        lin_expr=rows,
        senses=senses,
        rhs=rhs,
        names=constraint_names
    )

    mip_start_names = []
    mip_start_values = []

    for i in Ip:
        mip_start_names.append(X_names[i])
        mip_start_values.append(float(greedy_x_values[i]))

    for i in I:
        for k in K:
            mip_start_names.append(Y_names[(i, k)])
            mip_start_values.append(float(greedy_y_values[(i, k)]))

    for p in P:
        for s in S:
            mip_start_names.append(Q_names[(p, s)])
            mip_start_values.append(float(greedy_q_values[(p, s)]))

    try:
        model.MIP_starts.add(
            cplex.SparsePair(ind=mip_start_names, val=mip_start_values),
            model.MIP_starts.effort_level.auto,
            "greedy_forward_addition_start"
        )

        print("\n================ MIP START INFORMATION =================")
        print(f"Greedy lower bound passed to CPLEX = {GREEDY_LOWER_BOUND:.6f}")
        print(f"Greedy selected products = {greedy_selected_products}")
        print("=========================================================\n")

    except CplexError as e:
        print("\nWARNING: Could not add greedy MIP start to CPLEX.")
        print("CPLEX error:", e)

    build_end_time = time.perf_counter()

    solve_start_time = time.perf_counter()

    try:
        model.solve()
    except CplexError as e:
        raise RuntimeError(f"CPLEX failed while solving the model: {e}")

    solve_end_time = time.perf_counter()

    status_code = model.solution.get_status()
    status_string = model.solution.get_status_string(status_code)

    if not model.solution.is_primal_feasible():
        total_end_time = time.perf_counter()

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
            "status_code": status_code,
            "status": status_string,
            "time_limit_reached": "time limit" in status_string.lower(),
            "total_time_seconds": total_end_time - total_start_time,
            "solver_time_seconds": solve_end_time - solve_start_time,
            "model_build_time_seconds": build_end_time - build_start_time,
            "greedy_time_seconds": greedy_result["greedy_time_seconds"],
            "greedy_lower_bound": GREEDY_LOWER_BOUND,
            "selected_products": "",
            "greedy_selected_products": ", ".join(map(str, greedy_selected_products)),            
        }

    obj_value = model.solution.get_objective_value()

    try:
        best_bound = model.solution.MIP.get_best_objective()
    except CplexError:
        best_bound = None

    try:
        mip_gap = model.solution.MIP.get_mip_relative_gap()
        mip_gap_percent = 100.0 * mip_gap
    except CplexError:
        mip_gap = None
        mip_gap_percent = None

    time_limit_reached = "time limit" in status_string.lower()

    offered_products = [
        i for i in Ip
        if model.solution.get_values(X_names[i]) > 0.5
    ]

    total_end_time = time.perf_counter()

    print("\n================ CPLEX DIRECT SOLUTION =================")
    print("Final objective value:", obj_value)
    print("Best bound:", best_bound)
    print("Final MIP gap percent:", mip_gap_percent)
    print("CPLEX status:", status_string)
    print("Final offered products:", offered_products)

    # ========================================================
    # EXPORT DATAFRAMES
    # ========================================================

    df_x = pd.DataFrame([
        {
            "i": i,
            "X_value": model.solution.get_values(X_names[i]),
            "greedy_X_value": greedy_x_values[i]
        }
        for i in Ip
    ])

    df_y = pd.DataFrame([
        {
            "i": i,
            "k": k,
            "Y_value": model.solution.get_values(Y_names[(i, k)]),
            "greedy_Y_value": greedy_y_values[(i, k)]
        }
        for i in I
        for k in K
    ])

    df_q = pd.DataFrame([
        {
            "p": p,
            "s": s,
            "Q_value": model.solution.get_values(Q_names[(p, s)]),
            "greedy_Q_value": greedy_q_values[(p, s)]
        }
        for p in P
        for s in S
    ])

    df_summary = pd.DataFrame([
        {"metric": "sample_id", "value": sample_id},
        {"metric": "n", "value": n_value},
        {"metric": "P", "value": P_value},
        {"metric": "K", "value": K_value},
        {"metric": "V", "value": V_value},
        {"metric": "S", "value": S_value},
        {"metric": "greedy_lower_bound", "value": GREEDY_LOWER_BOUND},
        {"metric": "final_objective_value", "value": obj_value},
        {"metric": "best_bound", "value": best_bound},
        {"metric": "mip_gap_percent", "value": mip_gap_percent},
        {"metric": "time_limit_reached", "value": time_limit_reached},
        {"metric": "cplex_status_code", "value": status_code},
        {"metric": "cplex_status", "value": status_string},
        {"metric": "greedy_time_seconds", "value": greedy_result["greedy_time_seconds"]},
        {"metric": "model_build_time_seconds", "value": build_end_time - build_start_time},
        {"metric": "solver_time_seconds", "value": solve_end_time - solve_start_time},
        {"metric": "total_time_seconds", "value": total_end_time - total_start_time},
        {"metric": "time_limit_seconds_per_run", "value": TIME_LIMIT},
        {"metric": "selected_products", "value": ", ".join(map(str, offered_products))},
        {"metric": "greedy_selected_products", "value": ", ".join(map(str, greedy_selected_products))},
    ])

    df_offered_products = pd.DataFrame({"offered_product": offered_products})
    df_greedy_offered_products = pd.DataFrame({"greedy_offered_product": greedy_selected_products})
    df_greedy_steps = pd.DataFrame(greedy_steps)

    df_greedy_chosen_by_k = pd.DataFrame([
        {
            "k": k,
            "greedy_chosen_product": greedy_result["chosen_product_by_k"][k]
        }
        for k in K
    ])

    df_greedy_demand = pd.DataFrame([
        {
            "p": p,
            "greedy_demand": greedy_result["demand_p"][p]
        }
        for p in P
    ])

    df_x_nonzero = df_x[df_x["X_value"].abs() > eps].copy()
    df_y_nonzero = df_y[df_y["Y_value"].abs() > eps].copy()
    df_q_nonzero = df_q[df_q["Q_value"].abs() > eps].copy()

    df_greedy_x_nonzero = df_x[df_x["greedy_X_value"].abs() > eps].copy()
    df_greedy_y_nonzero = df_y[df_y["greedy_Y_value"].abs() > eps].copy()
    df_greedy_q_nonzero = df_q[df_q["greedy_Q_value"].abs() > eps].copy()

    output_dir = os.path.dirname(address)
    output_subdir = os.path.join(output_dir, "CPLEX_subset_results")
    os.makedirs(output_subdir, exist_ok=True)

    detailed_output_file = os.path.join(
        output_subdir,
        f"CPLEX_solution_sample_{sample_id}_n={n_value}_P={P_value}_K={K_value}_V={V_value}_S={S_value}.xlsx"
    )

    with pd.ExcelWriter(detailed_output_file, engine="xlsxwriter") as writer:
        df_summary.to_excel(writer, sheet_name="summary", index=False)
        df_offered_products.to_excel(writer, sheet_name="final_offered_products", index=False)
        df_greedy_offered_products.to_excel(writer, sheet_name="greedy_offered_products", index=False)
        df_greedy_steps.to_excel(writer, sheet_name="greedy_steps", index=False)
        df_greedy_chosen_by_k.to_excel(writer, sheet_name="greedy_chosen_by_k", index=False)
        df_greedy_demand.to_excel(writer, sheet_name="greedy_demand", index=False)

        df_x.to_excel(writer, sheet_name="X_all", index=False)
        df_y.to_excel(writer, sheet_name="Y_all", index=False)
        df_q.to_excel(writer, sheet_name="Q_all", index=False)

        df_x_nonzero.to_excel(writer, sheet_name="X_nonzero", index=False)
        df_y_nonzero.to_excel(writer, sheet_name="Y_nonzero", index=False)
        df_q_nonzero.to_excel(writer, sheet_name="Q_nonzero", index=False)

        df_greedy_x_nonzero.to_excel(writer, sheet_name="greedy_X_nonzero", index=False)
        df_greedy_y_nonzero.to_excel(writer, sheet_name="greedy_Y_nonzero", index=False)
        df_greedy_q_nonzero.to_excel(writer, sheet_name="greedy_Q_nonzero", index=False)

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
        "status_code": status_code,
        "status": status_string,
        "time_limit_reached": time_limit_reached,
        "total_time_seconds": total_end_time - total_start_time,
        "solver_time_seconds": solve_end_time - solve_start_time,
        "model_build_time_seconds": build_end_time - build_start_time,
        "greedy_time_seconds": greedy_result["greedy_time_seconds"],
        "greedy_lower_bound": GREEDY_LOWER_BOUND,
        "selected_products": ", ".join(map(str, offered_products)),
        "greedy_selected_products": ", ".join(map(str, greedy_selected_products)),
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

for sample_id, (n_value, P_value, K_value, V_value, S_value) in enumerate(experiments, start=1):

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
            "mip_gap": None,
            "mip_gap_percent": None,
            "status_code": None,
            "status": f"FAILED: {e}",
            "time_limit_reached": None,
            "total_time_seconds": None,
            "solver_time_seconds": None,
            "model_build_time_seconds": None,
            "greedy_time_seconds": None,
            "greedy_lower_bound": None,
            "selected_products": "",
            "greedy_selected_products": "",
        })

batch_end_time = time.perf_counter()


# ============================================================
# MASTER SUMMARY EXPORT
# ============================================================

summary_df = pd.DataFrame(all_results)

output_dir = os.path.dirname(address)
output_subdir = os.path.join(output_dir, "CPLEX_subset_results")
os.makedirs(output_subdir, exist_ok=True)

summary_output = os.path.join(output_subdir, "CPLEX_all_subset_runs_summary.xlsx")

with pd.ExcelWriter(summary_output, engine="xlsxwriter") as writer:
    summary_df.to_excel(writer, sheet_name="summary_all_runs", index=False)
    format_excel_worksheets(writer)

print("\n====================================================")
print("BATCH RUN FINISHED")
print(f"Number of runs attempted: {len(experiments)}")
print(f"Batch total wall-clock time: {batch_end_time - batch_start_time:.2f} seconds")
print(f"Each CPLEX run had its own time limit: {TIME_LIMIT} seconds")
print(f"Master summary exported to: {summary_output}")
print("====================================================")
