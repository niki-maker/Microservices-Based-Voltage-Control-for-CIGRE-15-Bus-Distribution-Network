import numpy as np
import logging
import multiprocessing

# ---------------- Timeout-enforced power flow ---------------- #
def _solve_worker(system, return_dict):
    try:
        from pyvolt.nv_powerflow import solve
        results_pf, _ = solve(system)
        return_dict["results"] = results_pf
    except Exception as e:
        return_dict["error"] = str(e)

def safe_powerflow(system, timeout=3.0):
    manager = multiprocessing.Manager()
    return_dict = manager.dict()
    process = multiprocessing.Process(target=_solve_worker, args=(system, return_dict))
    process.start()
    process.join(timeout)

    if process.is_alive():
        process.terminate()
        process.join()
        logging.warning("⏱️ Power flow forcibly terminated after %.1f seconds.", timeout)
        return None

    if "error" in return_dict:
        logging.warning("⚠️ Power flow failed: %s", return_dict["error"])
        return None

    return return_dict.get("results", None)
# ---------------------------------------------------------------- #

# ---------------- Combined Objective Function ---------------- #
def combined_cap_reac_objective_function(solution, system,
                                         capacitor_reactive_power,
                                         shunt_reactor_reactive_power,
                                         base_apparent_power):
    """
    Applies both capacitor and reactor states from a single decision vector.
    First len(capacitor_reactive_power) bits are capacitors (inject +Q),
    next bits are reactors (inject -Q).
    """
    n_caps = len(capacitor_reactive_power)
    binary_solution = [1 if s >= 0.5 else 0 for s in solution]
    cap_bits = binary_solution[:n_caps]
    reac_bits = binary_solution[n_caps:]

    # Reset all devices
    for node_name in capacitor_reactive_power.keys():
        node = system.get_node_by_uuid(node_name)
        if node:
            node.reactive_power = 0.0
    for node_name in shunt_reactor_reactive_power.keys():
        node = system.get_node_by_uuid(node_name)
        if node:
            node.reactive_power = 0.0

    # Apply capacitors
    for idx, (node_name, q_mvar) in enumerate(capacitor_reactive_power.items()):
        if cap_bits[idx] == 1:
            node = system.get_node_by_uuid(node_name)
            if node:
                node.reactive_power = q_mvar / base_apparent_power

    # Apply reactors
    for idx, (node_name, q_mvar) in enumerate(shunt_reactor_reactive_power.items()):
        if reac_bits[idx] == 1:
            node = system.get_node_by_uuid(node_name)
            if node:
                node.reactive_power = -(q_mvar / base_apparent_power)

    results_pf = safe_powerflow(system, timeout=3.0)
    if results_pf is None:
        return [float("inf"), float("inf")]

    voltages = {n.topology_node.name: abs(n.voltage_pu) for n in results_pf.nodes}
    voltage_deviation = sum((max(v - 1.05, 0) + max(0.95 - v, 0)) ** 2 for v in voltages.values())
    wear_and_tear = sum(binary_solution)
    return [voltage_deviation, wear_and_tear]
# ---------------------------------------------------------------- #

# ---------------- Pareto archive utilities ---------------- #
def update_pareto_archive(new_solution, archive):
    to_remove = []
    for idx, archived in enumerate(archive):
        if dominates(archived[-2:], new_solution[-2:]):
            return archive
        elif dominates(new_solution[-2:], archived[-2:]):
            to_remove.append(idx)
    archive = [archived for i, archived in enumerate(archive) if i not in to_remove]
    archive.append(new_solution)
    return archive

def dominates(obj1, obj2):
    return all(o1 <= o2 for o1, o2 in zip(obj1, obj2)) and any(o1 < o2 for o1, o2 in zip(obj1, obj2))

def extract_pareto_front(archive):
    non_dominated = []
    for sol in archive:
        if not any(dominates(other[-2:], sol[-2:]) for other in archive if other is not sol):
            non_dominated.append(sol)
    return non_dominated
# ----------------------------------------------------------- #

# ---------------- Fuzzy selection ---------------- #
def select_best_fuzzy(objectives):
    finite_mask = np.isfinite(objectives).all(axis=1)
    if not np.any(finite_mask):
        return 0
    objs = objectives[finite_mask]
    weights = np.array([0.95, 0.05])
    f_min = objs.min(axis=0)
    f_max = objs.max(axis=0)
    epsilon = 1e-9
    normalized_range = (f_max - f_min) + epsilon
    membership = (f_max - objs) / normalized_range
    fuzzy_scores = membership @ weights
    finite_indices = np.where(finite_mask)[0]
    return finite_indices[np.argmax(fuzzy_scores)]
# ----------------------------------------------- #

# ---------------- Optimizer ---------------- #
def initialize_population(pop_size, dim, lb, ub):
    return np.random.uniform(lb, ub, (pop_size, dim))

def fungal_growth_optimizer(N, Tmax, ub, lb, dim, fobj):
    M, Ep, R = 0.6, 0.7, 0.9
    S = initialize_population(N, dim, lb, ub)
    pareto_archive = []

    # Cache to avoid re-evaluating same binary pattern
    obj_cache = {}

    def eval_with_cache(sol):
        key = tuple(1 if v >= 0.5 else 0 for v in sol)
        if key in obj_cache:
            return obj_cache[key]
        val = fobj(sol)
        if val is None:
            val = [float("inf"), float("inf")]
        obj_cache[key] = val
        return val

    # Evaluate initial population
    for sol in S:
        objs = eval_with_cache(sol)
        pareto_archive = update_pareto_archive(np.hstack((sol, objs)), pareto_archive)

    for t in range(Tmax):
        nutrients = np.random.rand(N) if t <= Tmax / 2 else np.array([sol[-2] for sol in pareto_archive])
        nutrients /= (np.sum(nutrients) + 2 * np.random.rand())

        for i in range(N):
            a, b, c = np.random.choice([x for x in range(N) if x != i], 3, replace=False)
            p = 0.0
            Er = M + (1 - t / Tmax) * (1 - M)

            if p < Er:
                F = np.random.rand() * (1 - t / Tmax) ** (1 - t / Tmax)
                E = np.exp(F)
                r1, r2 = np.random.rand(dim), np.random.rand()
                U1 = r1 < r2
                S[i] = U1 * S[i] + (1 - U1) * (S[i] + E * (S[a] - S[b]))
            else:
                Ec = (np.random.rand(dim) - 0.5) * np.random.rand() * (S[a] - S[b])
                if np.random.rand() < np.random.rand():
                    De2 = np.random.rand(dim) * (S[i] - S[c]) * (np.random.rand(dim) > np.random.rand())
                    S[i] += De2 * nutrients[i] + Ec * (np.random.rand() > np.random.rand())
                else:
                    De = (np.random.rand() * (S[a] - S[i]) +
                          np.random.rand(dim) *
                          ((np.random.rand() > (np.random.rand() * 2 - 1)) * S[c] - S[i]) *
                          (np.random.rand() > R))
                    S[i] += De * nutrients[i] + Ec * (np.random.rand() > Ep)

            S[i] = np.clip(S[i], lb, ub)
            objs = eval_with_cache(S[i])
            pareto_archive = update_pareto_archive(np.hstack((S[i], objs)), pareto_archive)

        if len(obj_cache) >= (2 ** dim):
            break

    pareto_front = extract_pareto_front(pareto_archive)
    if not pareto_front:
        dummy_sol = np.zeros(dim)
        pareto_front = [np.hstack((dummy_sol, [float("inf"), float("inf")]))]

    objectives = np.array([sol[-2:] for sol in pareto_front])
    best_idx = select_best_fuzzy(objectives)
    best_solution = pareto_front[best_idx]
    return pareto_front, best_solution
# ------------------------------------------ #
