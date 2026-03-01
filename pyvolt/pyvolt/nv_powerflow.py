import numpy as np
import time
from .network import BusType
from .results import Results

def solve(system, tol=1e-8, max_iter=100, time_limit_sec=2.5):
    """
    Newton–Raphson power flow solver with iteration and time limits.

    Parameters
    ----------
    system : network.System
        The network model to solve.
    tol : float
        Convergence tolerance on state update infinity-norm.
    max_iter : int
        Maximum Newton iterations before aborting.
    time_limit_sec : float
        Hard wall-clock limit in seconds (slightly less than parent timeout).

    Returns
    -------
    Results, int
        Power flow results object and number of iterations.

    Raises
    ------
    RuntimeError
        If solver fails to converge within limits or numerical issues occur.
    """

    start_ts = time.time()
    nodes_num = system.get_nodes_num()

    # State vector: [V_re(0..N-1), V_im(0..N-1)]
    state = np.concatenate((np.ones(nodes_num), np.zeros(nodes_num)), axis=0)
    V = state[:nodes_num] + 1j * state[nodes_num:]

    # Pre-size arrays
    z = np.zeros(2 * nodes_num)
    h = np.zeros(2 * nodes_num)
    H = np.zeros((2 * nodes_num, 2 * nodes_num))

    # Apply tap effects before solving
    for branch in system.branches:
        if getattr(branch, "tap_updated", False):
            branch.calculate_tap_effect()

    diff = np.inf
    num_iter = 0

    while diff > tol:
        # Safety guards
        if num_iter >= max_iter:
            raise RuntimeError(f"Max iterations {max_iter} reached without convergence.")
        if (time.time() - start_ts) > time_limit_sec:
            raise RuntimeError(f"Time limit {time_limit_sec}s exceeded in solver.")

        # Reset arrays
        z.fill(0.0)
        h.fill(0.0)
        H.fill(0.0)

        # Build z, h, H
        for node in system.nodes:
            if getattr(node, "ideal_connected_with", "") == "":
                i = node.index
                m = 2 * i
                i2 = i + nodes_num
                node_type = node.type

                if node_type == BusType.SLACK:
                    z[m] = np.real(node.voltage_pu)
                    z[m + 1] = np.imag(node.voltage_pu)
                    H[m, i] = 1.0
                    H[m + 1, i2] = 1.0

                elif node_type == BusType.PQ:
                    reactive_power_total = node.reactive_power + np.imag(node.power_pu)
                    v_abs2 = max(np.abs(V[i]) ** 2, 1e-12)
                    z[m] = (np.real(node.power_pu) * np.real(V[i]) +
                            reactive_power_total * np.imag(V[i])) / v_abs2
                    z[m + 1] = (np.real(node.power_pu) * np.imag(V[i]) -
                                reactive_power_total * np.real(V[i])) / v_abs2
                    H[m, :nodes_num] = np.real(system.Ymatrix[i])
                    H[m, nodes_num:] = -np.imag(system.Ymatrix[i])
                    H[m + 1, :nodes_num] = np.imag(system.Ymatrix[i])
                    H[m + 1, nodes_num:] = np.real(system.Ymatrix[i])

                elif node_type == BusType.PV:
                    h[m] = (np.real(V[i]) *
                            (np.inner(np.real(system.Ymatrix[i]), np.real(V)) -
                             np.inner(np.imag(system.Ymatrix[i]), np.imag(V))) +
                            np.imag(V[i]) *
                            (np.inner(np.real(system.Ymatrix[i]), np.imag(V)) +
                             np.inner(np.imag(system.Ymatrix[i]), np.real(V))))
                    z[m] = np.real(node.power_pu)
                    h[m + 1] = np.abs(V[i])
                    z[m + 1] = np.abs(node.voltage_pu)

                    H[m, :nodes_num] = (np.real(V) * np.real(system.Ymatrix[i]) +
                                        np.imag(V) * np.imag(system.Ymatrix[i]))
                    H[m, i] += (np.inner(np.real(system.Ymatrix[i]), np.real(V)) -
                                np.inner(np.imag(system.Ymatrix[i]), np.imag(V)))
                    H[m, nodes_num:] = (np.imag(V) * np.real(system.Ymatrix[i]) -
                                        np.real(V) * np.imag(system.Ymatrix[i]))
                    H[m, i2] += (np.inner(np.real(system.Ymatrix[i]), np.imag(V)) +
                                 np.inner(np.imag(system.Ymatrix[i]), np.real(V)))

                    v_abs = max(np.abs(V[i]), 1e-12)
                    cos_ang = np.real(V[i]) / v_abs
                    sin_ang = np.imag(V[i]) / v_abs
                    H[m + 1, i] = cos_ang
                    H[m + 1, i2] = sin_ang

        # Fill h for Slack/PQ rows
        for node in system.nodes:
            if getattr(node, "ideal_connected_with", "") == "":
                i = node.index
                m = 2 * i
                i2 = i + nodes_num
                node_type = node.type
                if node_type in (BusType.SLACK, BusType.PQ):
                    h[m] = np.inner(H[m], state)
                    h[m + 1] = np.inner(H[m + 1], state)

        # Residual
        r = z - h

        if not np.all(np.isfinite(H)) or not np.all(np.isfinite(r)):
            raise RuntimeError("Non-finite values in Jacobian or residual.")

        # Solve for update
        try:
            delta_state = np.linalg.solve(H, r)
        except np.linalg.LinAlgError:
            reg = 1e-10
            delta_state = np.linalg.solve(H + reg * np.eye(H.shape[0]), r)

        state += delta_state
        diff = np.max(np.abs(delta_state))
        V = state[:nodes_num] + 1j * state[nodes_num:]
        num_iter += 1

    # Build results
    powerflow_results = Results(system)
    powerflow_results.load_voltages(V)
    powerflow_results.calculate_all()

    return powerflow_results, num_iter
