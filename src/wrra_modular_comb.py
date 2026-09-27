#!/usr/bin/env python3
"""WRRA modular prime-sieve optical-comb simulation.

This script tests an exact arithmetic construction and a deliberately simple
hardware-surrogate noise model. It does not simulate a complete laser,
modulator, detector, or feedback loop.
"""

from __future__ import annotations

import json
import math
from dataclasses import asdict, dataclass
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
RESULTS = ROOT / "results"
FIGURES = ROOT / "figures"
SEED = 20260927


def primes_up_to(n: int) -> np.ndarray:
    sieve = np.ones(n + 1, dtype=bool)
    sieve[:2] = False
    for p in range(2, int(math.isqrt(n)) + 1):
        if sieve[p]:
            sieve[p * p : n + 1 : p] = False
    return np.flatnonzero(sieve)


def prime_mask(n_max: int) -> np.ndarray:
    mask = np.zeros(n_max + 1, dtype=float)
    mask[primes_up_to(n_max)] = 1.0
    return mask


def base_primes(n_max: int) -> np.ndarray:
    return primes_up_to(math.isqrt(n_max))


def ideal_selector(n: np.ndarray, p: int) -> np.ndarray:
    """DFT selector equal to 1 when p divides n, and 0 otherwise."""
    k = np.arange(p)
    return np.exp(2j * np.pi * np.outer(n, k) / p).mean(axis=1)


def modular_field(
    n_max: int,
    rng: np.random.Generator | None = None,
    coeff_sigma: float = 0.0,
    phase_sigma_rad: float = 0.0,
    delay_fraction_sigma: float = 0.0,
) -> np.ndarray:
    """Return the cascaded modular-filter field on integer comb lines.

    The nominal p-stage response is

        H_p(n) = 1 - (1/p) sum_{k=0}^{p-1} exp(2 pi i k n / p).

    For n <= N, cascading stages for p <= sqrt(N) removes every composite.
    Each base prime is also removed by its own stage, so an ideal protected-line
    restoration is applied after the cascade.

    Noise terms are hardware surrogates:
      * coeff_sigma: fractional complex-tap amplitude calibration error.
      * phase_sigma_rad: fixed phase-offset error per tap.
      * delay_fraction_sigma: fractional delay error; its phase grows with n.
    """
    if rng is None:
        rng = np.random.default_rng(SEED)

    n = np.arange(n_max + 1, dtype=float)
    field = np.ones(n_max + 1, dtype=complex)
    field[:2] = 0.0

    for p in base_primes(n_max):
        k = np.arange(p, dtype=float)
        amp = 1.0 + rng.normal(0.0, coeff_sigma, p)
        phase_offset = rng.normal(0.0, phase_sigma_rad, p)
        delay_fraction = rng.normal(0.0, delay_fraction_sigma, p)

        nominal_phase = 2.0 * np.pi * np.outer(n, k) / p
        phase = nominal_phase * (1.0 + delay_fraction[None, :])
        phase += phase_offset[None, :]
        selector = (amp[None, :] * np.exp(1j * phase)).mean(axis=1)
        field *= 1.0 - selector

    # Protected-line restoration. Its physical implementation and error budget
    # are excluded from this first study and listed as a limitation.
    field[base_primes(n_max)] = 1.0 + 0.0j
    field[:2] = 0.0
    return field


@dataclass
class Metrics:
    relative_l2_error: float
    fourier_power_correlation: float
    mean_composite_leakage: float
    contrast_db: float
    prime_power_p05: float
    prime_power_cv: float
    precision: float
    recall: float
    f1: float


def measure(field: np.ndarray, target: np.ndarray) -> Metrics:
    x = np.asarray(field[2:], dtype=complex)
    y = np.asarray(target[2:], dtype=complex)
    prime = y.real > 0.5
    composite = ~prime

    denom = np.vdot(x, x)
    alpha = np.vdot(x, y) / denom if abs(denom) > 0 else 0.0
    x_aligned = alpha * x
    relative_l2_error = float(np.linalg.norm(x_aligned - y) / np.linalg.norm(y))

    pwr = np.abs(x) ** 2
    prime_mean = float(np.mean(pwr[prime])) if np.any(prime) else 1.0
    pwr_norm = pwr / max(prime_mean, 1e-300)
    mean_composite = float(np.mean(pwr_norm[composite]))
    contrast_db = float(10.0 * np.log10(1.0 / max(mean_composite, 1e-300)))
    prime_p05 = float(np.quantile(pwr_norm[prime], 0.05))
    prime_cv = float(np.std(pwr_norm[prime]) / max(np.mean(pwr_norm[prime]), 1e-300))

    pred = pwr_norm >= 0.5
    tp = int(np.count_nonzero(pred & prime))
    fp = int(np.count_nonzero(pred & composite))
    fn = int(np.count_nonzero((~pred) & prime))
    precision = tp / (tp + fp) if tp + fp else 0.0
    recall = tp / (tp + fn) if tp + fn else 0.0
    f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0

    ideal_fft = np.abs(np.fft.fft(y)) ** 2
    actual_fft = np.abs(np.fft.fft(x_aligned)) ** 2
    if np.std(actual_fft) == 0 or np.std(ideal_fft) == 0:
        corr = 0.0
    else:
        corr = float(np.corrcoef(ideal_fft, actual_fft)[0, 1])

    return Metrics(
        relative_l2_error=relative_l2_error,
        fourier_power_correlation=corr,
        mean_composite_leakage=mean_composite,
        contrast_db=contrast_db,
        prime_power_p05=prime_p05,
        prime_power_cv=prime_cv,
        precision=precision,
        recall=recall,
        f1=f1,
    )


def budget_matched_individual_field(n_max: int, budget: int) -> np.ndarray:
    """Best-case individual suppression with only `budget` binary controls."""
    target = prime_mask(n_max)
    field = np.ones(n_max + 1, dtype=complex)
    field[:2] = 0.0
    composites = np.flatnonzero((np.arange(n_max + 1) >= 2) & (target == 0))
    field[composites[: min(budget, len(composites))]] = 0.0
    return field


def uniform_group_field(n_max: int, groups: int) -> np.ndarray:
    """Piecewise-constant least-squares mask on equal-width spectral groups."""
    target = prime_mask(n_max)
    y = target[2:]
    out = np.zeros_like(target, dtype=complex)
    for idx in np.array_split(np.arange(len(y)), groups):
        if len(idx):
            out[idx + 2] = np.mean(y[idx])
    return out


def complexity_table(n_values: list[int]) -> pd.DataFrame:
    rows = []
    for n_max in n_values:
        bp = base_primes(n_max)
        pi_n = len(primes_up_to(n_max))
        line_weights = n_max - 1
        composites = line_weights - pi_n
        high_level_actions = 2 * len(bp)
        optical_paths = int(np.sum(bp + 1) + len(bp))
        ideal = measure(modular_field(n_max), prime_mask(n_max))
        budget_individual = measure(
            budget_matched_individual_field(n_max, high_level_actions),
            prime_mask(n_max),
        )
        uniform = measure(uniform_group_field(n_max, high_level_actions), prime_mask(n_max))
        rows.append(
            {
                "N": n_max,
                "prime_count": pi_n,
                "composite_count": composites,
                "base_primes": ",".join(map(str, bp)),
                "base_prime_count": len(bp),
                "direct_line_weights": line_weights,
                "direct_composite_suppressions": composites,
                "wrra_high_level_actions": high_level_actions,
                "wrra_optical_paths_surrogate": optical_paths,
                "action_reduction_vs_line_weights_pct": 100.0 * (1.0 - high_level_actions / line_weights),
                "path_reduction_vs_line_weights_pct": 100.0 * (1.0 - optical_paths / line_weights),
                "ideal_relative_l2_error": ideal.relative_l2_error,
                "ideal_fourier_correlation": ideal.fourier_power_correlation,
                "budget_individual_f1": budget_individual.f1,
                "uniform_group_f1": uniform.f1,
                "wrra_ideal_f1": ideal.f1,
            }
        )
    return pd.DataFrame(rows)


def ideal_ledger(n_max: int) -> pd.DataFrame:
    n = np.arange(n_max + 1)
    target = prime_mask(n_max)
    field = np.ones(n_max + 1, dtype=complex)
    field[:2] = 0.0
    rows = []
    for stage, p in enumerate(base_primes(n_max), start=1):
        before = np.abs(field) > 0.5
        h = 1.0 - ideal_selector(n, int(p))
        field *= h
        after = np.abs(field) > 0.5
        newly_removed = before & (~after)
        newly_removed_composites = newly_removed & (target == 0) & (n >= 2)
        remaining_composites = after & (target == 0) & (n >= 2)
        rows.append(
            {
                "N": n_max,
                "stage": stage,
                "prime_filter_p": int(p),
                "newly_removed_total": int(np.count_nonzero(newly_removed & (n >= 2))),
                "newly_removed_composites": int(np.count_nonzero(newly_removed_composites)),
                "protected_prime_temporarily_removed": int(newly_removed[int(p)]),
                "remaining_composites": int(np.count_nonzero(remaining_composites)),
            }
        )
    return pd.DataFrame(rows)


def run_noise_sweep(
    n_values: list[int],
    trials_by_n: dict[int, int],
    coefficient_phase_levels: list[float],
    delay_levels: list[float],
) -> pd.DataFrame:
    rng = np.random.default_rng(SEED)
    rows: list[dict[str, float | int | str]] = []

    def add_trials(n_max: int, kind: str, level: float, trials: int) -> None:
        target = prime_mask(n_max)
        for trial in range(trials):
            if kind == "coefficient_phase":
                field = modular_field(
                    n_max,
                    rng=rng,
                    coeff_sigma=level,
                    phase_sigma_rad=level,
                )
            elif kind == "delay_fraction":
                field = modular_field(
                    n_max,
                    rng=rng,
                    delay_fraction_sigma=level,
                )
            else:
                raise ValueError(kind)
            rows.append(
                {
                    "N": n_max,
                    "noise_kind": kind,
                    "noise_level": level,
                    "trial": trial,
                    **asdict(measure(field, target)),
                }
            )

    for n_max in n_values:
        trials = trials_by_n[n_max]
        for level in coefficient_phase_levels:
            add_trials(n_max, "coefficient_phase", level, trials)
        for level in delay_levels:
            add_trials(n_max, "delay_fraction", level, trials)

    return pd.DataFrame(rows)


def summarize_noise(raw: pd.DataFrame) -> pd.DataFrame:
    metric_cols = [
        "relative_l2_error",
        "fourier_power_correlation",
        "mean_composite_leakage",
        "contrast_db",
        "prime_power_p05",
        "prime_power_cv",
        "precision",
        "recall",
        "f1",
    ]
    grouped = raw.groupby(["N", "noise_kind", "noise_level"])[metric_cols]
    mean = grouped.mean().add_suffix("_mean")
    q05 = grouped.quantile(0.05).add_suffix("_q05")
    q95 = grouped.quantile(0.95).add_suffix("_q95")
    return pd.concat([mean, q05, q95], axis=1).reset_index()


def plot_complexity(df: pd.DataFrame) -> None:
    fig, ax = plt.subplots(figsize=(8.3, 5.0))
    ax.plot(df["N"], df["direct_line_weights"], "o-", label="Independent line weights")
    ax.plot(df["N"], df["direct_composite_suppressions"], "o-", label="Individual composite suppressions")
    ax.plot(df["N"], df["wrra_optical_paths_surrogate"], "o-", label="Modular optical paths surrogate")
    ax.plot(df["N"], df["wrra_high_level_actions"], "o-", label="WRRA high-level actions")
    ax.set_xscale("log")
    ax.set_yscale("log")
    ax.set_xlabel("Largest represented integer N")
    ax.set_ylabel("Count")
    ax.set_title("Control and hardware complexity")
    ax.grid(True, which="both", alpha=0.25)
    ax.legend(frameon=False)
    fig.tight_layout()
    fig.savefig(FIGURES / "figure_1_control_complexity.png", dpi=220)
    plt.close(fig)


def plot_ideal_mask(n_max: int = 210) -> None:
    target = prime_mask(n_max)
    field = modular_field(n_max)
    n = np.arange(2, n_max + 1)
    fig, axes = plt.subplots(2, 1, figsize=(9.2, 5.6), sharex=True, height_ratios=[2.2, 1.0])
    axes[0].vlines(n[target[2:] > 0.5], 0, 1, color="#1f4e79", lw=1.1, label="Prime target")
    axes[0].scatter(n, np.abs(field[2:]) ** 2, s=8, color="#d95f02", alpha=0.75, label="Modular-filter output")
    axes[0].set_ylabel("Normalized line power")
    axes[0].set_title(f"Ideal modular sieve on N = {n_max} comb lines")
    axes[0].legend(frameon=False, ncol=2)
    error = np.abs(np.abs(field[2:]) ** 2 - target[2:])
    axes[1].semilogy(n, np.maximum(error, 1e-18), color="#7f7f7f", lw=0.9)
    axes[1].set_ylabel("Absolute error")
    axes[1].set_xlabel("Arithmetic index mapped to comb line")
    axes[1].grid(True, alpha=0.25)
    fig.tight_layout()
    fig.savefig(FIGURES / "figure_2_ideal_exactness.png", dpi=220)
    plt.close(fig)


def plot_noise(summary: pd.DataFrame, kind: str, filename: str, x_label: str) -> None:
    data = summary[(summary["noise_kind"] == kind) & (summary["noise_level"] > 0)]
    fig, axes = plt.subplots(1, 2, figsize=(10.2, 4.3))
    for n_max, g in data.groupby("N"):
        g = g.sort_values("noise_level")
        x = g["noise_level"].to_numpy(float)
        axes[0].semilogx(x, g["f1_mean"], "o-", label=f"N={n_max}")
        axes[0].fill_between(x, g["f1_q05"], g["f1_q95"], alpha=0.12)
        axes[1].semilogx(x, g["contrast_db_mean"], "o-", label=f"N={n_max}")
        axes[1].fill_between(x, g["contrast_db_q05"], g["contrast_db_q95"], alpha=0.12)
    axes[0].set_xlabel(x_label)
    axes[0].set_ylabel("Prime-mask F1 score")
    axes[0].set_ylim(-0.02, 1.02)
    axes[1].set_xlabel(x_label)
    axes[1].set_ylabel("Prime/composite contrast dB")
    for ax in axes:
        ax.grid(True, which="both", alpha=0.25)
    axes[0].legend(frameon=False)
    fig.tight_layout()
    fig.savefig(FIGURES / filename, dpi=220)
    plt.close(fig)


def main() -> None:
    RESULTS.mkdir(parents=True, exist_ok=True)
    FIGURES.mkdir(parents=True, exist_ok=True)

    n_values = [30, 210, 2310]
    complexity = complexity_table(n_values)
    complexity.to_csv(RESULTS / "complexity_and_ideal_results.csv", index=False)

    ledgers = pd.concat([ideal_ledger(n) for n in n_values], ignore_index=True)
    ledgers.to_csv(RESULTS / "ideal_stage_ledger.csv", index=False)

    coefficient_phase_levels = [0.0, 1e-4, 3e-4, 1e-3, 3e-3, 1e-2, 3e-2, 1e-1, 3e-1]
    delay_levels = [0.0, 1e-8, 3e-8, 1e-7, 3e-7, 1e-6, 3e-6, 1e-5, 3e-5, 1e-4, 3e-4, 1e-3]
    trials_by_n = {30: 400, 210: 200, 2310: 40}
    raw = run_noise_sweep(n_values, trials_by_n, coefficient_phase_levels, delay_levels)
    raw.to_csv(RESULTS / "noise_monte_carlo_raw.csv", index=False)
    summary = summarize_noise(raw)
    summary.to_csv(RESULTS / "noise_monte_carlo_summary.csv", index=False)

    plot_complexity(complexity)
    plot_ideal_mask(210)
    plot_noise(
        summary,
        "coefficient_phase",
        "figure_3_coefficient_phase_robustness.png",
        "Coefficient sigma and phase sigma in radians",
    )
    plot_noise(
        summary,
        "delay_fraction",
        "figure_4_delay_robustness.png",
        "Fractional delay-error sigma",
    )

    mean_thresholds = {}
    conservative_thresholds = {}
    for kind in ["coefficient_phase", "delay_fraction"]:
        mean_thresholds[kind] = {}
        conservative_thresholds[kind] = {}
        for n_max in n_values:
            g = summary[(summary["noise_kind"] == kind) & (summary["N"] == n_max)].sort_values("noise_level")
            passing_mean = g[(g["f1_mean"] >= 0.95) & (g["contrast_db_mean"] >= 20.0)]
            passing_conservative = g[
                (g["f1_mean"] >= 0.95)
                & (g["f1_q05"] >= 0.95)
                & (g["contrast_db_q05"] >= 20.0)
            ]
            mean_thresholds[kind][str(n_max)] = (
                float(passing_mean["noise_level"].max()) if len(passing_mean) else None
            )
            conservative_thresholds[kind][str(n_max)] = (
                float(passing_conservative["noise_level"].max()) if len(passing_conservative) else None
            )

    payload = {
        "seed": SEED,
        "n_values": n_values,
        "trials_by_n": trials_by_n,
        "coefficient_phase_levels": coefficient_phase_levels,
        "delay_fraction_levels": delay_levels,
        "mean_pass_rule": "mean F1 >= 0.95 and mean contrast >= 20 dB",
        "conservative_pass_rule": "mean F1 >= 0.95, 5th-percentile F1 >= 0.95, and 5th-percentile contrast >= 20 dB",
        "largest_tested_passing_level_mean": mean_thresholds,
        "largest_tested_passing_level_conservative": conservative_thresholds,
        "model_scope": {
            "included": [
                "ideal DFT residue-class filters",
                "tap amplitude calibration error",
                "fixed tap phase-offset error",
                "fractional delay error with line-index phase accumulation",
                "prime/composite discrimination and Fourier-power observables",
            ],
            "excluded": [
                "laser dynamics",
                "nonlinear propagation",
                "detector and feedback electronics",
                "insertion loss and thermal crosstalk",
                "error in protected-prime restoration",
            ],
        },
    }
    (RESULTS / "run_summary.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")

    print(complexity.to_string(index=False))
    print("\nLargest tested passing noise levels:")
    print("Mean rule:")
    print(json.dumps(mean_thresholds, indent=2))
    print("Conservative rule:")
    print(json.dumps(conservative_thresholds, indent=2))


if __name__ == "__main__":
    main()
