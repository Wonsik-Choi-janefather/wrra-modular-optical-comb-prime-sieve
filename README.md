# WRRA Modular Optical-Comb Prime Sieve

Reproducible arithmetic and reduced transfer-function study for the report:

> **WRRA 기반 모듈러 광주파수 빗 소수체 구현 연구**  
> *Modular residue-class filters, control complexity, and error sensitivity*

Author: **Wonsik Choi (최원식)**  
Study date: **2026-09-27**

## Abstract

This repository tests a WRRA-style implementation in which one periodic
common-carrier comb is transformed by a small family of modular filters. For
each prime `p`, the ideal stage

$$
H_p(n)=1-\frac{1}{p}\sum_{k=0}^{p-1}\exp\!\left(\frac{2\pi i k n}{p}\right)
$$

rejects comb indices divisible by `p`. Cascading the stages for all
`p <= sqrt(N)`, followed by exact restoration of the small prime lines, produces
the prime indicator on `2,...,N` in the arithmetic model. The code also tests a
deliberately reduced hardware-surrogate model with coefficient, phase, and
fractional-delay errors.

This is a **toy-model implementation study**, not a claim that a complete
optical device has been demonstrated and not a new theorem about prime-number
distribution.

## WRRA evaluation chain

| Verified input | WRRA transform | Output | Falsification condition |
|---|---|---|---|
| Integer comb labels `n=2,...,N`; standard roots-of-unity identity | Common-carrier comb passed through modular DFT residue filters `H_p`, then protected-line restoration | Exact prime mask in the ideal arithmetic model | Any composite survives or any prime is absent beyond floating-point tolerance |
| Published/explicit numerical settings in `results/run_summary.json` | Same internal cascade with sampled coefficient, phase, or delay errors | F1, contrast, leakage, Fourier-power correlation, and uncertainty intervals | Re-running with seed `20260927` fails to reproduce the committed tables within numerical tolerance |
| Direct per-line control and equal-width grouped-control baselines | Compare control descriptions at equal high-level-action count | Action/path complexity and discrimination comparison | A baseline encoded by the stated rules outperforms the reported values when run from the same source |
| Engineering target: `F1 >= 0.95` and contrast `>= 20 dB` | Sweep error amplitudes for `N = 30, 210, 2310` | Largest tested passing error level | A repeated sweep under the same stochastic model moves the pass boundary outside the reported Monte Carlo uncertainty |

The reported agreement with known prime masks is treated here as a reality-
consistency and explanatory test of the internal WRRA transformation. Claims
about a physical instrument remain conditional on the explicitly excluded
hardware effects.

## Main numerical results

| `N` | Direct line weights | WRRA high-level actions | Path surrogate | Ideal relative L2 error | Ideal F1 |
|---:|---:|---:|---:|---:|---:|
| 30 | 29 | 6 | 16 | `3.911e-15` | 1.000 |
| 210 | 209 | 12 | 53 | `3.843e-14` | 1.000 |
| 2310 | 2309 | 30 | 358 | `4.497e-13` | 1.000 |

For `N=2310`, fractional-delay error `sigma=1e-4` gave mean `F1=0.911165`
and mean contrast `19.756 dB`. Under the conservative rule (mean F1, 5th-
percentile F1, and 5th-percentile contrast all passing), the largest tested
delay-error level was `3e-5`. These are results of the reduced surrogate, not
fabrication tolerances for a finished instrument.

## Repository contents

- `src/wrra_modular_comb.py` — exact construction, baselines, Monte Carlo
  sweeps, metrics, and figures.
- `results/` — committed ideal results, stage ledger, raw trials, aggregated
  robustness tables, and the run manifest.
- `figures/` — figures and rendered equations used in the report.
- `report/` — the complete Korean report in PDF and DOCX formats.
- `src/render_equations.py` — regenerates the equation images.
- `src/build_report.py` — regenerates the PDF/DOCX after the numerical run.

## Reproduce the numerical study

Python 3.11+ is recommended.

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\\Scripts\\activate
python -m pip install -r requirements.txt
python src/wrra_modular_comb.py
```

The stochastic sweep is deterministic under seed `20260927`. It may take
several minutes, especially for `N=2310`.

To rebuild the report after regenerating results:

```bash
python src/render_equations.py
python src/build_report.py
```

The report builder requires Korean-capable TrueType fonts. Set
`WRRA_FONT_REGULAR`, `WRRA_FONT_BOLD`, and optionally `WRRA_DOCX_FONT` if the
default Noto Sans KR paths are not present.

## Model boundary

Included:

- ideal DFT residue-class filters;
- tap-amplitude calibration error;
- fixed tap phase-offset error;
- fractional-delay error whose phase accumulates with comb index;
- prime/composite discrimination and Fourier-power observables.

Not included:

- laser dynamics and nonlinear propagation;
- detector and feedback electronics;
- insertion loss and thermal crosstalk;
- error in the protected-prime restoration channels;
- a fabricated device or laboratory measurement.

Accordingly, the calculation establishes exactness inside the stated model and
identifies delay alignment as the dominant tested bottleneck. The next empirical
step is a small `N=30` pilot implementation.

## Related work

- Related mathematical study: [*Prime-Sieve State-Space Renormalization and Irreducibility: Projective Lifting, Information Increment, and Falsifiable Physical Mappings*](https://doi.org/10.5281/zenodo.21869614), version 2.7.1. This DOI is **not** the DOI of the modular optical-comb report or this software package; no DOI for this exact work has been verified.

## Citation

Citation metadata is provided in `CITATION.cff`. This repository does not yet
assign a software or document license; reuse beyond normal citation therefore
requires the author's permission.

## Copyright

Copyright (C) 2026 Wonsik Choi

---

## Central corpus index

This work is part of the open research and publishing corpus of **Wonsik Choi (최원식)**.

- [Central Research & Publications Index](https://github.com/Wonsik-Choi-janefather/minimal-computing-cosmology-research-history/blob/main/PUBLICATIONS.md)
- [Public GitBook index](https://independent-research.gitbook.io/mcc-and-wrra-research-history/publications)
- [Machine-readable corpus index](https://github.com/Wonsik-Choi-janefather/minimal-computing-cosmology-research-history/blob/main/works.json)
- Identity: [janefather@gmail.com](mailto:janefather@gmail.com)

Rights remain those stated in this repository and its linked archival record.

**Copyright (C) 2026 Wonsik Choi**


**ORCID:** [0009-0001-4263-9772](https://orcid.org/0009-0001-4263-9772)
