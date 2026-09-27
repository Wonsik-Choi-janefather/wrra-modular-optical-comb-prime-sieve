#!/usr/bin/env python3
"""Render publication equations as high-resolution transparent PNGs."""

from pathlib import Path

import matplotlib.pyplot as plt


OUT = Path(__file__).resolve().parents[1] / "figures" / "equations"
OUT.mkdir(parents=True, exist_ok=True)

EQUATIONS = {
    "equation_1_filter": r"$H_p(n)=1-\frac{1}{p}\sum_{k=0}^{p-1}\exp\!\left(\frac{2\pi i k n}{p}\right)$",
    "equation_2_selector": r"$\frac{1}{p}\sum_{k=0}^{p-1}e^{2\pi i k n/p}=\mathbf{1}_{p\mid n},\qquad H_p(n)=\mathbf{1}_{p\nmid n}$",
    "equation_3_mask": r"$M_N(n)=\mathbf{1}_{\{2,\ldots,N\}}(n)\prod_{p\leq\sqrt{N}}H_p(n)+\sum_{p\leq\sqrt{N}}\mathbf{1}_{\{p\}}(n)=\mathbf{1}_{\mathbb{P}\cap[2,N]}(n)$",
    "equation_4_noisy": r"$\widetilde H_p(n)=1-\frac{1}{p}\sum_{k=0}^{p-1}(1+\varepsilon_{p,k})\exp\!\left[i\left(\frac{2\pi nk}{p}(1+\eta_{p,k})+\phi_{p,k}\right)\right]$",
    "equation_5_scaling": r"$A_{\mathrm{WRRA}}(N)=2\pi(\sqrt{N}),\qquad T_{\mathrm{paths}}(N)=\sum_{p\leq\sqrt{N}}(p+1)+\pi(\sqrt{N})$",
}


for name, equation in EQUATIONS.items():
    fig = plt.figure(figsize=(10.5, 0.9))
    fig.patch.set_alpha(0)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_axis_off()
    ax.text(0.5, 0.5, equation, ha="center", va="center", fontsize=19, color="black")
    fig.savefig(OUT / f"{name}.png", dpi=320, transparent=True, bbox_inches="tight", pad_inches=0.08)
    plt.close(fig)
