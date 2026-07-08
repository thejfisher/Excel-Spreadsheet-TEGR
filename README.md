# Excel-Spreadsheet-TEGR

A fully transparent, **formula-only** Excel workbook implementing the **10×10
Jacobian matrix** for relativistic particle dynamics in the
**Teleparallel Equivalent of General Relativity (TEGR)**.

No macros. No hidden code. Every cell is auditable.

---

## File

| File | Description |
|---|---|
| `TEGR_Jacobian.xlsx` | The main spreadsheet (5 worksheets, all formula-based) |
| `generate_tegr_jacobian.py` | Python script that builds the `.xlsx` from scratch using `openpyxl` |

---

## Physical Model

The workbook implements the **linearized TEGR geodesic equations** for a test
particle in the isotropic post-Newtonian metric

```
ds² = −(1+2Φ)dt² + (1−2Φ)(dx²+dy²+dz²),   Φ = −GM/r
```

which in TEGR arises from the diagonal tetrad

```
e^(0) = √(1+2Φ) dt,   e^(i) = √(1−2Φ) dxⁱ   (i = 1,2,3)
```

The torsion scalar along the worldline (leading weak-field term) is

```
T = −4GM/r³
```

### 10-Dimensional Phase-Space State Vector

```
z = ( t,  x,  y,  z,  u^t,  u^x,  u^y,  u^z,  T_scalar,  Ṫ )
     z₁  z₂  z₃  z₄   z₅    z₆    z₇    z₈      z₉       z₁₀
```

where `u^μ = dx^μ/dτ` are the four-velocity components and `Ṫ = dT/dτ`.

### Equations of Motion

```
f₁  = dt/dτ   = u^t
f₂  = dx/dτ   = u^x
f₃  = dy/dτ   = u^y
f₄  = dz/dτ   = u^z
f₅  = du^t/dτ = −2u^t(Φ_x u^x + Φ_y u^y + Φ_z u^z)
f₆  = du^x/dτ = −Φ_x(u^t²−u^x²+u^y²+u^z²) + 2Φ_y u^x u^y + 2Φ_z u^x u^z
f₇  = du^y/dτ = −Φ_y(u^t²+u^x²−u^y²+u^z²) + 2Φ_x u^x u^y + 2Φ_z u^y u^z
f₈  = du^z/dτ = −Φ_z(u^t²+u^x²+u^y²−u^z²) + 2Φ_x u^x u^z + 2Φ_y u^y u^z
f₉  = dT/dτ   = Ṫ
f₁₀ = d²T/dτ² = Σ T_αβ u^α u^β + T_x f₆ + T_y f₇ + T_z f₈
```

The **Jacobian** `J_{ij} = ∂fᵢ/∂zⱼ` is the main 10×10 output, evaluated at
a user-configurable point in phase space.

---

## Default Evaluation Point

| Quantity | Value | Meaning |
|---|---|---|
| (x₀, y₀, z₀) | (1, 0, 0) | On the positive x-axis at r₀ = 1 |
| M | 0.01 | Source mass (M/r₀ = 0.01, weak-field regime) |
| v | √(GM/r₀) ≈ 0.1 | Newtonian circular-orbit speed |
| γ | 1/√(1−v²) ≈ 1.005 | Lorentz factor |
| (u^t, u^x, u^y, u^z) | (γ, 0, γv, 0) | Tangential circular orbit |

---

## Workbook Structure

| Sheet | Contents |
|---|---|
| **Parameters** | Hard-coded inputs (G, c, M, evaluation point) and all derived quantities (r₀, γ, Φ, ∂Φ/∂x, ∂²Φ/∂x², T, ∂T/∂x, …) as formulas |
| **Tetrad** | Diagonal tetrad matrix `e^a_μ` and its inverse at the evaluation point |
| **Connections** | Non-zero Weitzenböck connection coefficients `Γ^ρ_{μν}` (= Christoffel symbols to first order in Φ) |
| **Jacobian** | The **10×10 Jacobian matrix** — every element is an Excel formula referencing the Parameters sheet |
| **Analysis** | Trace, Frobenius norm, tidal tensor components, orbital frequency checks, TEGR coupling strengths |

### Parameters Sheet — Key Cell Map

| Cell | Symbol | Formula |
|---|---|---|
| B3 | G | `1` (geometric units) |
| B5 | M | `0.01` |
| B9 | x₀ | `1` |
| B13 | r₀ | `=SQRT(B9²+B10²+B11²)` |
| B14 | v | `=SQRT(B3*B5/B13)` |
| B15 | γ | `=1/SQRT(1−B14²)` |
| B17 | u^t | `=B15` |
| B19 | u^y | `=B15*B14` |
| B23–B31 | Φ_x … Φ_yz | Second derivatives of Φ |
| B35–B43 | T_x … T_yz | First and second derivatives of T |
| B45–B49 | T_xxx … | Third derivatives (for row 10) |
| B52–B54 | f₆, f₇, f₈ | Evaluated equations of motion |

---

## Physical Validation

Key results at the default evaluation point:

| Property | Value | Check |
|---|---|---|
| Tr(J) | 0 | ✓ Liouville / symplectic structure |
| K_yy (transverse tidal) | −M/r₀³ = −0.01 | ✓ Exact Newtonian tidal |
| K_xx (radial tidal) | +2M(1+v²)γ²/r₀³ | ✓ Centrifugal + tidal |
| J[1,5] = J[2,6] = J[3,7] = J[4,8] = J[9,10] | 1 | ✓ Kinematic identity couplings |
| J[6,5] = J[5,6] | −2γΦ_x | ✓ Gravitoelectric coupling |

---

## How to Regenerate

```bash
pip install openpyxl
python generate_tegr_jacobian.py
```

---

## Units

All quantities are in **geometrized units** with G = c = 1.  
Lengths, times, and masses share the unit [m] (metres of geometric length).

---

## References

- Aldrovandi, R. & Pereira, J.G. *Teleparallel Gravity: An Introduction*. Springer, 2013.
- Maluf, J.W. "The teleparallel equivalent of general relativity." *Ann. Phys.* **525**, 339–357 (2013).
- Misner, C.W., Thorne, K.S. & Wheeler, J.A. *Gravitation*. W. H. Freeman, 1973.
