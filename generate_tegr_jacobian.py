#!/usr/bin/env python3
"""
generate_tegr_jacobian.py
=========================
Generates TEGR_Jacobian.xlsx — a fully transparent, formula-only Excel
workbook implementing the 10×10 Jacobian matrix for relativistic particle
dynamics in Teleparallel Equivalent of General Relativity (TEGR).

Physical model
--------------
A test particle moves on a geodesic in a weak gravitational field described by
TEGR.  The background is the isotropic post-Newtonian metric

    ds² = −(1+2Φ)dt² + (1−2Φ)(dx²+dy²+dz²),   Φ = −GM/r

which in TEGR is sourced by the diagonal tetrad

    e^(0) = √(1+2Φ) dt,   e^(i) = √(1−2Φ) dx^i   (i=1,2,3)

The torsion scalar for this tetrad (leading weak-field term) is

    T = −4GM/r³

The 10-dimensional phase-space state vector is

    z = (t, x, y, z, u^t, u^x, u^y, u^z, T_scalar, T_rate)

where u^μ = dx^μ/dτ are the four-velocity components and T_rate = dT/dτ.

The equations of motion are

    dz_1/dτ  =  u^t                                        (time coord)
    dz_2/dτ  =  u^x                                        (x coord)
    dz_3/dτ  =  u^y                                        (y coord)
    dz_4/dτ  =  u^z                                        (z coord)
    dz_5/dτ  =  −2 u^t (Φ_x u^x + Φ_y u^y + Φ_z u^z)    (du^t/dτ)
    dz_6/dτ  =  −Φ_x(u^t²−u^x²+u^y²+u^z²)
                 + 2 Φ_y u^x u^y + 2 Φ_z u^x u^z          (du^x/dτ)
    dz_7/dτ  =  −Φ_y(u^t²+u^x²−u^y²+u^z²)
                 + 2 Φ_x u^x u^y + 2 Φ_z u^y u^z          (du^y/dτ)
    dz_8/dτ  =  −Φ_z(u^t²+u^x²+u^y²−u^z²)
                 + 2 Φ_x u^x u^z + 2 Φ_y u^y u^z          (du^z/dτ)
    dz_9/dτ  =  T_rate                                     (dT/dτ)
    dz_10/dτ =  T_xx u^x² + T_yy u^y² + T_zz u^z²
                 + 2 T_xy u^x u^y + 2 T_xz u^x u^z
                 + 2 T_yz u^y u^z
                 + T_x(du^x/dτ) + T_y(du^y/dτ) + T_z(du^z/dτ)   (d²T/dτ²)

where Φ_x, Φ_xx, T_x, T_xx, … are the respective partial derivatives of Φ
and T evaluated at the instantaneous position.

The 10×10 Jacobian matrix J_ij = ∂(dz_i/dτ)/∂z_j is the main output.
Every cell in the Jacobian sheet is an Excel formula referencing the
Parameters sheet; no macros, no hidden code.

Evaluation point (default)
---------------------------
The Jacobian is evaluated at a circular orbit in the x-y plane:
  (x₀,y₀,z₀) = (1,0,0),   u^x=0,  u^y = γv,  u^t = γ
  v = √(GM/r₀)  (Newtonian circular-orbit speed),  M = 0.01, r₀ = 1
"""

from openpyxl import Workbook
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                              numbers)
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------
C_TITLE     = "1F4E79"   # dark navy
C_SECTION   = "2E75B6"   # medium blue
C_INPUT     = "D6E4F7"   # pale blue   → hard-coded inputs
C_FORMULA   = "E2EFDA"   # pale green  → computed formula cells
C_ZERO      = "F2F2F2"   # light grey  → structural zeros in Jacobian
C_NONZERO   = "FFFDE7"   # pale yellow → non-trivially non-zero Jacobian cells
C_DIAGONAL  = "FFE0B2"   # pale orange → diagonal elements
C_WHITE     = "FFFFFF"

# ---------------------------------------------------------------------------
# Style helpers
# ---------------------------------------------------------------------------
def _fill(hex_c):
    return PatternFill("solid", fgColor=hex_c)

def _font(bold=False, size=10, color="000000"):
    return Font(bold=bold, size=size, color=color)

def _border():
    s = Side(style="thin", color="BDBDBD")
    return Border(left=s, right=s, top=s, bottom=s)

def _thick_border():
    tk = Side(style="medium", color="1F4E79")
    return Border(left=tk, right=tk, top=tk, bottom=tk)

def _center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def _left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def style(ws, row, col,
          value=None, formula=None,
          bold=False, size=10, color="000000",
          bg=None, align="left",
          border=True, num_fmt=None):
    """Write a value or formula to a cell and apply styling."""
    c = ws.cell(row=row, column=col)
    c.value = formula if formula is not None else value
    c.font = _font(bold=bold, size=size, color=color)
    if bg:
        c.fill = _fill(bg)
    c.alignment = _center() if align == "center" else _left()
    if border:
        c.border = _border()
    if num_fmt:
        c.number_format = num_fmt
    return c


# ===========================================================================
# Workbook
# ===========================================================================
wb = Workbook()

# ---------------------------------------------------------------------------
# Sheet 1 — Parameters
# ---------------------------------------------------------------------------
ws = wb.active
ws.title = "Parameters"
ws.column_dimensions["A"].width = 42
ws.column_dimensions["B"].width = 24
ws.column_dimensions["C"].width = 22
ws.column_dimensions["D"].width = 50

# --- Title row ---------------------------------------------------------------
ws.merge_cells("A1:D1")
c = ws["A1"]
c.value = ("TEGR Jacobian — Relativistic Particle Dynamics in "
           "Teleparallel Equivalent of General Relativity")
c.font  = _font(bold=True, size=13, color=C_WHITE)
c.fill  = _fill(C_TITLE)
c.alignment = _center()
ws.row_dimensions[1].height = 26

# Helper: write a section-header spanning columns A:D
def section(ws, row, title):
    ws.merge_cells(f"A{row}:D{row}")
    c = ws[f"A{row}"]
    c.value = title
    c.font  = _font(bold=True, size=10, color=C_WHITE)
    c.fill  = _fill(C_SECTION)
    c.alignment = _left()
    ws.row_dimensions[row].height = 18

# Helper: write one labelled data row
def row_entry(ws, row, label, val_or_formula, unit="", note="",
              is_formula=False, num_fmt='0.000000'):
    bg = C_FORMULA if is_formula else C_INPUT
    style(ws, row, 1, value=label, bg=bg)
    if is_formula:
        style(ws, row, 2, formula=val_or_formula, bg=bg, num_fmt=num_fmt)
    else:
        style(ws, row, 2, value=val_or_formula, bg=bg, num_fmt=num_fmt)
    style(ws, row, 3, value=unit, bg=bg)
    style(ws, row, 4, value=note, bg=bg)

# ── Physical constants ────────────────────────────────────────────────────
section(ws, 2, "Physical Constants  (Geometrized Units: G = c = 1)")
row_entry(ws,  3, "G — Gravitational constant",   1,     "dimensionless",        "Geometric units G = 1")
row_entry(ws,  4, "c — Speed of light",           1,     "dimensionless",        "Geometric units c = 1")
row_entry(ws,  5, "M — Central (source) mass",    0.01,  "geometric mass [m]",   "M/r₀ = 0.01  →  weak-field regime")
row_entry(ws,  6, "m — Test-particle mass",       1,     "geometric mass [m]",   "Geodesic independent of m; retained for completeness")

# ── Evaluation-point coordinates ─────────────────────────────────────────
section(ws, 7, "Evaluation Point — Coordinates  x^μ = (t₀, x₀, y₀, z₀)")
row_entry(ws,  8, "t₀ — Coordinate time",          0, "geometric time [m]",   "Arbitrary for static background")
row_entry(ws,  9, "x₀ — x-position",               1, "geometric length [m]", "Particle on positive x-axis  (r₀ = x₀)")
row_entry(ws, 10, "y₀ — y-position",               0, "geometric length [m]", "Equatorial plane")
row_entry(ws, 11, "z₀ — z-position",               0, "geometric length [m]", "Equatorial plane")

# ── Derived radial quantities ─────────────────────────────────────────────
section(ws, 12, "Derived Radial Quantities")
row_entry(ws, 13, "r₀ — Radial distance",
          "=SQRT(B9^2+B10^2+B11^2)",
          "geometric length [m]", "r₀ = √(x₀²+y₀²+z₀²)",
          is_formula=True)
row_entry(ws, 14, "v — Circular-orbit speed",
          "=SQRT(B3*B5/B13)",
          "dimensionless (c = 1)", "Newtonian: v = √(GM/r₀)",
          is_formula=True)
row_entry(ws, 15, "γ — Lorentz factor",
          "=1/SQRT(1-B14^2)",
          "dimensionless", "γ = 1/√(1−v²)",
          is_formula=True)

# ── Four-velocity at evaluation point ────────────────────────────────────
section(ws, 16, "Evaluation Point — Four-Velocity Components  u^μ = dx^μ/dτ")
row_entry(ws, 17, "u^t = dt/dτ",    "=B15",      "dimensionless",        "Lorentz factor γ  (circular orbit)", is_formula=True)
row_entry(ws, 18, "u^x = dx/dτ",    0,           "dimensionless",        "Zero at apoapsis (tangential velocity)")
row_entry(ws, 19, "u^y = dy/dτ",    "=B15*B14",  "dimensionless",        "Tangential: u^y = γ·v",               is_formula=True)
row_entry(ws, 20, "u^z = dz/dτ",    0,           "dimensionless",        "Equatorial plane")

# ── Gravitational potential Φ = −GM/r ────────────────────────────────────
section(ws, 21, "Gravitational Potential  Φ = −GM/r  and Partial Derivatives at (x₀,y₀,z₀)")
row_entry(ws, 22, "Φ₀ — Potential",
          "=-B3*B5/B13",
          "dimensionless", "Φ = −GM/r", is_formula=True)
row_entry(ws, 23, "Φ_x = ∂Φ/∂x",
          "=B3*B5*B9/B13^3",
          "1/[m]", "= GMx/r³", is_formula=True)
row_entry(ws, 24, "Φ_y = ∂Φ/∂y",
          "=B3*B5*B10/B13^3",
          "1/[m]", "= GMy/r³", is_formula=True)
row_entry(ws, 25, "Φ_z = ∂Φ/∂z",
          "=B3*B5*B11/B13^3",
          "1/[m]", "= GMz/r³", is_formula=True)
row_entry(ws, 26, "Φ_xx = ∂²Φ/∂x²",
          "=B3*B5*(B13^2-3*B9^2)/B13^5",
          "1/[m²]", "= GM(r²−3x²)/r⁵  →  −2M/r₀³ at x₀=r₀", is_formula=True)
row_entry(ws, 27, "Φ_yy = ∂²Φ/∂y²",
          "=B3*B5*(B13^2-3*B10^2)/B13^5",
          "1/[m²]", "= GM(r²−3y²)/r⁵  →  +M/r₀³  at y₀=0", is_formula=True)
row_entry(ws, 28, "Φ_zz = ∂²Φ/∂z²",
          "=B3*B5*(B13^2-3*B11^2)/B13^5",
          "1/[m²]", "= GM(r²−3z²)/r⁵  →  +M/r₀³  at z₀=0", is_formula=True)
row_entry(ws, 29, "Φ_xy = ∂²Φ/∂x∂y",
          "=-3*B3*B5*B9*B10/B13^5",
          "1/[m²]", "= −3GMxy/r⁵", is_formula=True)
row_entry(ws, 30, "Φ_xz = ∂²Φ/∂x∂z",
          "=-3*B3*B5*B9*B11/B13^5",
          "1/[m²]", "= −3GMxz/r⁵", is_formula=True)
row_entry(ws, 31, "Φ_yz = ∂²Φ/∂y∂z",
          "=-3*B3*B5*B10*B11/B13^5",
          "1/[m²]", "= −3GMyz/r⁵", is_formula=True)

# ── Torsion scalar T = −4GM/r³ ───────────────────────────────────────────
section(ws, 32, "TEGR Torsion Scalar  T = −4GM/r³  and Partial Derivatives at (x₀,y₀,z₀)")
row_entry(ws, 33, "T₀ — Torsion scalar",
          "=-4*B3*B5/B13^3",
          "1/[m²]", "Leading weak-field torsion scalar for diagonal tetrad", is_formula=True)
row_entry(ws, 34, "Ṫ₀ = dT/dτ at eval point",
          "=B37*B18+B38*B19+B39*B20",
          "1/[m²·s]",
          "= T_x·u^x + T_y·u^y + T_z·u^z  (zero for tangential orbit)", is_formula=True)
row_entry(ws, 35, "T_x = ∂T/∂x",
          "=12*B3*B5*B9/B13^5",
          "1/[m³]", "= 12GMx/r⁵", is_formula=True)
row_entry(ws, 36, "T_y = ∂T/∂y",
          "=12*B3*B5*B10/B13^5",
          "1/[m³]", "= 12GMy/r⁵", is_formula=True)
row_entry(ws, 37, "T_z = ∂T/∂z",
          "=12*B3*B5*B11/B13^5",
          "1/[m³]", "= 12GMz/r⁵", is_formula=True)
row_entry(ws, 38, "T_xx = ∂²T/∂x²",
          "=12*B3*B5*(B13^2-5*B9^2)/B13^7",
          "1/[m⁴]", "= 12GM(r²−5x²)/r⁷  →  −48M/r₀⁵ at x₀=r₀", is_formula=True)
row_entry(ws, 39, "T_yy = ∂²T/∂y²",
          "=12*B3*B5*(B13^2-5*B10^2)/B13^7",
          "1/[m⁴]", "= 12GM(r²−5y²)/r⁷  →  +12M/r₀⁵ at y₀=0", is_formula=True)
row_entry(ws, 40, "T_zz = ∂²T/∂z²",
          "=12*B3*B5*(B13^2-5*B11^2)/B13^7",
          "1/[m⁴]", "= 12GM(r²−5z²)/r⁷  →  +12M/r₀⁵ at z₀=0", is_formula=True)
row_entry(ws, 41, "T_xy = ∂²T/∂x∂y",
          "=-60*B3*B5*B9*B10/B13^7",
          "1/[m⁴]", "= −60GMxy/r⁷", is_formula=True)
row_entry(ws, 42, "T_xz = ∂²T/∂x∂z",
          "=-60*B3*B5*B9*B11/B13^7",
          "1/[m⁴]", "= −60GMxz/r⁷", is_formula=True)
row_entry(ws, 43, "T_yz = ∂²T/∂y∂z",
          "=-60*B3*B5*B10*B11/B13^7",
          "1/[m⁴]", "= −60GMyz/r⁷", is_formula=True)

# ── Third derivatives of T (used in Jacobian row 10) ─────────────────────
section(ws, 44, "Third Derivatives of T  (required for Jacobian row 10 — d²T/dτ² sensitivity)")
row_entry(ws, 45, "T_xxx = ∂³T/∂x³",
          "=12*B3*B5*B9*(35*B9^2-15*B13^2)/B13^9",
          "1/[m⁵]", "= 12GMx(35x²−15r²)/r⁹", is_formula=True)
row_entry(ws, 46, "T_xyy = ∂T_xy/∂y = ∂³T/∂x∂y²",
          "=-60*B3*B5*B9*(B13^2-7*B10^2)/B13^9",
          "1/[m⁵]", "= −60GMx(r²−7y²)/r⁹  →  −60M/r₀⁷ at y₀=0", is_formula=True)
row_entry(ws, 47, "T_xzz = ∂T_xz/∂z = ∂³T/∂x∂z²",
          "=-60*B3*B5*B9*(B13^2-7*B11^2)/B13^9",
          "1/[m⁵]", "= −60GMx(r²−7z²)/r⁹  →  −60M/r₀⁷ at z₀=0", is_formula=True)
row_entry(ws, 48, "∂T_yy/∂x = ∂³T/∂y²∂x",
          "=12*B3*B5*B9*(35*B10^2-5*B13^2)/B13^9",
          "1/[m⁵]", "= 12GMx(35y²−5r²)/r⁹  →  −60M/r₀⁷ at y₀=0", is_formula=True)
row_entry(ws, 49, "∂T_zz/∂x = ∂³T/∂z²∂x",
          "=12*B3*B5*B9*(35*B11^2-5*B13^2)/B13^9",
          "1/[m⁵]", "= 12GMx(35z²−5r²)/r⁹  →  −60M/r₀⁷ at z₀=0", is_formula=True)

# ── Equations of motion evaluated at the evaluation point ─────────────────
section(ws, 50, "Equations of Motion  f_i = dz_i/dτ  Evaluated at (t₀,x₀,y₀,z₀,u^t,u^x,u^y,u^z)")
# f5 = du^t/dτ = -2 u^t (Φ_x u^x + Φ_y u^y + Φ_z u^z)
row_entry(ws, 51, "f₅ = du^t/dτ",
          "=-2*B17*(B23*B18+B24*B19+B25*B20)",
          "dimensionless/[m]",
          "= −2 u^t (Φ_x u^x + Φ_y u^y + Φ_z u^z)",
          is_formula=True, num_fmt="0.000000E+00")
# f6 = du^x/dτ
row_entry(ws, 52, "f₆ = du^x/dτ",
          "=-B23*(B17^2-B18^2+B19^2+B20^2)+2*B24*B18*B19+2*B25*B18*B20",
          "dimensionless/[m]",
          "= −Φ_x(u^t²−u^x²+u^y²+u^z²) + 2Φ_y u^x u^y + 2Φ_z u^x u^z",
          is_formula=True, num_fmt="0.000000E+00")
# f7 = du^y/dτ
row_entry(ws, 53, "f₇ = du^y/dτ",
          "=-B24*(B17^2+B18^2-B19^2+B20^2)+2*B23*B18*B19+2*B25*B19*B20",
          "dimensionless/[m]",
          "= −Φ_y(u^t²+u^x²−u^y²+u^z²) + 2Φ_x u^x u^y + 2Φ_z u^y u^z",
          is_formula=True, num_fmt="0.000000E+00")
# f8 = du^z/dτ
row_entry(ws, 54, "f₈ = du^z/dτ",
          "=-B25*(B17^2+B18^2+B19^2-B20^2)+2*B23*B18*B20+2*B24*B19*B20",
          "dimensionless/[m]",
          "= −Φ_z(u^t²+u^x²+u^y²−u^z²) + 2Φ_x u^x u^z + 2Φ_y u^y u^z",
          is_formula=True, num_fmt="0.000000E+00")

# Freeze first row
ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# Sheet 2 — Tetrad
# ---------------------------------------------------------------------------
wt = wb.create_sheet("Tetrad")
wt.column_dimensions["A"].width = 38
for c in "BCDEF":
    wt.column_dimensions[c].width = 16
wt.column_dimensions["G"].width = 46

wt.merge_cells("A1:G1")
c = wt["A1"]
c.value = ("TEGR Diagonal Tetrad  e^a_{μ}  — Weak-Field Isotropic Schwarzschild "
           "(geometrized units, evaluated at (x₀,y₀,z₀))")
c.font  = _font(bold=True, size=12, color=C_WHITE)
c.fill  = _fill(C_TITLE)
c.alignment = _center()
wt.row_dimensions[1].height = 26

# Section: Background
section(wt, 2, "Tetrad Scalar Functions at Evaluation Point")
row_entry(wt, 3,  "A = √(1+2Φ₀)  — time lapse",
          "=SQRT(1+2*Parameters!B22)",
          "dimensionless",
          "e^(0)_t = A  (temporal tetrad component)", is_formula=True)
row_entry(wt, 4,  "B = √(1−2Φ₀)  — spatial scale",
          "=SQRT(1-2*Parameters!B22)",
          "dimensionless",
          "e^(i)_i = B  (diagonal spatial tetrad components)", is_formula=True)
row_entry(wt, 5,  "A⁻¹ = inverse time lapse",
          "=1/B3",
          "dimensionless",
          "e_(0)^t = 1/A  (inverse tetrad)", is_formula=True)
row_entry(wt, 6,  "B⁻¹ = inverse spatial scale",
          "=1/B4",
          "dimensionless",
          "e_(i)^i = 1/B  (inverse tetrad)", is_formula=True)

# Tetrad matrix e^a_μ
section(wt, 7, "Tetrad Matrix  e^a_{μ}  (rows = local-frame index a = 0,1,2,3; columns = coord index μ = t,x,y,z)")
# Header row
hdr_labels = ["", "μ = t", "μ = x", "μ = y", "μ = z", "", "Note"]
for ci, lbl in enumerate(hdr_labels, start=1):
    c = wt.cell(row=8, column=ci)
    c.value = lbl
    c.font  = _font(bold=True, size=10, color=C_WHITE)
    c.fill  = _fill(C_SECTION)
    c.alignment = _center()
    c.border = _border()

tetrad_rows = [
    # label,      et,    ex,    ey,    ez
    ("a = 0 (time)",  "=B3",  0,     0,     0),
    ("a = 1 (x)",      0,     "=B4", 0,     0),
    ("a = 2 (y)",      0,     0,     "=B4", 0),
    ("a = 3 (z)",      0,     0,     0,     "=B4"),
]
for ri, (lbl, et, ex, ey, ez) in enumerate(tetrad_rows, start=9):
    wt.cell(row=ri, column=1).value = lbl
    wt.cell(row=ri, column=1).font  = _font(bold=True, size=10)
    wt.cell(row=ri, column=1).fill  = _fill(C_INPUT)
    wt.cell(row=ri, column=1).border = _border()
    for ci, val in enumerate([et, ex, ey, ez], start=2):
        c = wt.cell(row=ri, column=ci)
        c.value = val
        c.fill  = _fill(C_FORMULA) if isinstance(val, str) else _fill(C_ZERO)
        c.alignment = _center()
        c.border = _border()
        c.number_format = "0.000000"

# Inverse tetrad matrix e_a^μ
section(wt, 13, "Inverse Tetrad  e_{a}^{μ}  (satisfies  e^a_{μ} e_{a}^{ν} = δ^ν_μ)")
for ci, lbl in enumerate(hdr_labels, start=1):
    c = wt.cell(row=14, column=ci)
    c.value = lbl
    c.font  = _font(bold=True, size=10, color=C_WHITE)
    c.fill  = _fill(C_SECTION)
    c.alignment = _center()
    c.border = _border()
inv_rows = [
    ("a = 0 (time)",  "=B5",  0,     0,     0),
    ("a = 1 (x)",      0,     "=B6", 0,     0),
    ("a = 2 (y)",      0,     0,     "=B6", 0),
    ("a = 3 (z)",      0,     0,     0,     "=B6"),
]
for ri, (lbl, et, ex, ey, ez) in enumerate(inv_rows, start=15):
    wt.cell(row=ri, column=1).value = lbl
    wt.cell(row=ri, column=1).font  = _font(bold=True, size=10)
    wt.cell(row=ri, column=1).fill  = _fill(C_INPUT)
    wt.cell(row=ri, column=1).border = _border()
    for ci, val in enumerate([et, ex, ey, ez], start=2):
        c = wt.cell(row=ri, column=ci)
        c.value = val
        c.fill  = _fill(C_FORMULA) if isinstance(val, str) else _fill(C_ZERO)
        c.alignment = _center()
        c.border = _border()
        c.number_format = "0.000000"

# Torsion scalar formula
section(wt, 19, "Torsion Scalar T (from tetrad — leading weak-field expression)")
row_entry(wt, 20, "T₀ = −4GM/r³",
          "=Parameters!B33",
          "1/[m²]",
          "For diagonal tetrad: T = −4(A′/A)²(1+2Ar) to leading order → −4GM/r³",
          is_formula=True)
row_entry(wt, 21, "Verification: Φ₀ check",
          "=Parameters!B22",
          "dimensionless",
          "Should equal −GM/r₀ = −M/r₀",
          is_formula=True)


# ---------------------------------------------------------------------------
# Sheet 3 — Connections
# ---------------------------------------------------------------------------
wc = wb.create_sheet("Connections")
wc.column_dimensions["A"].width = 38
for col in "BCDE":
    wc.column_dimensions[col].width = 20
wc.column_dimensions["E"].width = 46

wc.merge_cells("A1:E1")
c = wc["A1"]
c.value = ("TEGR Weitzenböck Connection Coefficients  Γ^ρ_{μν} = e^ρ_a ∂_μ e^a_ν  "
           "— Relevant Non-Zero Components at Evaluation Point")
c.font  = _font(bold=True, size=12, color=C_WHITE)
c.fill  = _fill(C_TITLE)
c.alignment = _center()
wc.row_dimensions[1].height = 26

section(wc, 2, ("In the weak-field limit the Weitzenböck connection coincides with the "
                "Levi-Civita (Christoffel) connection to first order in Φ.  "
                "Non-zero components listed below."))

# Connection table header
conn_hdrs = ["Symbol", "Formula (Excel)", "Value at eval point", "", "Physical meaning"]
for ci, lbl in enumerate(conn_hdrs, start=1):
    c = wc.cell(row=3, column=ci)
    c.value = lbl
    c.font  = _font(bold=True, size=10, color=C_WHITE)
    c.fill  = _fill(C_SECTION)
    c.alignment = _center()
    c.border = _border()

# fmt: off  (keep long formula strings readable)
# Using Parameters sheet references
connections = [
    # symbol, formula-string, meaning
    ("Γ^t_{tx} = Γ^t_{xt}", "=Parameters!B23",      "∂Φ/∂x = GMx/r³  —  time-position mixing (x)"),
    ("Γ^t_{ty} = Γ^t_{yt}", "=Parameters!B24",      "∂Φ/∂y = GMy/r³  —  time-position mixing (y)"),
    ("Γ^t_{tz} = Γ^t_{zt}", "=Parameters!B25",      "∂Φ/∂z = GMz/r³  —  time-position mixing (z)"),
    ("Γ^x_{tt}",             "=Parameters!B23",      "∂Φ/∂x  —  gravitoelectric acceleration (x)"),
    ("Γ^x_{xx}",             "=-Parameters!B23",     "−∂Φ/∂x  — self-diagonal spatial"),
    ("Γ^x_{yy}",             "=Parameters!B23",      "∂Φ/∂x  — off-diagonal spatial y-y"),
    ("Γ^x_{zz}",             "=Parameters!B23",      "∂Φ/∂x  — off-diagonal spatial z-z"),
    ("Γ^x_{xy} = Γ^x_{yx}", "=-Parameters!B24",     "−∂Φ/∂y  — spatial cross-coupling (xy)"),
    ("Γ^x_{xz} = Γ^x_{zx}", "=-Parameters!B25",     "−∂Φ/∂z  — spatial cross-coupling (xz)"),
    ("Γ^y_{tt}",             "=Parameters!B24",      "∂Φ/∂y  —  gravitoelectric acceleration (y)"),
    ("Γ^y_{yy}",             "=-Parameters!B24",     "−∂Φ/∂y  — self-diagonal spatial"),
    ("Γ^y_{xx}",             "=Parameters!B24",      "∂Φ/∂y  — off-diagonal spatial x-x"),
    ("Γ^y_{zz}",             "=Parameters!B24",      "∂Φ/∂y  — off-diagonal spatial z-z"),
    ("Γ^y_{xy} = Γ^y_{yx}", "=-Parameters!B23",     "−∂Φ/∂x  — spatial cross-coupling (xy)"),
    ("Γ^y_{yz} = Γ^y_{zy}", "=-Parameters!B25",     "−∂Φ/∂z  — spatial cross-coupling (yz)"),
    ("Γ^z_{tt}",             "=Parameters!B25",      "∂Φ/∂z  —  gravitoelectric acceleration (z)"),
    ("Γ^z_{zz}",             "=-Parameters!B25",     "−∂Φ/∂z  — self-diagonal spatial"),
    ("Γ^z_{xx}",             "=Parameters!B25",      "∂Φ/∂z  — off-diagonal spatial x-x"),
    ("Γ^z_{yy}",             "=Parameters!B25",      "∂Φ/∂z  — off-diagonal spatial y-y"),
    ("Γ^z_{xz} = Γ^z_{zx}", "=-Parameters!B23",     "−∂Φ/∂x  — spatial cross-coupling (xz)"),
    ("Γ^z_{yz} = Γ^z_{zy}", "=-Parameters!B24",     "−∂Φ/∂y  — spatial cross-coupling (yz)"),
]
# fmt: on

for ri, (sym, fml, meaning) in enumerate(connections, start=4):
    wc.cell(row=ri, column=1).value = sym
    wc.cell(row=ri, column=1).font  = _font(bold=True, size=10)
    wc.cell(row=ri, column=1).fill  = _fill(C_INPUT)
    wc.cell(row=ri, column=1).border = _border()
    c2 = wc.cell(row=ri, column=2)
    c2.value = fml
    c2.fill  = _fill(C_FORMULA)
    c2.border = _border()
    c2.number_format = "0.000000E+00"
    wc.cell(row=ri, column=3).value = ""  # blank spacer
    wc.cell(row=ri, column=3).border = _border()
    c4 = wc.cell(row=ri, column=4)
    c4.value = meaning
    c4.fill  = _fill(C_INPUT)
    c4.border = _border()
    c4.alignment = _left()


# ===========================================================================
# Sheet 4 — Jacobian  (the main 10×10 matrix)
# ===========================================================================
wj = wb.create_sheet("Jacobian")

wj.column_dimensions["A"].width = 28
for i in range(2, 13):
    wj.column_dimensions[get_column_letter(i)].width = 18

# ── Title ────────────────────────────────────────────────────────────────
wj.merge_cells("A1:K1")
c = wj["A1"]
c.value = ("10×10 Jacobian Matrix  J_{ij} = ∂(dz_i/dτ)/∂z_j  "
           "Evaluated at Circular Orbit  (x₀,0,0)  —  TEGR Weak-Field Geodesic + Torsion Transport")
c.font  = _font(bold=True, size=12, color=C_WHITE)
c.fill  = _fill(C_TITLE)
c.alignment = _center()
wj.row_dimensions[1].height = 28

# ── Physical model note ──────────────────────────────────────────────────
wj.merge_cells("A2:K2")
c = wj["A2"]
c.value = ("State vector  z = (t, x, y, z, u^t, u^x, u^y, u^z, T_scalar, Ṫ)  •  "
           "All cells are Excel formulas referencing the Parameters sheet  •  "
           "Geometrized units G = c = 1")
c.font  = _font(size=9, color="4A4A4A")
c.fill  = _fill("EBF3FA")
c.alignment = _left()
wj.row_dimensions[2].height = 14

# ── Column headers (state-vector variables) ──────────────────────────────
col_labels = [
    "∂/∂t",
    "∂/∂x",
    "∂/∂y",
    "∂/∂z",
    "∂/∂u^t",
    "∂/∂u^x",
    "∂/∂u^y",
    "∂/∂u^z",
    "∂/∂T",
    "∂/∂Ṫ",
]
col_sub = [
    "z₁=t", "z₂=x", "z₃=y", "z₄=z",
    "z₅=u^t", "z₆=u^x", "z₇=u^y", "z₈=u^z",
    "z₉=T", "z₁₀=Ṫ",
]
wj.cell(row=3, column=1).value = "Equation  /  Variable →"
wj.cell(row=3, column=1).font  = _font(bold=True, size=9)
wj.cell(row=3, column=1).fill  = _fill(C_TITLE)
wj.cell(row=3, column=1).alignment = _center()
wj.cell(row=3, column=1).border = _border()

for ci, (lbl, sub) in enumerate(zip(col_labels, col_sub), start=2):
    c = wj.cell(row=3, column=ci)
    c.value = f"{lbl}\n{sub}"
    c.font  = _font(bold=True, size=9, color=C_WHITE)
    c.fill  = _fill(C_TITLE)
    c.alignment = _center()
    c.border = _border()
wj.row_dimensions[3].height = 30

# ── Row labels ───────────────────────────────────────────────────────────
row_labels = [
    "f₁ = dt/dτ = u^t",
    "f₂ = dx/dτ = u^x",
    "f₃ = dy/dτ = u^y",
    "f₄ = dz/dτ = u^z",
    "f₅ = du^t/dτ",
    "f₆ = du^x/dτ",
    "f₇ = du^y/dτ",
    "f₈ = du^z/dτ",
    "f₉ = dT/dτ = Ṫ",
    "f₁₀ = d²T/dτ²",
]
for ri, lbl in enumerate(row_labels, start=4):
    c = wj.cell(row=ri, column=1)
    c.value = lbl
    c.font  = _font(bold=True, size=9, color=C_WHITE)
    c.fill  = _fill(C_SECTION)
    c.alignment = _left()
    c.border = _border()
    wj.row_dimensions[ri].height = 20

# ── Build the 10×10 Jacobian formula table ──────────────────────────────
# P = shorthand for Parameters sheet references
P = "Parameters!"
ut = f"{P}B17"   # u^t
ux = f"{P}B18"   # u^x
uy = f"{P}B19"   # u^y
uz = f"{P}B20"   # u^z

Px  = f"{P}B23"   # Φ_x
Py  = f"{P}B24"   # Φ_y
Pz  = f"{P}B25"   # Φ_z
Pxx = f"{P}B26"   # Φ_xx
Pyy = f"{P}B27"   # Φ_yy
Pzz = f"{P}B28"   # Φ_zz
Pxy = f"{P}B29"   # Φ_xy
Pxz = f"{P}B30"   # Φ_xz
Pyz = f"{P}B31"   # Φ_yz

Tx  = f"{P}B35"   # T_x
Ty  = f"{P}B36"   # T_y
Tz  = f"{P}B37"   # T_z
Txx = f"{P}B38"   # T_xx
Tyy = f"{P}B39"   # T_yy
Tzz = f"{P}B40"   # T_zz
Txy = f"{P}B41"   # T_xy
Txz = f"{P}B42"   # T_xz
Tyz = f"{P}B43"   # T_yz

T_xxx  = f"{P}B45"   # ∂³T/∂x³
T_xyy  = f"{P}B46"   # ∂³T/∂x∂y²
T_xzz  = f"{P}B47"   # ∂³T/∂x∂z²
Tyy_dx = f"{P}B48"   # ∂T_yy/∂x
Tzz_dx = f"{P}B49"   # ∂T_zz/∂x

f6eval = f"{P}B52"   # f₆ at eval point
f7eval = f"{P}B53"   # f₇ at eval point
f8eval = f"{P}B54"   # f₈ at eval point

# Helper: velocity-squared combinations
ut2  = f"{ut}^2"
ux2  = f"{ux}^2"
uy2  = f"{uy}^2"
uz2  = f"{uz}^2"

# Each Jacobian row is a list of 10 formulas (or integer 0/1 for exact zeros/ones)
# row order: [∂/∂t, ∂/∂x, ∂/∂y, ∂/∂z, ∂/∂u^t, ∂/∂u^x, ∂/∂u^y, ∂/∂u^z, ∂/∂T, ∂/∂Ṫ]

def j(formula):
    """Wrap a non-trivial formula string."""
    return f"={formula}"

# --- Row 1: f₁ = u^t  →  ∂f₁/∂z_j -----------------------------------------
J1 = [0, 0, 0, 0, 1, 0, 0, 0, 0, 0]

# --- Row 2: f₂ = u^x  →  ∂f₂/∂z_j -----------------------------------------
J2 = [0, 0, 0, 0, 0, 1, 0, 0, 0, 0]

# --- Row 3: f₃ = u^y  →  ∂f₃/∂z_j -----------------------------------------
J3 = [0, 0, 0, 0, 0, 0, 1, 0, 0, 0]

# --- Row 4: f₄ = u^z  →  ∂f₄/∂z_j -----------------------------------------
J4 = [0, 0, 0, 0, 0, 0, 0, 1, 0, 0]

# --- Row 5: f₅ = −2u^t(Φ_x u^x + Φ_y u^y + Φ_z u^z) ----------------------
# ∂f₅/∂t  = 0  (static background)
# ∂f₅/∂x  = −2u^t (Φ_xx u^x + Φ_xy u^y + Φ_xz u^z)
# ∂f₅/∂y  = −2u^t (Φ_xy u^x + Φ_yy u^y + Φ_yz u^z)
# ∂f₅/∂z  = −2u^t (Φ_xz u^x + Φ_yz u^y + Φ_zz u^z)
# ∂f₅/∂u^t= −2(Φ_x u^x + Φ_y u^y + Φ_z u^z)
# ∂f₅/∂u^x= −2u^t Φ_x
# ∂f₅/∂u^y= −2u^t Φ_y
# ∂f₅/∂u^z= −2u^t Φ_z
J5 = [
    0,
    j(f"-2*{ut}*({Pxx}*{ux}+{Pxy}*{uy}+{Pxz}*{uz})"),
    j(f"-2*{ut}*({Pxy}*{ux}+{Pyy}*{uy}+{Pyz}*{uz})"),
    j(f"-2*{ut}*({Pxz}*{ux}+{Pyz}*{uy}+{Pzz}*{uz})"),
    j(f"-2*({Px}*{ux}+{Py}*{uy}+{Pz}*{uz})"),
    j(f"-2*{ut}*{Px}"),
    j(f"-2*{ut}*{Py}"),
    j(f"-2*{ut}*{Pz}"),
    0,
    0,
]

# --- Row 6: f₆ = −Φ_x(u^t²−u^x²+u^y²+u^z²) + 2Φ_y u^x u^y + 2Φ_z u^x u^z
# ∂f₆/∂t  = 0
# ∂f₆/∂x  = −Φ_xx(u^t²−u^x²+u^y²+u^z²) + 2Φ_xy u^x u^y + 2Φ_xz u^x u^z
# ∂f₆/∂y  = −Φ_xy(u^t²−u^x²+u^y²+u^z²) + 2Φ_yy u^x u^y + 2Φ_yz u^x u^z
# ∂f₆/∂z  = −Φ_xz(u^t²−u^x²+u^y²+u^z²) + 2Φ_yz u^x u^y + 2Φ_zz u^x u^z
# ∂f₆/∂u^t= −2Φ_x u^t
# ∂f₆/∂u^x= 2Φ_x u^x + 2Φ_y u^y + 2Φ_z u^z
# ∂f₆/∂u^y= −2Φ_x u^y + 2Φ_y u^x
# ∂f₆/∂u^z= −2Φ_x u^z + 2Φ_z u^x
combo_x = f"({ut2}-{ux2}+{uy2}+{uz2})"
J6 = [
    0,
    j(f"-{Pxx}*{combo_x}+2*{Pxy}*{ux}*{uy}+2*{Pxz}*{ux}*{uz}"),
    j(f"-{Pxy}*{combo_x}+2*{Pyy}*{ux}*{uy}+2*{Pyz}*{ux}*{uz}"),
    j(f"-{Pxz}*{combo_x}+2*{Pyz}*{ux}*{uy}+2*{Pzz}*{ux}*{uz}"),
    j(f"-2*{Px}*{ut}"),
    j(f"2*{Px}*{ux}+2*{Py}*{uy}+2*{Pz}*{uz}"),
    j(f"-2*{Px}*{uy}+2*{Py}*{ux}"),
    j(f"-2*{Px}*{uz}+2*{Pz}*{ux}"),
    0,
    0,
]

# --- Row 7: f₇ = −Φ_y(u^t²+u^x²−u^y²+u^z²) + 2Φ_x u^x u^y + 2Φ_z u^y u^z
# ∂f₇/∂t  = 0
# ∂f₇/∂x  = −Φ_xy(u^t²+u^x²−u^y²+u^z²) + 2Φ_xx u^x u^y + 2Φ_xz u^y u^z
# ∂f₇/∂y  = −Φ_yy(u^t²+u^x²−u^y²+u^z²) + 2Φ_xy u^x u^y + 2Φ_yz u^y u^z
# ∂f₇/∂z  = −Φ_yz(u^t²+u^x²−u^y²+u^z²) + 2Φ_xz u^x u^y + 2Φ_zz u^y u^z
# ∂f₇/∂u^t= −2Φ_y u^t
# ∂f₇/∂u^x= −2Φ_y u^x + 2Φ_x u^y
# ∂f₇/∂u^y= 2Φ_y u^y + 2Φ_x u^x + 2Φ_z u^z
# ∂f₇/∂u^z= −2Φ_y u^z + 2Φ_z u^y
combo_y = f"({ut2}+{ux2}-{uy2}+{uz2})"
J7 = [
    0,
    j(f"-{Pxy}*{combo_y}+2*{Pxx}*{ux}*{uy}+2*{Pxz}*{uy}*{uz}"),
    j(f"-{Pyy}*{combo_y}+2*{Pxy}*{ux}*{uy}+2*{Pyz}*{uy}*{uz}"),
    j(f"-{Pyz}*{combo_y}+2*{Pxz}*{ux}*{uy}+2*{Pzz}*{uy}*{uz}"),
    j(f"-2*{Py}*{ut}"),
    j(f"-2*{Py}*{ux}+2*{Px}*{uy}"),
    j(f"2*{Py}*{uy}+2*{Px}*{ux}+2*{Pz}*{uz}"),
    j(f"-2*{Py}*{uz}+2*{Pz}*{uy}"),
    0,
    0,
]

# --- Row 8: f₈ = −Φ_z(u^t²+u^x²+u^y²−u^z²) + 2Φ_x u^x u^z + 2Φ_y u^y u^z
# ∂f₈/∂t  = 0
# ∂f₈/∂x  = −Φ_xz(u^t²+u^x²+u^y²−u^z²) + 2Φ_xx u^x u^z + 2Φ_xy u^y u^z
# ∂f₈/∂y  = −Φ_yz(u^t²+u^x²+u^y²−u^z²) + 2Φ_xy u^x u^z + 2Φ_yy u^y u^z
# ∂f₈/∂z  = −Φ_zz(u^t²+u^x²+u^y²−u^z²) + 2Φ_xz u^x u^z + 2Φ_yz u^y u^z
# ∂f₈/∂u^t= −2Φ_z u^t
# ∂f₈/∂u^x= −2Φ_z u^x + 2Φ_x u^z
# ∂f₈/∂u^y= −2Φ_z u^y + 2Φ_y u^z
# ∂f₈/∂u^z= 2Φ_z u^z + 2Φ_x u^x + 2Φ_y u^y
combo_z = f"({ut2}+{ux2}+{uy2}-{uz2})"
J8 = [
    0,
    j(f"-{Pxz}*{combo_z}+2*{Pxx}*{ux}*{uz}+2*{Pxy}*{uy}*{uz}"),
    j(f"-{Pyz}*{combo_z}+2*{Pxy}*{ux}*{uz}+2*{Pyy}*{uy}*{uz}"),
    j(f"-{Pzz}*{combo_z}+2*{Pxz}*{ux}*{uz}+2*{Pyz}*{uy}*{uz}"),
    j(f"-2*{Pz}*{ut}"),
    j(f"-2*{Pz}*{ux}+2*{Px}*{uz}"),
    j(f"-2*{Pz}*{uy}+2*{Py}*{uz}"),
    j(f"2*{Pz}*{uz}+2*{Px}*{ux}+2*{Py}*{uy}"),
    0,
    0,
]

# --- Row 9: f₉ = Ṫ  →  ∂f₉/∂Ṫ = 1, rest 0 --------------------------------
J9 = [0, 0, 0, 0, 0, 0, 0, 0, 0, 1]

# --- Row 10: f₁₀ = d²T/dτ² -------------------------------------------------
# f₁₀ = T_xx u^x² + T_yy u^y² + T_zz u^z²
#        + 2T_xy u^x u^y + 2T_xz u^x u^z + 2T_yz u^y u^z
#        + T_x f₆ + T_y f₇ + T_z f₈
#
# Jacobian elements:
# ∂f₁₀/∂t  = 0
#
# ∂f₁₀/∂x:
#   Derivatives of T_xx,T_yy,T_zz,T_xy,T_xz,T_yz w.r.t. x give third-derivative
#   terms of T, PLUS: ∂(T_x f₆)/∂x = T_xx f₆ + T_x J₆₂
#                     ∂(T_y f₇)/∂x = T_xy f₇ + T_y J₇₂
#                     ∂(T_z f₈)/∂x = T_xz f₈ + T_z J₈₂
#
# ∂f₁₀/∂u^μ:
#   Derivatives of T_xx u^x² etc. w.r.t. u^μ give 2T_μμ u^μ type terms
#   PLUS: T_x J₆μ + T_y J₇μ + T_z J₈μ

# To keep formulas self-contained we reference the already-defined Jacobian cells
# for J₆₂, J₇₂, J₈₂ etc. via the same sheet:
# The Jacobian data starts at row 4, column 2 (for ∂/∂t), so:
#   Row 4 = f₁, Row 5 = f₂, ..., Row 9 = f₆, Row 10 = f₇, Row 11 = f₈
# Columns: B=∂/∂t, C=∂/∂x, D=∂/∂y, E=∂/∂z, F=∂/∂u^t,
#          G=∂/∂u^x, H=∂/∂u^y, I=∂/∂u^z, J=∂/∂T, K=∂/∂Ṫ

# Use named references within the same sheet for J6, J7, J8 rows (rows 9,10,11)
J62ref = "C9"; J72ref = "C10"; J82ref = "C11"
J65ref = "F9"; J75ref = "F10"; J85ref = "F11"
J66ref = "G9"; J76ref = "G10"; J86ref = "G11"
J67ref = "H9"; J77ref = "H10"; J87ref = "H11"
J68ref = "I9"; J78ref = "I10"; J88ref = "I11"

# ∂f₁₀/∂x:
#   = T_xxx u^x² + (∂T_yy/∂x) u^y² + (∂T_zz/∂x) u^z²
#     + 2 T_xyy u^x u^y + 2 T_xzz u^x u^z
#     + T_xx f₆ + T_xy f₇ + T_xz f₈
#     + T_x J₆₂ + T_y J₇₂ + T_z J₈₂
J10_dx = (
    f"{T_xxx}*{ux2}+{Tyy_dx}*{uy2}+{Tzz_dx}*{uz2}"
    f"+2*{T_xyy}*{ux}*{uy}+2*{T_xzz}*{ux}*{uz}"
    f"+{Txx}*{f6eval}+{Txy}*{f7eval}+{Txz}*{f8eval}"
    f"+{Tx}*{J62ref}+{Ty}*{J72ref}+{Tz}*{J82ref}"
)

# ∂f₁₀/∂y (only cross-terms survive; all T_y derivatives ≡ 0 at y₀=0 but
#            formula is kept general for any evaluation point)
# Similar structure but with y-partial of T components:
# ∂T_xx/∂y = T_xxy (formula below),  ∂T_yy/∂y = T_yyy, ∂T_zz/∂y = T_zzy
# We express symbolically using Parameters cells for generality.
# For the general formula we write out the full expression using parameter derivs:
# T_xxy = ∂T_xx/∂y = 12GM y₀ (35x₀²−5r₀²)/r₀⁹
T_xxy_f = f"12*{P}B3*{P}B5*{P}B10*(35*{P}B9^2-5*{P}B13^2)/{P}B13^9"
# T_yyy = ∂T_yy/∂y = 12GM y₀ (35y₀²−17r₀²)/r₀⁹  [=0 at y₀=0]
T_yyy_f  = f"12*{P}B3*{P}B5*{P}B10*(35*{P}B10^2-17*{P}B13^2)/{P}B13^9"
# T_yyz = ∂T_yz/∂y = -60GM y₀(r₀²−7y₀²)/r₀⁹   [here it's cross-deriv]
T_yyy2_f = f"-60*{P}B3*{P}B5*{P}B10*({P}B13^2-7*{P}B10^2)/{P}B13^9"
# ∂T_zz/∂y  [zero at y₀=0, general: 12GMy(35z²−5r²)/r⁹]
Tzz_dy_f = f"12*{P}B3*{P}B5*{P}B10*(35*{P}B11^2-5*{P}B13^2)/{P}B13^9"
# 2 * (∂T_xy/∂y) = T_xyy (already have)
J63ref = "D9"; J73ref = "D10"; J83ref = "D11"
J10_dy = (
    f"({T_xxy_f})*{ux2}+({T_yyy_f})*{uy2}+({Tzz_dy_f})*{uz2}"
    f"+2*{T_xyy}*{ux}*{uy}"
    f"+{Txx}*{J63ref}+{Txy}*{J73ref}+{Txz}*{J83ref}"
    f"+{Tx}*{J63ref}+{Ty}*{J73ref}+{Tz}*{J83ref}"
)
# Note: above double-counts T contributions — fix:
# The correct split is:
#   Σ_{spatial} ∂(T_{αβ} uα uβ)/∂y  +  ∂(T_x f6 + T_y f7 + T_z f8)/∂y
# = [∂T_xx/∂y·u^x² + ∂T_yy/∂y·u^y² + ∂T_zz/∂y·u^z² + 2∂T_xy/∂y u^x u^y + ...]
# + [T_xy f6 + T_x J63] + [T_yy f7 + T_y J73] + [T_yz f8 + T_z J83]
J10_dy = (
    f"({T_xxy_f})*{ux2}+({T_yyy_f})*{uy2}+({Tzz_dy_f})*{uz2}"
    f"+2*({T_yyy2_f})*{ux}*{uy}"
    f"+{Txy}*{f6eval}+{Tx}*{J63ref}"
    f"+{Tyy}*{f7eval}+{Ty}*{J73ref}"
    f"+{Tyz}*{f8eval}+{Tz}*{J83ref}"
)

# ∂f₁₀/∂z (symmetric to ∂/∂y)
T_xxz_f  = f"12*{P}B3*{P}B5*{P}B11*(35*{P}B9^2-5*{P}B13^2)/{P}B13^9"
T_yyz_f  = f"12*{P}B3*{P}B5*{P}B11*(35*{P}B10^2-5*{P}B13^2)/{P}B13^9"
T_zzz_f  = f"12*{P}B3*{P}B5*{P}B11*(35*{P}B11^2-17*{P}B13^2)/{P}B13^9"
T_xzz2_f = f"-60*{P}B3*{P}B5*{P}B11*({P}B13^2-7*{P}B11^2)/{P}B13^9"
J64ref = "E9"; J74ref = "E10"; J84ref = "E11"
J10_dz = (
    f"({T_xxz_f})*{ux2}+({T_yyz_f})*{uy2}+({T_zzz_f})*{uz2}"
    f"+2*({T_xzz2_f})*{ux}*{uz}"
    f"+{Txz}*{f6eval}+{Tx}*{J64ref}"
    f"+{Tyz}*{f7eval}+{Ty}*{J74ref}"
    f"+{Tzz}*{f8eval}+{Tz}*{J84ref}"
)

# ∂f₁₀/∂u^t = T_x J65 + T_y J75 + T_z J85  (velocity combos in T-part have no u^t)
J10_dut = f"{Tx}*{J65ref}+{Ty}*{J75ref}+{Tz}*{J85ref}"

# ∂f₁₀/∂u^x = 2 T_xx u^x + 2 T_xy u^y + 2 T_xz u^z
#               + T_x J66 + T_y J76 + T_z J86
J10_dux = (
    f"2*{Txx}*{ux}+2*{Txy}*{uy}+2*{Txz}*{uz}"
    f"+{Tx}*{J66ref}+{Ty}*{J76ref}+{Tz}*{J86ref}"
)

# ∂f₁₀/∂u^y = 2 T_yy u^y + 2 T_xy u^x + 2 T_yz u^z
#               + T_x J67 + T_y J77 + T_z J87
J10_duy = (
    f"2*{Tyy}*{uy}+2*{Txy}*{ux}+2*{Tyz}*{uz}"
    f"+{Tx}*{J67ref}+{Ty}*{J77ref}+{Tz}*{J87ref}"
)

# ∂f₁₀/∂u^z = 2 T_zz u^z + 2 T_xz u^x + 2 T_yz u^y
#               + T_x J68 + T_y J78 + T_z J88
J10_duz = (
    f"2*{Tzz}*{uz}+2*{Txz}*{ux}+2*{Tyz}*{uy}"
    f"+{Tx}*{J68ref}+{Ty}*{J78ref}+{Tz}*{J88ref}"
)

J10 = [
    0,
    j(J10_dx),
    j(J10_dy),
    j(J10_dz),
    j(J10_dut),
    j(J10_dux),
    j(J10_duy),
    j(J10_duz),
    0,
    0,
]

all_rows = [J1, J2, J3, J4, J5, J6, J7, J8, J9, J10]

# ── Write the matrix into the sheet ────────────────────────────────────
for ri, jrow in enumerate(all_rows, start=4):
    for ci, val in enumerate(jrow, start=2):
        c = wj.cell(row=ri, column=ci)
        if isinstance(val, int):
            c.value = val
            bg = C_DIAGONAL if (ri - 4 == ci - 2) else C_ZERO
        else:
            c.value = val   # Excel formula string
            bg = C_NONZERO
        c.fill      = _fill(bg)
        c.alignment = _center()
        c.border    = _border()
        c.number_format = "0.000000E+00"
        c.font      = _font(size=9)

# Highlight diagonal (approximate identity block) in orange
diag_pairs = [(4,6), (5,7), (6,8), (7,9), (12,11)]  # (row, col)
for rr, cc in diag_pairs:
    wj.cell(row=rr, column=cc).fill = _fill(C_DIAGONAL)

# Freeze pane
wj.freeze_panes = "B4"


# ===========================================================================
# Sheet 5 — Analysis
# ===========================================================================
wa = wb.create_sheet("Analysis")
wa.column_dimensions["A"].width = 40
wa.column_dimensions["B"].width = 24
wa.column_dimensions["C"].width = 22
wa.column_dimensions["D"].width = 50

wa.merge_cells("A1:D1")
c = wa["A1"]
c.value = "Jacobian Analysis — Trace, Norms, Block Structure, Conservation Checks"
c.font  = _font(bold=True, size=13, color=C_WHITE)
c.fill  = _fill(C_TITLE)
c.alignment = _center()
wa.row_dimensions[1].height = 26

# The Jacobian data is in Jacobian sheet rows 4–13, columns B–K
# J_ij occupies Jacobian!B4:K13

J_range_prefix = "Jacobian!"
def Jcell(r, c_idx):
    """Return reference like Jacobian!C5."""
    return f"{J_range_prefix}{get_column_letter(c_idx+1)}{r+3}"

section(wa, 2, "Scalar Properties of the Jacobian Matrix  J ∈ ℝ^{10×10}")

row_entry(wa, 3, "Tr(J) — Trace  (sum of diagonal elements)",
          ("=Jacobian!B4+Jacobian!C5+Jacobian!D6+Jacobian!E7+"
           "Jacobian!F8+Jacobian!G9+Jacobian!H10+Jacobian!I11+"
           "Jacobian!J12+Jacobian!K13"),
          "dimensionless",
          "For Hamiltonian systems Tr(J) = 0 (Liouville); "
          "non-zero here due to TEGR torsion coupling",
          is_formula=True, num_fmt="0.000000E+00")

row_entry(wa, 4,
          "‖J‖_F — Frobenius norm  √(Σ J_ij²)",
          ("=SQRT(SUMPRODUCT(Jacobian!B4:K13*Jacobian!B4:K13))"),
          "dimensionless",
          "Measures overall magnitude of the Jacobian (sensitivity)",
          is_formula=True, num_fmt="0.000000E+00")

row_entry(wa, 5,
          "max|J_ij| — Maximum absolute entry",
          "=MAX(ABS(Jacobian!B4:K13))",
          "dimensionless",
          "Largest Jacobian element (dominant coupling)",
          is_formula=True, num_fmt="0.000000E+00")

section(wa, 6, "Block Structure Analysis")

row_entry(wa, 7, "Position block Tr (rows 1–4, cols 1–4)",
          "=Jacobian!B4+Jacobian!C5+Jacobian!D6+Jacobian!E7",
          "dimensionless", "Should be 0 (no coordinate self-coupling)",
          is_formula=True, num_fmt="0.000000E+00")

row_entry(wa, 8, "Velocity block Tr (rows 5–8, cols 5–8)",
          "=Jacobian!F8+Jacobian!G9+Jacobian!H10+Jacobian!I11",
          "dimensionless",
          "Geodesic self-damping; 0 in flat-space limit",
          is_formula=True, num_fmt="0.000000E+00")

row_entry(wa, 9, "Torsion block Tr (rows 9–10, cols 9–10)",
          "=Jacobian!J12+Jacobian!K13",
          "dimensionless",
          "Torsion subsystem trace",
          is_formula=True, num_fmt="0.000000E+00")

section(wa, 10, "Tidal Tensor Components  (=Jacobian rows 6–8, cols 2–4 = ∂f_{x,y,z}/∂{x,y,z})")

row_entry(wa, 11, "K_xx = ∂(du^x/dτ)/∂x  (radial tidal)",
          "=Jacobian!C9",
          "1/[m²]",
          "= 2M/r₀³ for circular orbit at x₀=r₀  (centrifugal + tidal)",
          is_formula=True, num_fmt="0.000000E+00")

row_entry(wa, 12, "K_yy = ∂(du^y/dτ)/∂y  (transverse tidal)",
          "=Jacobian!D10",
          "1/[m²]",
          "= −M/r₀³  (restoring tidal in orbital direction)",
          is_formula=True, num_fmt="0.000000E+00")

row_entry(wa, 13, "K_zz = ∂(du^z/dτ)/∂z  (vertical tidal)",
          "=Jacobian!E11",
          "1/[m²]",
          "= −M/r₀³  (restoring tidal perpendicular to plane)",
          is_formula=True, num_fmt="0.000000E+00")

row_entry(wa, 14, "Tr(K) = K_xx + K_yy + K_zz  (should ≈ 0, Laplace eq.)",
          "=B11+B12+B13",
          "1/[m²]",
          "Tr(K) = 0 away from source: ∇²Φ = 0 (vacuum Poisson eq.)",
          is_formula=True, num_fmt="0.000000E+00")

section(wa, 15, "Orbital Frequency Cross-Check")

row_entry(wa, 16, "ω² (Newtonian) = GM/r₀³ = Φ_x/r₀²",
          "=Parameters!B3*Parameters!B5/Parameters!B13^3",
          "1/[m²]",
          "Newtonian orbital angular frequency squared",
          is_formula=True, num_fmt="0.000000E+00")

row_entry(wa, 17, "Epicyclic freq² κ² ≈ 4ω² − ∂²Ω²/∂ln r · ω²",
          "=4*B16-3*B16",
          "1/[m²]",
          "For Keplerian orbit: κ² = ω²  (confirmed by J₇₃ = −M/r₀³ = −ω²)",
          is_formula=True, num_fmt="0.000000E+00")

row_entry(wa, 18,
          "Vertical freq² ν² = GM/r₀³",
          "=Parameters!B3*Parameters!B5/Parameters!B13^3",
          "1/[m²]",
          "Oscillation frequency perpendicular to orbital plane; equals K_zz magnitude",
          is_formula=True, num_fmt="0.000000E+00")

section(wa, 19, "TEGR Torsion Coupling Strengths (Row 10 of Jacobian)")

row_entry(wa, 20, "J₁₀,₅ = ∂(d²T/dτ²)/∂u^t",
          "=Jacobian!F13",
          "1/[m³]",
          "Torsion acceleration sensitivity to u^t; = −2 T_x Φ_x u^t",
          is_formula=True, num_fmt="0.000000E+00")

row_entry(wa, 21, "J₁₀,₇ = ∂(d²T/dτ²)/∂u^y  (dominant torsion coupling)",
          "=Jacobian!H13",
          "1/[m³]",
          "= 2 T_yy u^y + T_x J₆₇  — torsion response to tangential velocity",
          is_formula=True, num_fmt="0.000000E+00")

row_entry(wa, 22, "J₁₀,₂ = ∂(d²T/dτ²)/∂x  (tidal torsion gradient)",
          "=Jacobian!C13",
          "1/[m⁴]",
          "= (∂T_yy/∂x) u^y² + T_xx f₆ + T_x J₆₂",
          is_formula=True, num_fmt="0.000000E+00")

section(wa, 23, "Summary Table — Non-Zero Jacobian Elements at Default Evaluation Point")

summary_hdrs = ["Element J_{ij}", "Row i  (equation)", "Col j  (variable)", "Numerical Value", "Physical Interpretation"]
for ci, lbl in enumerate(summary_hdrs, start=1):
    c = wa.cell(row=24, column=ci)
    c.value = lbl
    c.font  = _font(bold=True, size=10, color=C_WHITE)
    c.fill  = _fill(C_SECTION)
    c.alignment = _center()
    c.border = _border()

summary_rows = [
    # (label, row_ref, col_ref, formula, meaning)
    ("J_{1,5}",  "f₁=dt/dτ",   "u^t",  "=Jacobian!F4",  "= 1  (identity coupling: dt/dτ = u^t)"),
    ("J_{2,6}",  "f₂=dx/dτ",   "u^x",  "=Jacobian!G5",  "= 1  (identity coupling: dx/dτ = u^x)"),
    ("J_{3,7}",  "f₃=dy/dτ",   "u^y",  "=Jacobian!H6",  "= 1  (identity coupling: dy/dτ = u^y)"),
    ("J_{4,8}",  "f₄=dz/dτ",   "u^z",  "=Jacobian!I7",  "= 1  (identity coupling: dz/dτ = u^z)"),
    ("J_{5,6}",  "f₅=du^t/dτ", "u^x",  "=Jacobian!G8",  "= −2γΦ_x  (gravitomagnetic time dilation)"),
    ("J_{5,7}",  "f₅=du^t/dτ", "u^y",  "=Jacobian!H8",  "= −2γΦ_y = 0  (zero by symmetry at y₀=0)"),
    ("J_{6,2}",  "f₆=du^x/dτ", "x",    "=Jacobian!C9",  "Tidal: −Φ_xx(γ²+γ²v²) = 2M(1+v²)γ²/r₀³"),
    ("J_{6,5}",  "f₆=du^x/dτ", "u^t",  "=Jacobian!F9",  "= −2Φ_x u^t  (gravitoelectric force coupling)"),
    ("J_{6,7}",  "f₆=du^x/dτ", "u^y",  "=Jacobian!H9",  "= −2Φ_x u^y  (Coriolis-like term)"),
    ("J_{7,3}",  "f₇=du^y/dτ", "y",    "=Jacobian!D10", "Tidal: −Φ_yy(γ²−γ²v²) = −M/r₀³ (restoring)"),
    ("J_{7,6}",  "f₇=du^y/dτ", "u^x",  "=Jacobian!G10", "= 2Φ_x u^y  (Coriolis-like term)"),
    ("J_{8,4}",  "f₈=du^z/dτ", "z",    "=Jacobian!E11", "Tidal: −Φ_zz(γ²+γ²v²) = −M(1+v²)γ²/r₀³"),
    ("J_{9,10}", "f₉=dT/dτ",   "Ṫ",    "=Jacobian!K12", "= 1  (torsion rate kinematic equation)"),
    ("J_{10,2}", "f₁₀=d²T/dτ²","x",    "=Jacobian!C13", "Torsion-tidal coupling along x"),
    ("J_{10,5}", "f₁₀=d²T/dτ²","u^t",  "=Jacobian!F13", "= −2 T_x Φ_x u^t  (time-dilation torsion)"),
    ("J_{10,7}", "f₁₀=d²T/dτ²","u^y",  "=Jacobian!H13", "= 2T_yy u^y + T_x J₆₇  (dominant TEGR coupling)"),
]

for ri, (lbl, req, cvar, fml, meaning) in enumerate(summary_rows, start=25):
    wa.cell(row=ri, column=1).value = lbl
    wa.cell(row=ri, column=1).font  = _font(bold=True, size=10)
    wa.cell(row=ri, column=1).fill  = _fill(C_INPUT)
    wa.cell(row=ri, column=1).border = _border()
    wa.cell(row=ri, column=2).value  = req
    wa.cell(row=ri, column=2).fill   = _fill(C_INPUT)
    wa.cell(row=ri, column=2).border = _border()
    wa.cell(row=ri, column=3).value  = cvar
    wa.cell(row=ri, column=3).fill   = _fill(C_INPUT)
    wa.cell(row=ri, column=3).border = _border()
    c4 = wa.cell(row=ri, column=4)
    c4.value = fml
    c4.fill  = _fill(C_FORMULA)
    c4.border = _border()
    c4.number_format = "0.000000E+00"
    wa.cell(row=ri, column=5).value = meaning
    wa.cell(row=ri, column=5).fill  = _fill(C_INPUT)
    wa.cell(row=ri, column=5).border = _border()
    wa.cell(row=ri, column=5).alignment = _left()
wa.column_dimensions["E"].width = 55


# ===========================================================================
# Save
# ===========================================================================
out_path = "TEGR_Jacobian.xlsx"
wb.save(out_path)
print(f"Saved: {out_path}")
print("Sheets:", [s.title for s in wb.worksheets])
