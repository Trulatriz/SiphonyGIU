import pandas as pd

# Okabe-Ito color-blind-safe palette
OKABE_ITO = [
    "#000000",  # black
    "#E69F00",  # orange
    "#56B4E9",  # sky blue
    "#009E73",  # bluish green
    "#F0E442",  # yellow
    "#0072B2",  # blue
    "#D55E00",  # vermillion
    "#CC79A7",  # reddish purple
]

RHO_FOAM_G = "\u03c1 foam (g/cm^3)"
RHO_FOAM_KG = "\u03c1 foam (kg/m^3)"
RHO_REL = "\u03c1\u1d63"
DESV_RHO_FOAM_G = "Desvest \u03c1 foam (g/cm^3)"
DESV_RHO_FOAM_KG = "Desvest \u03c1 foam (kg/m^3)"

# Exact canonical column names from CombineModule.new_column_order
INDEPENDENT_OPTIONS = [
    ("m(g)", "m(g)", r"$m\;(\mathrm{g})$"),
    ("Water (g)", "Water (g)", r"$\mathrm{Water}\;(\mathrm{g})$"),
    ("Tsat (\u00b0C)", "T (\u00b0C)", r"$T_{\mathrm{sat}}\;({}^\circ\mathrm{C})$"),
    ("Psat (MPa)", "Psat (MPa)", r"$P_{\mathrm{sat}}\;(\mathrm{MPa})$"),
    ("tsat (min)", "t (min)", r"$t_{\mathrm{sat}}\;(\mathrm{min})$"),
    ("PDR (MPa/s)", "PDR (MPa/s)", r"$\mathrm{PDR}\;(\mathrm{MPa}/\mathrm{s})$"),
    ("Additive %", "Additive %", r"\text{Additive }(\%)"),
    ("Additive", "Additive", r"\text{Additive}"),
    ("Base Polymer", "Base Polymer", r"\text{Base Polymer}"),
]

INDEPENDENTS = [display for display, _column, _latex in INDEPENDENT_OPTIONS]
INDEPENDENT_TO_COLUMN = {display: column for display, column, _latex in INDEPENDENT_OPTIONS}
COLUMN_TO_INDEPENDENT = {column: display for display, column, _latex in INDEPENDENT_OPTIONS}
INDEPENDENT_COLUMNS = [column for _display, column, _latex in INDEPENDENT_OPTIONS]
INDEPENDENT_LATEX = {display: latex for display, _column, latex in INDEPENDENT_OPTIONS}

DEPENDENT_OPTIONS = [
    ("\u00f8 (\u00b5m)", "\u00f8 (\u00b5m)", r"$\varnothing\;(\mu\mathrm{m})$"),
    ("N\u1d65 (cells\u00b7cm^3)", "N\u1d65 (cells\u00b7cm^3)", r"$N_{\mathrm{v}}\;(\mathrm{cells}/\mathrm{cm}^3)$"),
    ("\u03c1f (g/cm^3)", RHO_FOAM_G, r"$\rho_{f}\;(\mathrm{g}/\mathrm{cm}^3)$"),
    ("\u03c1f (kg/m^3)", RHO_FOAM_KG, r"$\rho_{f}\;(\mathrm{kg}/\mathrm{m}^3)$"),
    ("\u03c1\u1d63", RHO_REL, r"$\rho_{r}$"),
    ("PDR (MPa/s)", "PDR (MPa/s)", r"$\mathrm{PDR}\;(\mathrm{MPa}/\mathrm{s})$"),
    ("X", "X", r"$X$"),
    ("Ov (%)", "OC (%)", r"$O_{\mathrm{v}}\;(\%)$"),
    ("Tm (\u00b0C)", "DSC Tm (\u00b0C)", r"$T_{\mathrm{m}}\;({}^\circ\mathrm{C})$"),
    ("Tg (\u00b0C)", "DSC Tg (\u00b0C)", r"$T_{\mathrm{g}}\;({}^\circ\mathrm{C})$"),
    ("\u03c7c (%)", "DSC Xc (%)", r"$\chi_{c}\;(\%)$"),
]

DEPENDENT_LABELS = [label for label, _column, _latex in DEPENDENT_OPTIONS]
DEPENDENT_MAP = {label: column for label, column, _latex in DEPENDENT_OPTIONS}
DEPENDENT_COLUMN_TO_LABEL = {column: label for label, column, _latex in DEPENDENT_OPTIONS}
DEPENDENT_LATEX = {label: latex for label, _column, latex in DEPENDENT_OPTIONS}
DEPENDENT_COLUMNS = [column for _label, column, _latex in DEPENDENT_OPTIONS]
DEPENDENTS = DEPENDENT_LABELS

DEPENDENT_TO_DEVIATION = {
    "\u00f8 (\u00b5m)": "Desvest \u00f8 (\u00b5m)",
    "N\u1d65 (cells\u00b7cm^3)": "Desvest N\u1d65 (cells\u00b7cm^3)",
    "\u03c1f (g/cm^3)": DESV_RHO_FOAM_G,
    "\u03c1f (kg/m^3)": DESV_RHO_FOAM_KG,
}

DEVIATIONS = DEPENDENT_TO_DEVIATION.copy()

LEGACY_DEPENDENT_LABELS = {
    "\u03c1f (g/cm^3)": RHO_FOAM_G,
    "\u03c1f (kg/m^3)": RHO_FOAM_KG,
    "Ov (%)": "OC (%)",
    "Tm (\u00b0C)": "DSC Tm (\u00b0C)",
    "Tg (\u00b0C)": "DSC Tg (\u00b0C)",
    "\u03c7c (%)": "DSC Xc (%)",
    "DSC Xc (\u00b0C)": "DSC Xc (%)",
}


def normalize_numeric_series(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype=float)
    cleaned = series.astype(str).str.replace(r"\s", "", regex=True).str.replace(",", ".", regex=False)
    return pd.to_numeric(cleaned, errors="coerce")


def dependent_latex(label: str) -> str:
    return DEPENDENT_LATEX.get(label, label)


def independent_latex(label: str) -> str:
    return INDEPENDENT_LATEX.get(label, label)


def friendly_column_name(column: str) -> str:
    if column in DEPENDENT_COLUMN_TO_LABEL:
        return DEPENDENT_COLUMN_TO_LABEL[column]
    if column in COLUMN_TO_INDEPENDENT:
        return COLUMN_TO_INDEPENDENT[column]
    return column


def augment_density_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure derived density columns exist for plotting."""
    result = df.copy()

    if RHO_REL not in result.columns:
        for legacy in ("\u03c1\u1d63", "rho_r"):
            if legacy in result.columns:
                result[RHO_REL] = result[legacy]
                break
    for legacy in ("\u03c1\u1d63", "rho_r"):
        if legacy in result.columns and legacy != RHO_REL:
            result = result.drop(columns=[legacy])

    if RHO_FOAM_G in result.columns and RHO_FOAM_KG not in result.columns:
        result[RHO_FOAM_KG] = normalize_numeric_series(result[RHO_FOAM_G]) * 1000
    if DESV_RHO_FOAM_G in result.columns and DESV_RHO_FOAM_KG not in result.columns:
        result[DESV_RHO_FOAM_KG] = normalize_numeric_series(result[DESV_RHO_FOAM_G]) * 1000

    if "Psat (MPa)" not in result.columns and "P CO2 (bar)" in result.columns:
        result["Psat (MPa)"] = pd.to_numeric(result["P CO2 (bar)"], errors="coerce") / 10

    return result
