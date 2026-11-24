"""
Motor de càlcul GRADISTAT (mostra única) en Python.

Port fidel de les fórmules del full "Calculations" i de les macros VBA
FolkAndWardMacro, Folk54Macro, Triangle i Triangle2 de GRADISTAT v9.1
(Blott & Pye, 2001, Earth Surface Processes and Landforms 26, 1237-1248).

Convencions (com a l'Excel):
  * Obertures ordenades de més grossa a més fina; l'última fila és el "pan"
    (obertura 0) amb el material que ha passat per tots els tamisos.
  * El pes de cada fila és el material retingut en aquell tamís.
  * phi = -log2(mm).
  * Els valors d'error d'Excel (#NUM!) es representen amb NaN.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

import numpy as np

NROWS = 221  # files del full Calculations (A30:A250)
NAN = float("nan")


def _isnan(x) -> bool:
    return isinstance(x, float) and math.isnan(x)


# ---------------------------------------------------------------------------
# Plantilles d'obertures (µm), copiades de les macros de l'Excel
# ---------------------------------------------------------------------------
TEMPLATES = {
    "Tamisos ½ φ (per defecte)": [90000, 63000, 45000, 31500, 22400, 16000, 11200, 8000, 5600,
                                  4000, 2800, 2000, 1400, 1000, 710, 500, 355, 250, 180, 125, 90, 63],
    "Tamisos ¼ φ": [75000, 63000, 53000, 45000, 37500, 31500, 26500, 22400, 19000, 16000, 13200,
                    11200, 9500, 8000, 6700, 5600, 4750, 4000, 3350, 2800, 2360, 2000, 1700, 1400,
                    1180, 1000, 850, 710, 600, 500, 425, 355, 300, 250, 212, 180, 150, 125, 106,
                    90, 75, 63],
    "Làser Coulter LS230": [75000, 63000, 45000, 31500, 22400, 16000, 11200, 8000, 5600, 4000, 2800,
                            2000, 1822, 1660, 1512, 1377, 1255, 1143, 1041, 948.2, 863.9, 786.9,
                            716.9, 653, 594.9, 541.9, 493.6, 449.7, 409.6, 373.1, 339.8, 309.6,
                            282.1, 256.8, 234.1, 213.2, 194.2, 176.8, 161.2, 146.8, 133.7, 121.8,
                            111, 101.1, 92.09, 83.9, 76.43, 69.62, 63.41, 57.77, 52.63, 47.93,
                            43.66, 39.77, 36.24, 33, 30.07, 27.38, 24.95, 22.73, 20.7, 18.86,
                            17.17, 15.65, 14.26, 12.99, 11.83, 10.78, 9.819, 8.944, 8.147, 7.421,
                            6.761, 6.158, 5.611, 5.111, 4.656, 4.241, 3.862, 3.519, 3.206, 2.92,
                            2.66, 2.423, 2.207, 2.01, 1.832, 1.669, 1.52, 1.385, 1.261, 1.149,
                            1.047, 0.953, 0.869, 0.791, 0.721, 0.657, 0.598, 0.545, 0.496, 0.452,
                            0.412, 0.375, 0.342, 0.311, 0.284, 0.258, 0.235, 0.214, 0.195, 0.178,
                            0.162, 0.148, 0.134, 0.122, 0.112, 0.102, 0.093, 0.084, 0.077, 0.07,
                            0.064, 0.058, 0.053, 0.048, 0.044, 0.04],
}

CLASS_KEYS = ["vc_gravel", "c_gravel", "m_gravel", "f_gravel", "vf_gravel",
              "vc_sand", "c_sand", "m_sand", "f_sand", "vf_sand",
              "vc_silt", "c_silt", "m_silt", "f_silt", "vf_silt", "clay"]
CLASS_LABELS = {
    "vc_gravel": "V Coarse Gravel", "c_gravel": "Coarse Gravel", "m_gravel": "Medium Gravel",
    "f_gravel": "Fine Gravel", "vf_gravel": "V Fine Gravel",
    "vc_sand": "V Coarse Sand", "c_sand": "Coarse Sand", "m_sand": "Medium Sand",
    "f_sand": "Fine Sand", "vf_sand": "V Fine Sand",
    "vc_silt": "V Coarse Silt", "c_silt": "Coarse Silt", "m_silt": "Medium Silt",
    "f_silt": "Fine Silt", "vf_silt": "V Fine Silt", "clay": "Clay",
}


# ---------------------------------------------------------------------------
# Resultat
# ---------------------------------------------------------------------------
@dataclass
class Result:
    # taula de càlcul (com el full Calculations)
    table: dict = field(default_factory=dict)
    # mètode dels moments
    moments: dict = field(default_factory=dict)
    # Folk & Ward
    folk_ward: dict = field(default_factory=dict)
    descriptions: dict = field(default_factory=dict)
    # percentils i modes
    modes: list = field(default_factory=list)          # [(µm o "Pan", phi o "Pan"), ...]
    d_values: dict = field(default_factory=dict)
    sample_type: str = ""
    # composició
    gravel: float | None = None
    sand: float | None = None
    mud: float | None = None
    classes: dict = field(default_factory=dict)        # valor (%) o "-"
    textural_group: str | None = None
    sediment_name: str | None = None
    sand_silt_clay_group: str | None = None
    silt: float | None = None
    clay: float | None = None
    triangle_gsm: tuple | None = None                  # (x, y) en coordenades del diagrama
    triangle_ssc: tuple | None = None
    sieving_error: float | None = None                 # fracció (0.012 = 1.2 %)
    warnings: list = field(default_factory=list)


# ---------------------------------------------------------------------------
# Full "Calculations"
# ---------------------------------------------------------------------------
def build_rows(apertures_um, weights, pan_weight=0.0):
    """Construeix les columnes A (mm) i G (pes) tal com les copia la macro."""
    pairs = [(float(a), float(w or 0.0)) for a, w in zip(apertures_um, weights)
             if a is not None and not _isnan(float(a)) and float(a) > 0]
    pairs.sort(key=lambda p: -p[0])
    A = np.zeros(NROWS)
    G = np.zeros(NROWS)
    n = min(len(pairs), NROWS - 1)
    for k in range(n):
        A[k] = pairs[k][0] / 1000.0
        G[k] = pairs[k][1]
    G[n] = float(pan_weight or 0.0)  # fila del pan (A = 0)
    return A, G


def calculations_sheet(A, G):
    n = NROWS
    total = float(np.sum(G))
    F = G * 100.0 / total
    E = np.zeros(n)
    E[0] = F[0]
    for k in range(1, n):
        E[k] = F[k] + E[k - 1]

    # B: punt mig aritmètic (mm). B30 és buit.
    B = np.zeros(n)
    if n > 1:
        B[1] = (A[0] + A[1]) / 2
    for k in range(2, n):
        if A[k] == 0:
            if A[k - 1] > 0:
                B[k] = ((A[k - 1] + 2 ** -10) / 2) if B[k - 1] > 0.001 else 0.0
            else:
                B[k] = 0.0
        else:
            B[k] = (A[k - 1] + A[k]) / 2

    # C: obertura en phi (amb el pan = 10 phi)
    C = np.full(n, NAN)
    prev = 0.0  # C29 és buit
    for k in range(n):
        if A[k] == 0:
            if B[k] > 0 and not _isnan(prev) and prev < 10:
                C[k] = 10.0
            else:
                C[k] = NAN
        else:
            C[k] = -math.log(A[k], 2)
        prev = C[k]

    # D: punt mig en phi (D30 buit -> 0 en les fórmules)
    D = np.zeros(n)
    D[1:] = (C[:-1] + C[1:]) / 2

    # U: log10 del punt mig geomètric en µm (U30 buit -> 0)
    with np.errstate(over="ignore", invalid="ignore"):
        U = np.zeros(n)
        U[1:] = np.log10((2.0 ** (-D[1:])) * 1000)

    def s(x):
        x = np.asarray(x, dtype=float)
        return float(np.nansum(x[np.isfinite(x)]))

    # Moments logarítmics (phi)
    mean_phi = s(D * F) / 100
    with np.errstate(invalid="ignore", divide="ignore"):
        sd_phi = math.sqrt(s(F * (D - mean_phi) ** 2) / 100)
        sk_phi = _div(s(F * (D - mean_phi) ** 3), 100 * sd_phi ** 3)
        ku_phi = _div(s(F * (D - mean_phi) ** 4), 100 * sd_phi ** 4)

        # Moments aritmètics (µm)
        Bu = B * 1000
        mean_a = s(Bu * F) / 100
        sd_a = math.sqrt(s(F * (Bu - mean_a) ** 2) / 100)
        sk_a = _div(s(F * (Bu - mean_a) ** 3), 100 * sd_a ** 3)
        ku_a = _div(s(F * (Bu - mean_a) ** 4), 100 * sd_a ** 4)

        # Moments geomètrics (µm)
        mean_g = 10 ** (s(U * F) / 100)
        lm = math.log10(mean_g)
        sd_g = 10 ** math.sqrt(s(F * (U - lm) ** 2) / 100)
        sk_g = _div(s(F * (U - lm) ** 3), 100 * math.log10(sd_g) ** 3)
        ku_g = _div(s(F * (U - lm) ** 4), 100 * math.log10(sd_g) ** 4)

    # Pes normalitzat per amplada de classe (N) i en % (O)
    kpan = _first_pan(A)
    N = np.zeros(n)
    for k in range(1, kpan):
        diff = C[k] - C[k - 1]
        N[k] = F[k] / diff
    a13 = float(np.sum(N[1:][N[1:] > 0]))
    O = np.zeros(n)
    O[1:] = N[1:] * 100 / a13 if a13 else 0.0

    moments = {
        "arithmetic": dict(mean=mean_a, sorting=sd_a, skewness=sk_a, kurtosis=ku_a),
        "geometric": dict(mean=mean_g, sorting=sd_g, skewness=sk_g, kurtosis=ku_g),
        "logarithmic": dict(mean=mean_phi, sorting=sd_phi, skewness=sk_phi, kurtosis=ku_phi),
    }
    table = dict(A=A, B=B, C=C, D=D, E=E, F=F, G=G, N=N, O=O, total=total, kpan=kpan)
    return table, moments


def _div(a, b):
    try:
        return a / b
    except ZeroDivisionError:
        return NAN


def _first_pan(A):
    for k in range(NROWS):
        if A[k] == 0:
            return k
    return NROWS


# ---------------------------------------------------------------------------
# FolkAndWardMacro
# ---------------------------------------------------------------------------
class _Pct:
    """Interpolació de percentils amb l'estat que arrossega la macro VBA."""

    def __init__(self, t, res: Result):
        self.t = t
        self.res = res
        self.gp = 0.0   # GreaterPercent
        self.gphi = 0.0  # GreaterPhi
        self.lp = 0.0   # LessPercent
        self.lphi = 0.0  # LessPhi

    def __call__(self, p, pan_check=True):
        A, C, E = self.t["A"], self.t["C"], self.t["E"]
        for k in range(NROWS):
            if E[k] > p:
                self.gp = E[k]
                if pan_check and A[k] == 0:
                    self.gphi = 8.0
                    self.res.warnings.append("pan")
                else:
                    self.gphi = C[k]
                break
        for k in range(NROWS - 1, -1, -1):
            if E[k] < p:
                self.lp = E[k]
                self.lphi = C[k]
                break
        for k in range(NROWS):
            if E[k] == p:
                return C[k], A[k]
        r1 = (p - self.lp) / (self.gp - self.lp)
        phi = self.lphi + (self.gphi - self.lphi) * r1
        return phi, 1 / (2 ** phi)


def mean_desc(m):
    bounds = [(-5, "Very Coarse Gravel"), (-4, "Coarse Gravel"), (-3, "Medium Gravel"),
              (-2, "Fine Gravel"), (-1, "Very Fine Gravel"), (0, "Very Coarse Sand"),
              (1, "Coarse Sand"), (2, "Medium Sand"), (3, "Fine Sand"), (4, "Very Fine Sand"),
              (5, "Very Coarse Silt"), (6, "Coarse Silt"), (7, "Medium Silt"), (8, "Fine Silt"),
              (9, "Very Fine Silt")]
    for b, name in bounds:
        if m <= b:
            return name
    return "Clay"


def sorting_desc(s):
    if s < 0.35:
        return "Very Well Sorted"
    if s < 0.5:
        return "Well Sorted"
    if s < 0.7:
        return "Moderately Well Sorted"
    if s < 1:
        return "Moderately Sorted"
    if s < 2:
        return "Poorly Sorted"
    if s < 4:
        return "Very Poorly Sorted"
    return "Extremely Poorly Sorted"


def skew_desc(sk):
    if sk >= 0.3:
        return "Very Fine Skewed"
    if sk >= 0.1:
        return "Fine Skewed"
    if sk > -0.1:
        return "Symmetrical"
    if sk > -0.3:
        return "Coarse Skewed"
    return "Very Coarse Skewed"


def kurt_desc(k):
    if k < 0.67:
        return "Very Platykurtic"
    if k < 0.9:
        return "Platykurtic"
    if k <= 1.11:
        return "Mesokurtic"
    if k <= 1.5:
        return "Leptokurtic"
    if k <= 3:
        return "Very Leptokurtic"
    return "Extremely Leptokurtic"


def _modes(t):
    A, B, D, O = t["A"], t["B"], t["D"], t["O"]

    def o(k):
        return O[k] if 0 <= k < NROWS else 0.0

    mode_max = float(np.max(O))
    smallest = o(_first_pan(A))
    class_modes = []   # [(valor, phi, mm)]
    peaks = 0
    for level in range(5):
        found = None
        for i in range(1, NROWS + 1):
            k = i  # fila i+30 -> índex 0-based i
            if o(k - 1) < o(k) and o(k) >= o(k + 1) and o(k) > 0.15 * mode_max:
                if any(o(k) == cm[0] for cm in class_modes):
                    continue
                if o(k) == smallest:
                    found = (o(k), "Pan", "Pan")
                    break
                found = (o(k), D[k], B[k])
        if found is not None:
            class_modes.append(found)
            peaks = level + 1
        else:
            class_modes.append((None, None, None))
    vals = sorted([cm[0] for cm in class_modes if cm[0] is not None], reverse=True)
    out = []
    for rank in range(3):
        if rank >= peaks or rank >= len(vals):
            break
        target = vals[rank]
        for cm in class_modes:
            if cm[0] == target:
                phi, mm = cm[1], cm[2]
                out.append(("Pan", "Pan") if mm == "Pan" else (mm * 1000, phi))
                break
    dist = {1: "Unimodal", 2: "Bimodal", 3: "Trimodal"}.get(peaks, "Polymodal" if peaks > 3 else "")
    return out, peaks, dist


def folk_ward(t, res: Result):
    P = _Pct(t, res)
    phi5, mm5 = P(5, pan_check=False)
    phi16, mm16 = P(16)
    phi25, mm25 = P(25)
    phi50, mm50 = P(50)
    phi75, mm75 = P(75)
    phi84, mm84 = P(84)
    phi95, mm95 = P(95)
    u5, u16, u25, u50, u75, u84, u95 = (x * 1000 for x in (mm5, mm16, mm25, mm50, mm75, mm84, mm95))
    ln = math.log

    mean = (phi16 + phi50 + phi84) / 3
    sd = (phi84 - phi16) / 4 + (phi95 - phi5) / 6.6
    sk = ((phi16 + phi84 - 2 * phi50) / (2 * (phi84 - phi16))
          + (phi5 + phi95 - 2 * phi50) / (2 * (phi95 - phi5)))
    ku = (phi95 - phi5) / (2.44 * (phi75 - phi25))

    mean_um = math.exp((ln(u16) + ln(u50) + ln(u84)) / 3)
    sd_um = math.exp((ln(u16) - ln(u84)) / 4 + (ln(u5) - ln(u95)) / 6.6)
    sk_um = ((ln(u16) + ln(u84) - 2 * ln(u50)) / (2 * (ln(u16) - ln(u84)))
             + (ln(u5) + ln(u95) - 2 * ln(u50)) / (2 * (ln(u5) - ln(u95))))
    ku_um = (ln(u5) - ln(u95)) / (2.44 * (ln(u25) - ln(u75)))

    res.folk_ward = {
        "geometric": dict(mean=mean_um, sorting=sd_um, skewness=sk_um, kurtosis=ku_um),
        "logarithmic": dict(mean=mean, sorting=sd, skewness=sk, kurtosis=ku),
    }
    res.descriptions = dict(mean=mean_desc(mean), sorting=sorting_desc(sd),
                            skewness=skew_desc(sk), kurtosis=kurt_desc(ku))

    modes, peaks, dist = _modes(t)
    res.modes = modes
    res.sample_type = f"{dist}, {res.descriptions['sorting']}"
    if peaks > 1:
        res.warnings.append("multimodal")

    phi10, _ = P(10, pan_check=False)
    phi90, _ = P(90)
    d10mm, d90mm = 1 / 2 ** phi10, 1 / 2 ** phi90
    # Atenció: a GRADISTAT, la columna µm de "D10" és la mida del 10 % més fi (= phi90)
    res.d_values = {
        "D10": (d90mm * 1000, phi10),
        "D50": (1 / 2 ** phi50 * 1000, phi50),
        "D90": (d10mm * 1000, phi90),
        "D90/D10": (d10mm / d90mm, phi90 / phi10),
        "D90-D10": ((d10mm - d90mm) * 1000, phi90 - phi10),
        "D75/D25": (mm25 / mm75, phi75 / phi25),
        "D75-D25": ((mm25 - mm75) * 1000, phi75 - phi25),
    }
    return peaks


# ---------------------------------------------------------------------------
# Folk54Macro
# ---------------------------------------------------------------------------
class _FolkError(Exception):
    pass


def folk54(t, res: Result):
    A, C, E = t["A"], t["C"], t["E"]
    kp = _first_pan(A)
    i_vba = kp + 1  # índex 1-based de la fila del pan
    last_sieve_phi = C[kp - 1] if kp >= 1 else NAN
    j = i_vba if last_sieve_phi < 8 else i_vba - 1  # nombre de files usades

    def less(b):
        for k in range(j - 1, -1, -1):
            if C[k] < b:
                return C[k], E[k]
        return None

    def greater(b):
        for k in range(j):
            if C[k] > b:
                return C[k], E[k]
        return None

    def at(b, L, Gr):
        for k in range(j):
            if C[k] == b:
                return E[k]
        return L[1] + (Gr[1] - L[1]) * ((b - L[0]) / (Gr[0] - L[0]))

    def near(x, v, tol):
        return v - tol < x < v + tol

    # --- grava / sorra / fang ----------------------------------------------
    gravel = sand = mud = 0.0
    done = False
    L = less(-1)
    if L is not None:
        Gr = greater(-1)
        if Gr is None:
            if near(L[0], -1, 0.1) and near(L[1], 100, 0.001):
                gravel, sand, mud, done = 100.0, 0.0, 0.0, True
            else:
                raise _FolkError
        else:
            gravel = at(-1, L, Gr)
    if not done:
        L = less(4)
        mud_set = False
        if L is None:
            sand = 0.0
        else:
            Gr = greater(4)
            if Gr is None:
                if near(L[0], 4, 0.1):
                    sand = L[1] - gravel
                    if near(L[1], 100, 0.001):
                        mud, mud_set = 0.0, True
                else:
                    raise _FolkError
            else:
                sand = at(4, L, Gr) - gravel
        if not mud_set:
            mud = 100 - gravel - sand
            if mud < 0.0001:
                mud = 0.0
    res.gravel, res.sand, res.mud = gravel, sand, mud

    if mud == 0 and sand == 0:
        ms = 0
    elif mud == 0:
        ms = 10
    elif sand == 0:
        ms = 0.01
    else:
        ms = sand / mud

    def gclass(a, b, c, d, e):
        if gravel > 80:
            return "Gravel"
        if gravel > 30:
            return a
        if gravel > 5:
            return b
        if gravel > 0:
            return c
        return d if gravel == 0 else e

    if ms >= 9:
        tex = gclass("Sandy Gravel", "Gravelly Sand", "Slightly Gravelly Sand", "Sand", None)
    elif ms >= 1:
        tex = gclass("Muddy Sandy Gravel", "Gravelly Muddy Sand", "Slightly Gravelly Muddy Sand",
                     "Muddy Sand", None)
    elif ms >= 1 / 9:
        tex = gclass("Muddy Gravel", "Gravelly Mud", "Slightly Gravelly Sandy Mud", "Sandy Mud", None)
    else:
        tex = gclass("Muddy Gravel", "Gravelly Mud", "Slightly Gravelly Mud", "Mud", None)
    res.textural_group = tex

    # --- classes detallades -------------------------------------------------
    vals = [0.0] * 16
    bounds = list(range(-5, 10))
    stop = False

    # Very coarse gravel (límit -5 phi)
    L = less(-5)
    if L is not None:
        Gr = greater(-5)
        if Gr is None:
            if near(L[0], -5, 0.1):
                if near(L[1], 100, 0.001):
                    vals[0] = 100.0
                stop = True
            else:
                vals[0] = "-"
        else:
            vals[0] = at(-5, L, Gr)
            if near(vals[0], 100, 0.00001):
                stop = True

    cum = 0.0
    s = 1
    while not stop and s < 15:
        b = bounds[s]
        prev = vals[s - 1]
        L = less(b)
        if L is None:
            vals[s] = 0.0
            s += 1
            continue
        Gr = greater(b)
        if prev == "-":
            vals[s] = "-"
            s += 1
            continue
        if s == 1:
            base, check = vals[0], 0.0  # a la macro CumPercent encara és buit aquí
        else:
            cum = (vals[0] + vals[1]) if s == 2 else cum + prev
            base = check = cum
        if Gr is None:
            if near(L[0], b, 0.1):
                vals[s] = L[1] - base
                if near(L[1], 100, 0.001):
                    stop = True
                s += 1
                continue
            if near(check, 100, 0.0001):
                vals[s] = "-"
                s += 1
                continue
            vals[s] = 100 - base
            stop = True
            break
        vals[s] = at(b, L, Gr) - base
        tol = 0.00001 if s == 1 else 0.0001
        if near(vals[s] + base, 100, tol):
            stop = True
        s += 1
    if not stop:
        if vals[14] == "-":
            vals[15] = "-"
        else:
            cum = cum + vals[14]
            vals[15] = 100 - cum
    res.classes = dict(zip(CLASS_KEYS, vals))
    v = res.classes

    # --- nom del sediment ---------------------------------------------------
    silt_tag = 0
    silt = 0.0

    def pick(xs, names):
        # el primer que és >= tots els següents i > tots els anteriors
        for idx, x in enumerate(xs):
            if all(x > y for y in xs[:idx]) and all(x >= y for y in xs[idx + 1:]):
                return names[idx]
        return None

    sed = None
    gkeys = CLASS_KEYS[0:5]
    skeys = CLASS_KEYS[5:10]
    mkeys = CLASS_KEYS[10:15]
    if any(v[k] == "-" for k in gkeys) or any(v[k] == "-" for k in skeys):
        res.warnings.append("nonstandard_name")
    else:
        gname = pick([v[k] for k in gkeys], ["Very Coarse Gravel", "Coarse Gravel", "Medium Gravel",
                                             "Fine Gravel", "Very Fine Gravel"])
        sname = pick([v[k] for k in skeys], ["Very Coarse Sand", "Coarse Sand", "Medium Sand",
                                             "Fine Sand", "Very Fine Sand"])
        gdes = gname.replace("Gravel", "Gravelly")
        sdes = sname.replace("Sand", "Sandy")
        sorting = res.descriptions.get("sorting", "")
        if any(v[k] == "-" for k in mkeys) or v["clay"] == "-":
            silt_tag = 1
            mname, mdes = "Mud", "Muddy"
        else:
            silt = sum(v[k] for k in mkeys)
            clay = v["clay"]
            siltname = pick([v[k] for k in mkeys], ["Very Coarse Silt", "Coarse Silt", "Medium Silt",
                                                    "Fine Silt", "Very Fine Silt"])
            siltdes = siltname.replace("Silt", "Silty")
            silt_tag = 2
            if silt == 0 and clay > 0:
                sc = 0.01
            elif clay == 0 and silt > 0:
                sc = 10
            elif clay == 0 and silt == 0:
                sc = None
                sed = {"Gravel": gname, "Sandy Gravel": "Sandy " + gname,
                       "Gravelly Sand": gdes + " " + sname,
                       "Slightly Gravelly Sand": "Slightly " + gdes + " " + sname,
                       "Sand": sorting + " " + sname}.get(tex)
            else:
                sc = silt / clay
            if sc is not None:
                if sc >= 2:
                    mname, mdes = siltname, siltdes
                elif sc <= 0.5:
                    mname, mdes = "Clay", "Clayey"
                else:
                    mname, mdes = "Mud", "Muddy"
        if sed is None and not (silt_tag == 2 and v["clay"] == 0 and silt == 0):
            sed = {
                "Gravel": gname,
                "Sandy Gravel": "Sandy " + gname,
                "Muddy Sandy Gravel": f"{mdes} Sandy {gname}",
                "Muddy Gravel": f"{mdes} {gname}",
                "Gravelly Sand": f"{gdes} {sname}",
                "Gravelly Muddy Sand": f"{gdes} {mdes} {sname}",
                "Gravelly Mud": f"{gdes} {mname}",
                "Slightly Gravelly Sand": f"Slightly {gdes} {sname}",
                "Slightly Gravelly Muddy Sand": f"Slightly {gdes} {mdes} {sname}",
                "Slightly Gravelly Sandy Mud": f"Slightly {gdes} {sdes} {mname}",
                "Slightly Gravelly Mud": f"Slightly {gdes} {mname}",
                "Sand": f"{sorting} {sname}",
                "Muddy Sand": f"{mdes} {sname}",
                "Sandy Mud": f"{sdes} {mname}",
                "Mud": mname,
            }.get(tex)
    res.sediment_name = sed
    res.silt = silt if silt_tag == 2 else None
    res.clay = v["clay"] if silt_tag == 2 else None

    # --- triangles (Triangle / Triangle2) -----------------------------------
    res.triangle_gsm = gsm_xy(gravel, sand, mud)
    if silt_tag == 1:
        res.sand_silt_clay_group = "Unable to determine"
    else:
        clay = v["clay"] if isinstance(v["clay"], float) else 0.0
        tot = sand + silt + clay
        if tot > 0:
            res.triangle_ssc = ssc_xy(sand, silt, clay)
            s14, si14, c14 = sand * 100 / tot, silt * 100 / tot, clay * 100 / tot
            r = 0 if si14 == 0 else (10 if c14 == 0 else si14 / c14)
            if s14 >= 90:
                g = "Sand"
            elif s14 >= 50:
                g = "Silty Sand" if r >= 2 else ("Muddy Sand" if r >= 0.5 else "Clayey Sand")
            elif s14 >= 10:
                g = "Sandy Silt" if r >= 2 else ("Sandy Mud" if r >= 0.5 else "Sandy Clay")
            else:
                g = "Silt" if r >= 2 else ("Mud" if r >= 0.5 else "Clay")
            res.sand_silt_clay_group = g


# Coordenades dels diagrames triangulars (mateixa geometria que GRADISTAT:
# base 0-460, alçada 0-398.372; el vèrtex superior és a x = 230)
H = 398.372


def gsm_y(grav):
    if grav <= 0.5:
        return grav * 19.919 / 0.5
    if grav <= 5:
        return (grav - 0.5) * 39.837 / 4.5 + 19.919
    if grav <= 30:
        return (grav - 5) * 79.674 / 25 + 59.756
    if grav <= 80:
        return (grav - 30) * 179.267 / 50 + 139.43
    return (grav - 80) * 79.675 / 20 + 318.697


def gsm_x1(san):
    if san <= 10:
        return san * 6.9
    if san <= 50:
        return (san - 10) * (230 - 69) / 40 + 69
    if san <= 90:
        return (san - 50) * (391 - 230) / 40 + 230
    return (san - 90) * (460 - 391) / 10 + 391


def _tri_xy(x1, y1):
    return x1 + y1 * (230 - x1) / H, y1


def gsm_xy(gravel, sand, mud):
    grav = gravel / (gravel + sand + mud) * 100 if gravel > 0 else 0.0
    san = sand / (sand + mud) * 100 if sand > 0 else 0.0
    return _tri_xy(gsm_x1(san), gsm_y(grav))


def ssc_xy(sand, silt, clay):
    s13 = sand / (sand + silt + clay) * 100 if sand > 0 else 0.0
    si13 = silt / (silt + clay) * 100 if silt > 0 else 0.0
    return _tri_xy(si13 / 100 * 460, s13 / 100 * H)


# ---------------------------------------------------------------------------
# Punt d'entrada
# ---------------------------------------------------------------------------
def analyse(apertures_um, weights, pan_weight=0.0, initial_weight=None) -> Result:
    res = Result()
    A, G = build_rows(apertures_um, weights, pan_weight)
    if np.sum(G) <= 0:
        raise ValueError("No hi ha cap pes introduït.")
    t, moments = calculations_sheet(A, G)
    res.table, res.moments = t, moments
    if t["F"][0] > 0:
        res.warnings.append("top_sieve")
    try:
        folk_ward(t, res)
    except (ZeroDivisionError, ValueError, OverflowError):
        res.warnings.append("fatal")
        return res
    try:
        folk54(t, res)
    except _FolkError:
        res.warnings.append("nonstandard")
    except (ZeroDivisionError, ValueError, TypeError, OverflowError):
        res.warnings.append("fatal")
    if initial_weight:
        res.sieving_error = ((initial_weight / t["total"]) * 100 - 100) / 100
    res.warnings = list(dict.fromkeys(res.warnings))
    return res
