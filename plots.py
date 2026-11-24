"""Gràfics (Plotly) per a l'app GRADISTAT."""
from __future__ import annotations

import numpy as np
import plotly.graph_objects as go

from gradistat import H, gsm_x1, gsm_y, NROWS

INK = "#1f2933"
MUTED = "#7b8794"
GRID = "#e4e7eb"
BAR = "#2f6f9f"
LINE = "#c2410c"
POINT = "#c2410c"


def distribution_chart(res, title=""):
    t = res.table
    A, F = t["A"], t["F"]
    kp = t["kpan"]
    lo, hi, pct = [], [], []
    for k in range(1, min(kp, NROWS)):
        lo.append(A[k] * 1000)
        hi.append(A[k - 1] * 1000)
        pct.append(F[k])
    # el pan es dibuixa a part, amb l'amplada (log) de l'última classe
    last = A[kp - 1] * 1000
    ratio = (A[kp - 2] / A[kp - 1]) if kp >= 2 else 2.0
    pan_lo, pan_pct = last / ratio, F[kp] if kp < NROWS else 0.0
    lo, hi, pct = np.array(lo), np.array(hi), np.array(pct)
    centers = np.sqrt(lo * hi)
    widths = hi - lo

    # % acumulat més fi que cada obertura
    finer = 100 - t["E"][: kp]
    sizes = A[: kp] * 1000

    fig = go.Figure()
    fig.add_bar(x=centers, y=pct, width=widths, name="Pes de la classe (%)",
                marker=dict(color=BAR, line=dict(color="white", width=0.5)),
                customdata=np.stack([lo, hi], axis=1),
                hovertemplate="%{customdata[0]:.4g}–%{customdata[1]:.4g} µm<br>%{y:.2f} %<extra></extra>")
    if pan_pct > 0:
        fig.add_bar(x=[np.sqrt(pan_lo * last)], y=[pan_pct], width=[last - pan_lo], name="Pan",
                    marker=dict(color="#9aa5b1", line=dict(color="white", width=0.5)),
                    hovertemplate=f"Pan (< {last:.4g} µm)<br>%{{y:.2f}} %<extra></extra>")
    fig.add_scatter(x=sizes, y=finer, mode="lines+markers", name="% acumulat més fi",
                    yaxis="y2", line=dict(color=LINE, width=2), marker=dict(size=5),
                    hovertemplate="%{x:.4g} µm<br>%{y:.1f} % més fi<extra></extra>")
    fig.update_layout(
        title=dict(text=title, x=0, font=dict(size=14, color=INK)),
        height=420, margin=dict(l=60, r=60, t=50, b=50), plot_bgcolor="white",
        legend=dict(orientation="h", y=-0.18, x=0),
        bargap=0,
        xaxis=dict(type="log", title="Mida de partícula (µm)", gridcolor=GRID,
                   range=[np.log10(min(lo.min(), pan_lo)) - 0.05, np.log10(hi.max()) + 0.05]),
        yaxis=dict(title="Pes de la classe (%)", gridcolor=GRID, rangemode="tozero"),
        yaxis2=dict(title="% acumulat més fi", overlaying="y", side="right", range=[0, 102],
                    tickvals=[0, 25, 50, 75, 100], showgrid=False),
    )
    return fig


def _line(fig, xs, ys, **kw):
    fig.add_scatter(x=xs, y=ys, mode="lines", hoverinfo="skip", showlegend=False,
                    line=dict(color=kw.get("color", INK), width=kw.get("width", 1)))


def _ratio_x(x1, y):
    return x1 + y * (230 - x1) / H


def _label(fig, x, y, text, size=11, color=INK):
    fig.add_annotation(x=x, y=y, text=text, showarrow=False, font=dict(size=size, color=color))


def _base_layout(fig, title):
    fig.update_layout(
        title=dict(text=title, x=0, font=dict(size=14, color=INK)),
        height=540, margin=dict(l=10, r=10, t=40, b=10), plot_bgcolor="white",
        xaxis=dict(visible=False, range=[-30, 500], fixedrange=True),
        yaxis=dict(visible=False, range=[-42, H + 25], scaleanchor="x", scaleratio=1,
                   fixedrange=True),
        showlegend=False,
    )


def gsm_triangle(res, title="Gravel–Sand–Mud (Folk, 1954)"):
    fig = go.Figure()
    _line(fig, [0, 460, 230, 0], [0, 0, H, 0], width=1.6)
    # nivells de grava: trace, 5, 30, 80 %
    for g, lab in [(0.5, "Trace"), (5, "5%"), (30, "30%"), (80, "80%")]:
        y = gsm_y(g)
        xl, xr = y * 230 / H, 460 - y * 230 / H
        _line(fig, [xl, xr], [y, y])
        _label(fig, xr + 22, y, lab, size=10, color=MUTED)
    # proporcions sorra:fang
    for san, top in [(10, 5), (50, 80), (90, 80)]:
        x1 = gsm_x1(san)
        yt = gsm_y(top)
        _line(fig, [x1, _ratio_x(x1, yt)], [0, yt])
    # etiquetes de camp
    bands = {"top": (80, 100), "hi": (30, 80), "mid": (5, 30), "low": (0.5, 5), "zero": (0, 0.5)}

    def put(band, s0, s1, text):
        g0, g1 = bands[band]
        y = (gsm_y(g0) + gsm_y(g1)) / 2
        x1 = (gsm_x1(s0) + gsm_x1(s1)) / 2
        _label(fig, _ratio_x(x1, y), y, text, size=10)

    _label(fig, 230, (gsm_y(80) + H) / 2 - 10, "G", size=11)
    put("hi", 0, 50, "mG"); put("hi", 50, 90, "msG"); put("hi", 90, 100, "sG")
    put("mid", 0, 50, "gM"); put("mid", 50, 90, "gmS"); put("mid", 90, 100, "gS")
    put("low", 0, 10, "(g)M"); put("low", 10, 50, "(g)sM"); put("low", 50, 90, "(g)mS")
    put("low", 90, 100, "(g)S")
    put("zero", 0, 10, "M"); put("zero", 10, 50, "sM"); put("zero", 50, 90, "mS")
    put("zero", 90, 100, "S")
    _label(fig, 230, H + 15, "<b>Grava</b>")
    _label(fig, -5, -22, "<b>Fang</b>")
    _label(fig, 465, -22, "<b>Sorra</b>")
    for r, lab in [(10, "1:9"), (50, "1:1"), (90, "9:1")]:
        _label(fig, gsm_x1(r), -12, lab, size=9, color=MUTED)
    _label(fig, 230, -32, "Proporció sorra:fang", size=9, color=MUTED)
    if res.triangle_gsm:
        x, y = res.triangle_gsm
        fig.add_scatter(x=[x], y=[y], mode="markers",
                        marker=dict(size=12, color=POINT, line=dict(color="white", width=1.5)),
                        hovertemplate=(f"Grava {res.gravel:.1f} %<br>Sorra {res.sand:.1f} %"
                                       f"<br>Fang {res.mud:.1f} %<extra></extra>"))
    _base_layout(fig, title)
    return fig


def ssc_triangle(res, title="Sand–Silt–Clay (Folk, 1954)"):
    fig = go.Figure()
    _line(fig, [0, 460, 230, 0], [0, 0, H, 0], width=1.6)
    for s in (10, 50, 90):
        y = s / 100 * H
        xl, xr = y * 230 / H, 460 - y * 230 / H
        _line(fig, [xl, xr], [y, y])
        _label(fig, xr + 22, y, f"{s}%", size=10, color=MUTED)
    yt = 0.9 * H
    for frac in (1 / 3, 2 / 3):
        x1 = frac * 460
        _line(fig, [x1, _ratio_x(x1, yt)], [0, yt])

    def put(s0, s1, f0, f1, text):
        y = (s0 + s1) / 200 * H
        x1 = (f0 + f1) / 2 * 460
        _label(fig, _ratio_x(x1, y), y, text, size=10)

    _label(fig, 230, (0.9 * H + H) / 2 - 8, "S", size=11)
    for (s0, s1), labs in {(50, 90): ("cS", "mS", "zS"), (10, 50): ("sC", "sM", "sZ"),
                           (0, 10): ("C", "M", "Z")}.items():
        for (f0, f1), lab in zip([(0, 1 / 3), (1 / 3, 2 / 3), (2 / 3, 1)], labs):
            put(s0, s1, f0, f1, lab)
    _label(fig, 230, H + 15, "<b>Sorra</b>")
    _label(fig, -5, -22, "<b>Argila</b>")
    _label(fig, 465, -22, "<b>Llim</b>")
    _label(fig, 460 / 3, -12, "1:2", size=9, color=MUTED)
    _label(fig, 2 * 460 / 3, -12, "2:1", size=9, color=MUTED)
    _label(fig, 230, -32, "Proporció llim:argila", size=9, color=MUTED)
    if res.triangle_ssc:
        x, y = res.triangle_ssc
        fig.add_scatter(x=[x], y=[y], mode="markers",
                        marker=dict(size=12, color=POINT, line=dict(color="white", width=1.5)),
                        hovertemplate=(f"Sorra {res.sand:.1f} %<br>Llim {res.silt:.1f} %"
                                       f"<br>Argila {res.clay:.1f} %<extra></extra>"))
    _base_layout(fig, title)
    return fig
