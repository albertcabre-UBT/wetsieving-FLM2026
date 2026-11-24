"""
GRADISTAT web · mostra única
Streamlit app basada en GRADISTAT v9.1 (Blott & Pye, 2001).

Executa:  streamlit run app.py
"""
from __future__ import annotations

import datetime as dt
import io
import math

import numpy as np
import pandas as pd
import streamlit as st

from gradistat import CLASS_KEYS, CLASS_LABELS, TEMPLATES, analyse
from plots import distribution_chart, gsm_triangle, ssc_triangle

st.set_page_config(page_title="GRADISTAT web", page_icon="⏳", layout="wide")

COL_AP = "Obertura (µm)"
COL_W = "Pes retingut"

WARNINGS = {
    "top_sieve": "Hi ha pes retingut al tamís més gran. Com a l'Excel, aquesta fracció no té punt mig "
                 "i no entra bé al mètode dels moments: afegeix un tamís més gran amb pes 0.",
    "pan": "Més del 5 % de la mostra és a la classe més fina (el pan). GRADISTAT n'ha interpolat la mida; "
           "per més precisió, analitza millor les fraccions fines.",
    "multimodal": "La mostra no és unimodal: el sorting, l'asimetria i la curtosi són poc fiables.",
    "nonstandard": "No s'han fet servir espaiats de tamís estàndard: no es poden calcular els percentatges "
                   "per classe, el grup textural, el nom del sediment ni el triangle.",
    "nonstandard_name": "No s'han fet servir espaiats de tamís estàndard: no es pot deduir el nom del sediment.",
    "fatal": "S'ha produït un error de càlcul i algunes estadístiques poden faltar. Revisa les dades.",
}


# ---------------------------------------------------------------------------
# Utilitats
# ---------------------------------------------------------------------------
def fmt(v):
    """Format de números com fa GRADISTAT (3, 2 o 1 decimals segons la mida)."""
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    v = float(v)
    if not math.isfinite(v):
        return "–"
    if v < 10:
        return f"{v:.3f}"
    if v < 100:
        return f"{v:.2f}"
    return f"{v:.1f}"


def pct(v):
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    return f"{float(v):.1f} %"


def template_df(name):
    ap = TEMPLATES[name]
    return pd.DataFrame({COL_AP: [float(a) for a in ap], COL_W: [None] * len(ap)})


def parse_upload(file):
    """Llegeix un CSV o Excel amb dues columnes: obertura (µm) i pes.
    Una fila amb obertura 0, buida o 'pan' es pren com el pan."""
    name = file.name.lower()
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        raw = pd.read_excel(file, header=None)
    else:
        raw = pd.read_csv(file, header=None, sep=None, engine="python", decimal=_guess_decimal(file))
    raw = raw.iloc[:, :2]
    raw.columns = ["ap", "w"]
    pan = 0.0
    rows = []
    for ap, w in raw.itertuples(index=False):
        wv = pd.to_numeric(str(w).replace(",", "."), errors="coerce") if not isinstance(w, (int, float)) else w
        is_pan = isinstance(ap, str) and ap.strip().lower() in ("pan", "fons", "fondo", "<")
        apv = pd.to_numeric(ap, errors="coerce") if not isinstance(ap, (int, float)) else ap
        if is_pan or (apv is not None and not pd.isna(apv) and apv == 0):
            if wv is not None and not pd.isna(wv):
                pan += float(wv)
            continue
        if apv is None or pd.isna(apv):
            continue  # capçaleres o files buides
        rows.append((float(apv), None if pd.isna(wv) else float(wv)))
    df = pd.DataFrame(rows, columns=[COL_AP, COL_W]).sort_values(COL_AP, ascending=False)
    return df.reset_index(drop=True), pan


def _guess_decimal(file):
    head = file.getvalue()[:2000].decode("utf-8", errors="ignore")
    file.seek(0)
    return "," if (";" in head and "," in head) else "."


def summary_row(res, meta):
    fw, mo = res.folk_ward or {}, res.moments
    row = {
        "Mostra": meta["id"], "Analista": meta["analyst"], "Data": meta["date"],
        "Tipus de mostra": res.sample_type, "Grup textural": res.textural_group,
        "Nom del sediment": res.sediment_name,
    }
    for i, m in enumerate(res.modes[:3], 1):
        row[f"Moda {i} (µm)"], row[f"Moda {i} (φ)"] = m
    for k, (um, phi) in res.d_values.items():
        row[f"{k} (µm)"] = um
        row[f"{k} (φ)"] = phi
    for nm, key in [("Mitjana", "mean"), ("Sorting", "sorting"), ("Asimetria", "skewness"),
                    ("Curtosi", "kurtosis")]:
        if fw:
            row[f"F&W {nm} (µm)"] = fw["geometric"][key]
            row[f"F&W {nm} (φ)"] = fw["logarithmic"][key]
            row[f"F&W {nm} descripció"] = res.descriptions[key]
        row[f"Moments arit. {nm} (µm)"] = mo["arithmetic"][key]
        row[f"Moments geom. {nm} (µm)"] = mo["geometric"][key]
        row[f"Moments log. {nm} (φ)"] = mo["logarithmic"][key]
    row["Grava (%)"], row["Sorra (%)"], row["Fang (%)"] = res.gravel, res.sand, res.mud
    for k in CLASS_KEYS:
        row[f"{CLASS_LABELS[k]} (%)"] = res.classes.get(k)
    if res.sieving_error is not None:
        row["Error de tamisatge (%)"] = res.sieving_error * 100
    return {k: (float(v) if isinstance(v, (np.floating,)) else v) for k, v in row.items()}


def to_excel(rows, data_df=None, pan=None, res=None):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(rows).to_excel(xw, sheet_name="Resultats", index=False)
        if data_df is not None:
            d = data_df.copy().astype(object)
            d.loc[len(d)] = ["Pan", pan]
            d.to_excel(xw, sheet_name="Dades", index=False)
        if res is not None:
            t = res.table
            kp = t["kpan"]
            calc = pd.DataFrame({
                "Obertura (µm)": (t["A"][: kp + 1] * 1000).astype(object),
                "Obertura (φ)": t["C"][: kp + 1],
                "Pes (%)": t["F"][: kp + 1],
                "% acumulat retingut": t["E"][: kp + 1],
                "% acumulat més fi": 100 - t["E"][: kp + 1],
            })
            calc.loc[kp, "Obertura (µm)"] = "Pan"
            calc.to_excel(xw, sheet_name="Distribució", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = max(12, min(40, len(str(col[0].value)) + 2))
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Estat
# ---------------------------------------------------------------------------
if "data" not in st.session_state:
    st.session_state.template = list(TEMPLATES)[0]
    st.session_state.data = template_df(st.session_state.template)
    st.session_state.pan = 0.0
    st.session_state.editor_key = 0
if "history" not in st.session_state:
    st.session_state.history = []

# ---------------------------------------------------------------------------
# Barra lateral
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Mostra")
    sid = st.text_input("Identificador de la mostra", value="")
    analyst = st.text_input("Analista", value="")
    date = st.date_input("Data", value=dt.date.today(), format="DD/MM/YYYY")
    init_w = st.number_input("Pes inicial de la mostra (opcional)", min_value=0.0, value=0.0,
                             step=0.1, help="Si l'indiques, es calcula l'error de tamisatge.")

    st.header("Dades")
    tpl = st.selectbox("Plantilla d'obertures", list(TEMPLATES),
                       index=list(TEMPLATES).index(st.session_state.template))
    if st.button("Carrega la plantilla (esborra els pesos)", width="stretch"):
        st.session_state.template = tpl
        st.session_state.data = template_df(tpl)
        st.session_state.pan = 0.0
        st.session_state.editor_key += 1
        st.rerun()

    up = st.file_uploader("…o puja un fitxer (CSV o Excel)", type=["csv", "txt", "xlsx", "xls"],
                          help="Dues columnes: obertura en µm i pes retingut. "
                               "La fila del pan pot tenir obertura 0 o el text «pan».")
    if up is not None and st.session_state.get("last_upload") != (up.name, up.size):
        try:
            df_up, pan_up = parse_upload(up)
            if df_up.empty:
                st.error("No he trobat dades numèriques al fitxer.")
            else:
                st.session_state.data = df_up
                st.session_state.pan = pan_up
                st.session_state.editor_key += 1
                st.session_state.last_upload = (up.name, up.size)
                st.rerun()
        except Exception as e:  # noqa: BLE001
            st.error(f"No he pogut llegir el fitxer: {e}")

    st.divider()
    st.caption("Mètodes de GRADISTAT v9.1 · Blott, S.J. & Pye, K. (2001). *GRADISTAT: a grain size "
               "distribution and statistics package for the analysis of unconsolidated sediments.* "
               "Earth Surface Processes and Landforms 26, 1237–1248.")

# ---------------------------------------------------------------------------
# Cos principal
# ---------------------------------------------------------------------------
st.title("GRADISTAT web")
st.caption("Estadístiques granulomètriques d'una mostra (tamisatge o granulòmetre làser).")

left, right = st.columns([1.05, 2.4], gap="large")

with left:
    st.subheader("Entrada de dades")
    st.caption("Obertures de més grossa a més fina. El pes és el material **retingut** a cada tamís.")
    edited = st.data_editor(
        st.session_state.data,
        key=f"editor_{st.session_state.editor_key}",
        num_rows="dynamic",
        width="stretch",
        height=min(38 * (len(st.session_state.data) + 2), 640),
        column_config={
            COL_AP: st.column_config.NumberColumn(COL_AP, min_value=0.0, format="%g", width="small"),
            COL_W: st.column_config.NumberColumn(COL_W, min_value=0.0, format="%g", width="small",
                                                 help="Grams o percentatge"),
        },
        hide_index=True,
    )
    pan = st.number_input("Pes al pan (passa per l'última obertura)", min_value=0.0,
                          value=float(st.session_state.pan), step=0.1, key=f"pan_{st.session_state.editor_key}")

aps = pd.to_numeric(edited[COL_AP], errors="coerce").tolist()
ws = pd.to_numeric(edited[COL_W], errors="coerce").fillna(0).tolist()
total = float(np.nansum(ws)) + float(pan or 0)

with right:
    if total <= 0:
        st.info("Introdueix els pesos retinguts (o puja un fitxer) per veure els resultats.")
        st.stop()

    try:
        res = analyse(aps, ws, pan_weight=pan, initial_weight=init_w or None)
    except Exception as e:  # noqa: BLE001
        st.error(f"No s'han pogut calcular les estadístiques: {e}")
        st.stop()

    meta = {"id": sid or "(sense nom)", "analyst": analyst, "date": date.strftime("%d/%m/%Y")}

    # --- capçalera ----------------------------------------------------------
    st.subheader(f"Resultats · {meta['id']}")
    c1, c2, c3 = st.columns([1.3, 1, 1])
    c1.markdown(f"**Nom del sediment**  \n{res.sediment_name or '–'}")
    c2.markdown(f"**Grup textural**  \n{res.textural_group or '–'}")
    c3.markdown(f"**Tipus de mostra**  \n{res.sample_type or '–'}")
    if res.sieving_error is not None:
        se = res.sieving_error * 100
        msg = f"Error de tamisatge: {'pèrdua' if se > 0 else 'guany'} del {abs(se):.1f} % respecte del pes inicial."
        (st.warning if abs(se) >= 2 else st.caption)(msg)
    for w in res.warnings:
        st.warning(WARNINGS.get(w, w))

    tab_res, tab_graf, tab_tri = st.tabs(["Estadístiques", "Distribució", "Triangles"])

    with tab_res:
        a, b = st.columns(2, gap="medium")
        with a:
            st.markdown("**Modes i percentils**")
            rows = []
            for i, (um, phi) in enumerate(res.modes[:3], 1):
                rows.append((f"Moda {i}", fmt(um), fmt(phi)))
            names = {"D10": "D10", "D50": "Mediana (D50)", "D90": "D90", "D90/D10": "D90 / D10",
                     "D90-D10": "D90 − D10", "D75/D25": "D75 / D25", "D75-D25": "D75 − D25"}
            for k, (um, phi) in res.d_values.items():
                rows.append((names[k], fmt(um), fmt(phi)))
            st.dataframe(pd.DataFrame(rows, columns=["", "µm", "φ"]), hide_index=True,
                         width="stretch")
        with b:
            st.markdown("**Composició**")
            comp = [("Grava", pct(res.gravel)), ("Sorra", pct(res.sand)), ("Fang", pct(res.mud))]
            comp += [(CLASS_LABELS[k], pct(res.classes.get(k))) for k in CLASS_KEYS] if res.classes else []
            st.dataframe(pd.DataFrame(comp, columns=["Fracció", "%"]), hide_index=True,
                         width="stretch", height=36 * (len(comp) + 1))

        st.markdown("**Mètode dels moments i Folk & Ward**")
        mo, fw, de = res.moments, res.folk_ward, res.descriptions
        stats = []
        for lab, key in [("Mitjana", "mean"), ("Sorting (σ)", "sorting"), ("Asimetria (Sk)", "skewness"),
                         ("Curtosi (K)", "kurtosis")]:
            stats.append({
                "": lab,
                "Moments · aritmètic (µm)": fmt(mo["arithmetic"][key]),
                "Moments · geomètric (µm)": fmt(mo["geometric"][key]),
                "Moments · logarítmic (φ)": fmt(mo["logarithmic"][key]),
                "F&W · geomètric (µm)": fmt(fw["geometric"][key]) if fw else "",
                "F&W · logarítmic (φ)": fmt(fw["logarithmic"][key]) if fw else "",
                "Descripció": de.get(key, ""),
            })
        st.dataframe(pd.DataFrame(stats), hide_index=True, width="stretch")

    with tab_graf:
        st.plotly_chart(distribution_chart(res, f"Distribució granulomètrica · {meta['id']}"),
                        width="stretch")

    with tab_tri:
        tri_choice = st.radio("Diagrama", ["Grava–sorra–fang", "Sorra–llim–argila"], horizontal=True,
                              label_visibility="collapsed")
        if tri_choice == "Grava–sorra–fang":
            if res.triangle_gsm:
                st.plotly_chart(gsm_triangle(res), width="stretch")
            else:
                st.info("No es pot situar la mostra al triangle grava-sorra-fang.")
        elif res.triangle_ssc:
            st.plotly_chart(ssc_triangle(res), width="stretch")
            st.caption(f"Classificació sorra-llim-argila: **{res.sand_silt_clay_group}**")
        else:
            st.info("No es pot situar la mostra al triangle sorra-llim-argila "
                    "(falten dades de les fraccions fines).")

    # --- exportació i llistat ----------------------------------------------
    st.divider()
    row = summary_row(res, meta)
    d1, d2 = st.columns(2)
    d1.download_button("⬇ Descarrega aquesta mostra (Excel)",
                       data=to_excel([row], edited, pan, res),
                       file_name=f"GRADISTAT_{meta['id']}.xlsx".replace(" ", "_"),
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       width="stretch")
    if d2.button("➕ Afegeix al llistat de la sessió", width="stretch"):
        st.session_state.history = [h for h in st.session_state.history if h["Mostra"] != row["Mostra"]]
        st.session_state.history.append(row)
        st.toast(f"Mostra {row['Mostra']} afegida al llistat")

if st.session_state.history:
    st.divider()
    st.subheader(f"Mostres analitzades en aquesta sessió ({len(st.session_state.history)})")
    hist = pd.DataFrame(st.session_state.history)
    show = ["Mostra", "Nom del sediment", "Grup textural", "Tipus de mostra", "D50 (µm)",
            "F&W Mitjana (φ)", "F&W Sorting (φ)", "Grava (%)", "Sorra (%)", "Fang (%)"]
    st.dataframe(hist[[c for c in show if c in hist]].round(3), hide_index=True, width="stretch")
    h1, h2, _ = st.columns([1, 1, 2])
    h1.download_button("⬇ Descarrega el llistat (Excel)", data=to_excel(st.session_state.history),
                       file_name="GRADISTAT_llistat.xlsx", width="stretch",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if h2.button("Buida el llistat", width="stretch"):
        st.session_state.history = []
        st.rerun()
    st.caption("El llistat només es guarda mentre la sessió està oberta: descarrega'l abans de tancar.")
