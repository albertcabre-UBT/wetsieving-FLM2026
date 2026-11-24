# GRADISTAT web (mostra única)

App de Streamlit que reprodueix els càlculs de **GRADISTAT v9.1** (Blott & Pye, 2001) per analitzar mostres d'una en una, de tamisatge o de granulòmetre làser.

## Com executar-la

```bash
pip install -r requirements.txt
streamlit run app.py
```

S'obre al navegador a `http://localhost:8501`.

Per publicar-la, puja aquesta carpeta a un repositori de GitHub i desplega-la a [Streamlit Community Cloud](https://streamlit.io/cloud): el fitxer principal és `app.py`.

## Ús

1. Omple l'identificador, l'analista, la data i, si vols, el pes inicial (per calcular l'error de tamisatge).
2. Tria una plantilla d'obertures (½ φ, ¼ φ o làser LS230) i escriu els pesos retinguts a la taula. També pots pujar un CSV o Excel.
3. Escriu el pes del **pan** (el material que ha passat per l'última obertura) al camp de sota la taula.
4. Els resultats s'actualitzen sols. Pots descarregar cada mostra en Excel, o afegir-la al llistat de la sessió i descarregar-lo tot junt al final.

### Format del fitxer

Dues columnes: obertura en µm i pes retingut. Pot tenir capçalera i fer servir `,` o `.` com a separador decimal. La fila del pan pot tenir obertura `0` o el text `pan`. Hi ha un exemple a `exemple_mostra.csv`.

## Què calcula

Fa el mateix que el full *Single Sample Statistics* de l'Excel:

- **Mètode dels moments:** aritmètic (µm), geomètric (µm) i logarítmic (φ).
- **Folk & Ward (1957):** geomètric (µm) i logarítmic (φ), amb les descripcions.
- **Valors de la distribució:** modes (fins a 3), D10, D50, D90, D90/D10, D90−D10, D75/D25 i D75−D25.
- **Composició:** grava, sorra i fang, i les 16 classes d'Udden–Wentworth.
- **Classificació:** grup textural i nom del sediment (Folk, 1954), i tipus de mostra.
- **Gràfics:** histograma amb la corba acumulada, i els triangles grava–sorra–fang i sorra–llim–argila.

## Validació

`gradistat.py` és un port fidel de les fórmules del full *Calculations* i de les macros `FolkAndWardMacro`, `Folk54Macro`, `Triangle` i `Triangle2`. Fins i tot reprodueix les peculiaritats de l'original. Per exemple, la columna µm de D10 correspon al 10 % més fi, i el pan s'interpola a 10 φ.

S'ha comprovat amb 20 mostres (tamisos ½ φ, ¼ φ i làser; sorres, graves, fangs, mostres bimodals i mostres amb molt de pan) executant la macro original. Els 1.267 valors coincideixen fins a 9 decimals. Per tornar-ho a comprovar:

```bash
python -m pytest tests
```

## Referència

Blott, S.J. & Pye, K. (2001). GRADISTAT: a grain size distribution and statistics package for the analysis of unconsolidated sediments. *Earth Surface Processes and Landforms* 26, 1237–1248.

GRADISTAT és obra de Simon J. Blott (Kenneth Pye Associates Ltd.). Aquesta app reimplementa els seus mètodes; si la feu pública, citeu l'article.
