"""CSV till Excel – läser försäljningsfiler, rensar data och skapar output/rapport.xlsx.

Kör:  .venv\\Scripts\\python.exe "<sökväg>\\rapport.py"
Originalfilerna i data/ ändras aldrig. Programmet kan köras om; rapporten skrivs över.
"""
import logging
import sys
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

KOLUMNER = ["order_id", "date", "product", "quantity", "unit_price_sek"]
PROJEKT = Path(__file__).resolve().parent
DATA = PROJEKT / "data"
OUTPUT = PROJEKT / "output"


def starta_loggning() -> logging.Logger:
    OUTPUT.mkdir(exist_ok=True)
    log = logging.getLogger("rapport")
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fmt = logging.Formatter("%(asctime)s  %(levelname)s  %(message)s", "%Y-%m-%d %H:%M:%S")
    for h in (logging.StreamHandler(sys.stdout), logging.FileHandler(OUTPUT / "rapport.log", mode="w", encoding="utf-8")):
        h.setFormatter(fmt)
        log.addHandler(h)
    return log


def las_filer(log: logging.Logger) -> pd.DataFrame:
    # Filerna läses i namnordning, så "första förekomsten" av en dubblett blir alltid densamma.
    # Döp filerna så att namnordning = tidsordning, t.ex. forsaljning_2026-01.csv (inte _jan/_feb).
    filer = sorted(DATA.glob("*.csv"))
    if not filer:
        raise SystemExit(f"Inga CSV-filer hittades i {DATA}")
    delar = []
    for f in filer:
        # Läs allt som text: då jämförs dubbletter exakt som de står i filen och inget gissas
        df = pd.read_csv(f, dtype=str, keep_default_na=False)
        saknas = [k for k in KOLUMNER if k not in df.columns]
        if saknas:
            raise SystemExit(f"{f.name} saknar kolumner: {saknas}")
        df = df[KOLUMNER].apply(lambda s: s.str.strip())
        df["kallfil"] = f.name
        df["rad_i_fil"] = range(2, len(df) + 2)  # radnummer som i filen (rad 1 = rubriker)
        log.info("Läste %s: %d rader", f.name, len(df))
        delar.append(df)
    return pd.concat(delar, ignore_index=True)


def rensa(data: pd.DataFrame, log: logging.Logger):
    # 1) Exakta dubbletter över alla fem kolumner, även mellan filer. Första förekomsten behålls.
    dubbl = data.duplicated(subset=KOLUMNER, keep="first")
    problem_dubbl = data[dubbl].assign(orsak="Exakt dubblett av tidigare rad")
    kvar = data[~dubbl]

    # 2) Rader som saknar någon uppgift flaggas och utesluts – inga värden gissas.
    tom = kvar[KOLUMNER].eq("")
    problem_saknas = kvar[tom.any(axis=1)].copy()
    problem_saknas["orsak"] = tom[tom.any(axis=1)].apply(
        lambda r: "Saknar: " + ", ".join(k for k in KOLUMNER if r[k]), axis=1)
    rensad = kvar[~tom.any(axis=1)].copy()

    problem = pd.concat([problem_dubbl, problem_saknas], ignore_index=True)
    for _, r in problem.iterrows():
        log.info("Utesluten: %s rad %d (order %s) – %s", r.kallfil, r.rad_i_fil, r.order_id or "?", r.orsak)

    # 3) Typer och försäljning (indata är giltig enligt specifikationen, men felet ska synas om den inte är det)
    rensad["date"] = pd.to_datetime(rensad["date"], format="%Y-%m-%d")
    rensad["quantity"] = pd.to_numeric(rensad["quantity"])
    rensad["unit_price_sek"] = pd.to_numeric(rensad["unit_price_sek"])
    rensad["forsaljning_sek"] = (rensad["quantity"] * rensad["unit_price_sek"]).round(2)
    rensad["manad"] = rensad["date"].dt.strftime("%Y-%m")

    log.info("Inlästa rader: %d", len(data))
    log.info("Borttagna dubbletter: %d", len(problem_dubbl))
    log.info("Rader med saknade värden: %d", len(problem_saknas))
    log.info("Godkända rader: %d", len(rensad))
    rensad = rensad.sort_values(["date", "order_id"], kind="stable").reset_index(drop=True)
    return rensad, problem


def sammanfatta(rensad: pd.DataFrame):
    per_manad = rensad.pivot_table(index="manad", columns="product", values="forsaljning_sek",
                                   aggfunc="sum", fill_value=0).round(2)
    per_manad["Totalt"] = per_manad.sum(axis=1).round(2)
    per_manad = per_manad.reset_index().rename(columns={"manad": "Månad"})
    per_manad.columns.name = None

    per_produkt = (rensad.groupby("product")
                   .agg(antal_salda=("quantity", "sum"), forsaljning_sek=("forsaljning_sek", "sum"))
                   .round(2).sort_values("forsaljning_sek", ascending=False).reset_index())
    return per_manad, per_produkt


def tal_dar_mojligt(df: pd.DataFrame) -> pd.DataFrame:
    """Gör textkolumner som bara innehåller siffror (eller tomt) till tal, så att Excel inte varnar
    för "tal lagrat som text". Kolumner med annan text lämnas orörda."""
    df = df.copy()
    for kol in df.columns:
        if not (pd.api.types.is_object_dtype(df[kol]) or pd.api.types.is_string_dtype(df[kol])):
            continue
        s = df[kol].astype(object).replace("", None)
        try:
            df[kol] = pd.to_numeric(s)
        except (ValueError, TypeError):
            pass
    return df


def skriv_excel(rensad, per_manad, per_produkt, problem, log):
    fil = OUTPUT / "rapport.xlsx"
    blad = {
        "Rensad_data": rensad[["order_id", "date", "product", "quantity", "unit_price_sek", "forsaljning_sek", "kallfil"]],
        "Per_manad": per_manad,
        "Per_produkt": per_produkt,
        "Dataproblem": problem[["kallfil", "rad_i_fil"] + KOLUMNER + ["orsak"]],
    }
    try:
        with pd.ExcelWriter(fil, engine="openpyxl", datetime_format="YYYY-MM-DD") as xw:
            for namn, df in blad.items():
                df = tal_dar_mojligt(df)
                df.to_excel(xw, sheet_name=namn, index=False)
                ws = xw.sheets[namn]
                ws.freeze_panes = "A2"
                for i, kol in enumerate(df.columns, 1):   # datum utan klockslag
                    if pd.api.types.is_datetime64_any_dtype(df[kol]):
                        for cell in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                            cell[0].number_format = "YYYY-MM-DD"
                for i, kol in enumerate(df.columns, 1):
                    bredd = max([len(str(kol)), 10] + [len(str(v)) for v in df[kol].head(200)])
                    ws.column_dimensions[get_column_letter(i)].width = min(bredd + 3, 45)

            # Diagram över total försäljning per månad
            ws = xw.sheets["Per_manad"]
            tot_kol = list(per_manad.columns).index("Totalt") + 1
            n = len(per_manad)
            diagram = BarChart()
            diagram.title = "Försäljning per månad (SEK)"
            diagram.legend = None
            diagram.varyColors = False             # samma färg på alla staplar
            diagram.y_axis.delete = False          # nyare Excel döljer annars axlarna
            diagram.x_axis.delete = False
            diagram.y_axis.scaling.min = 0         # staplar ska börja på noll
            diagram.y_axis.numFmt = "# ##0"
            diagram.y_axis.majorGridlines = None
            diagram.add_data(Reference(ws, min_col=tot_kol, min_row=1, max_row=n + 1), titles_from_data=True)
            diagram.set_categories(Reference(ws, min_col=1, min_row=2, max_row=n + 1))
            diagram.width, diagram.height = 18, 9
            ws.add_chart(diagram, f"{get_column_letter(tot_kol + 2)}2")
    except PermissionError:
        raise SystemExit(f"Kan inte skriva {fil} – stäng filen i Excel och kör igen.")
    log.info("Rapport skapad: %s", fil)


def main():
    log = starta_loggning()
    data = las_filer(log)
    rensad, problem = rensa(data, log)
    per_manad, per_produkt = sammanfatta(rensad)
    skriv_excel(rensad, per_manad, per_produkt, problem, log)


if __name__ == "__main__":
    main()
