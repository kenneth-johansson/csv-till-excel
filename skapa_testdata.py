"""Skapar tre CSV-filer med påhittad försäljningsdata (30 rader: 3 dubbletter, 3 rader med saknade värden)."""
from pathlib import Path

rader = {
    "forsaljning_2026-01.csv": [
        "1001,2026-01-05,Kaffe,3,49.90",
        "1002,2026-01-07,Te,2,39.00",
        "1003,2026-01-12,Kanelbulle,10,25.00",
        "1004,2026-01-15,Kaffe,1,49.90",
        "1002,2026-01-07,Te,2,39.00",          # dubblett inom filen
        "1005,2026-01-20,Smörgås,4,65.00",
        "1006,,Te,1,39.00",                     # saknar datum
        "1007,2026-01-25,Kaffe,5,49.90",
        "1008,2026-01-28,Kanelbulle,6,25.00",
        "1009,2026-01-30,Smörgås,2,65.00",
    ],
    "forsaljning_2026-02.csv": [
        "1010,2026-02-02,Kaffe,4,49.90",
        "1011,2026-02-04,Te,3,39.00",
        "1009,2026-01-30,Smörgås,2,65.00",      # dubblett mellan filer
        "1012,2026-02-10,Kanelbulle,8,25.00",
        "1013,2026-02-12,,2,49.90",             # saknar produkt
        "1014,2026-02-15,Smörgås,3,65.00",
        "1015,2026-02-18,Kaffe,2,49.90",
        "1016,2026-02-21,Te,5,39.00",
        "1017,2026-02-24,Kanelbulle,12,25.00",
        "1018,2026-02-27,Kaffe,6,49.90",
    ],
    "forsaljning_2026-03.csv": [
        "1019,2026-03-03,Smörgås,5,65.00",
        "1020,2026-03-05,Kaffe,3,49.90",
        "1021,2026-03-09,Te,2,",                # saknar pris
        "1022,2026-03-12,Kanelbulle,9,25.00",
        "1023,2026-03-16,Kaffe,7,49.90",
        "1017,2026-02-24,Kanelbulle,12,25.00",  # dubblett mellan filer
        "1024,2026-03-20,Smörgås,1,65.00",
        "1025,2026-03-23,Te,4,39.00",
        "1026,2026-03-26,Kaffe,2,49.90",
        "1027,2026-03-30,Kanelbulle,5,25.00",
    ],
}

mapp = Path(__file__).resolve().parent / "data"
mapp.mkdir(exist_ok=True)
for namn, r in rader.items():
    (mapp / namn).write_text("order_id,date,product,quantity,unit_price_sek\n" + "\n".join(r) + "\n", encoding="utf-8")
print("Skapade", len(rader), "filer i", mapp)
