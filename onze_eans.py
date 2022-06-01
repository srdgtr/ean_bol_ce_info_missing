# verkrijgen van eans met missende ce_ handelmarkering, die in ons assortiment voorkomen

import pandas as pd
import os
from pathlib import Path
basis_file_path = Path.home() / "Dropbox" / "MACRO" / "Basisbestanden"


laatste_basisbestand = (
    pd.read_excel(
        max(
            basis_file_path.glob("**/BasisBestand*.xlsm"),
            key=os.path.getmtime,
        ),
        converters={
            "EAN (handmatig)": lambda x: pd.to_numeric(x, errors="coerce"),
            "EAN": lambda x: pd.to_numeric(x, errors="coerce"),
        },
        usecols=[
            "Product ID eigen",
            "Product ID eigen (nieuw)",
            "United actie",
            "EAN",
            "EAN (handmatig)",

        ],
        engine="openpyxl",
    )
    .query("`Product ID eigen`.notnull()")
    .assign(
        Product_ID_eigen=lambda x: x["United actie"].fillna(x["Product ID eigen (nieuw)"]).fillna(x["Product ID eigen"]),
        ean=lambda x: x["EAN (handmatig)"].fillna(x["EAN"]).fillna("0").astype("int64"),
    )
)
nodige_colums_basis = laatste_basisbestand[["Product_ID_eigen","ean"]]

ean_ce_prob = pd.read_excel(
        max(
            Path.cwd().glob("**/EANs-P*.xlsx"),
            key=os.path.getmtime,
        ),engine="openpyxl",).rename(columns={"EAN artikelen":"ean"})

welke_producten = nodige_colums_basis.merge(ean_ce_prob, on="ean",how="left").dropna(subset=["Nog te vullen attributen"])

welke_producten.to_excel("onze_ce_prod.xlsx", index=False)