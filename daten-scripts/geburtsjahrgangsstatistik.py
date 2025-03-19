import pandas as pd
import os
import datetime

from common.gebiet_schluessel import gebiet_schluessel
from common.sorted_gebiet_to_gemeinde import sorted_gebiet_to_gemeinde

FILENAME = "geburtsjahrgangsstatistik.xlsx"
INPUT_DIR = "data"
OUTPUT_DIR = "dist"
SHEET_NAME = "dadigesamt"

INPUT_FILENAME = os.path.join(INPUT_DIR, FILENAME)
OUTPUT_FILENAME = os.path.join(OUTPUT_DIR, FILENAME)

def parse_excel():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if not os.path.exists(INPUT_FILENAME):
        print(f"Error: File {INPUT_FILENAME} not found.")
        return

    xls = pd.ExcelFile(INPUT_FILENAME)
    if SHEET_NAME not in xls.sheet_names:
        print(f"Error: Sheet '{SHEET_NAME}' not found in the file. Available sheets: {xls.sheet_names}")
        return

    df = pd.read_excel(xls, sheet_name=SHEET_NAME, dtype=str)
    df = df.dropna(subset=["Gebiet", "Jahrgang"])
    df["Gebiet"] = df["Gebiet"].str.strip()

    df["Jahrgang"] = pd.to_numeric(df["Jahrgang"], errors='coerce')
    df = df.dropna(subset=["Jahrgang"])
    df["Jahrgang"] = df["Jahrgang"].astype(int)

    df["Gemeinde"] = df["Gebiet"].map(lambda x: sorted_gebiet_to_gemeinde.get(x, x))

    columns_to_keep = ["Gemeinde", "Jahrgang", "M gesamt", "W gesamt", "EW gesamt"]
    missing_columns = [col for col in columns_to_keep if col not in df.columns]
    if missing_columns:
        print(f"Error: Missing columns {missing_columns}")
        return

    df = df[columns_to_keep]

    for col in ["M gesamt", "W gesamt", "EW gesamt"]:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

    current_year = datetime.datetime.now().year
    def classify_age_group(year):
        age = current_year - year
        if age < 20:
            return "Unter 20 Jährige"
        elif age > 64:
            return "65 Jahre und älter"
        return "20 Jahre - 64 Jahre"

    df["Gruppe"] = df["Jahrgang"].apply(classify_age_group)

    grouped_df = df.groupby(["Gemeinde", "Gruppe"])[["M gesamt", "W gesamt", "EW gesamt"]].sum().reset_index()

    total_population = df.groupby("Gemeinde")[["M gesamt", "W gesamt", "EW gesamt"]].sum().reset_index()
    total_population = total_population.rename(columns={
        "M gesamt": "M total", "W gesamt": "W total", "EW gesamt": "EW total"
    })

    grouped_df = grouped_df.merge(total_population, on="Gemeinde", how="left")

    grouped_df.insert(1, "Amtlicher Gemeinde-schlüssel", grouped_df["Gemeinde"].map(lambda x: gebiet_schluessel.get(x, ("", ""))[0]))
    grouped_df.insert(2, "ISO", grouped_df["Gemeinde"].map(lambda x: gebiet_schluessel.get(x, ("", ""))[1]))

    grouped_df["M quotient"] = ((grouped_df["M gesamt"] / grouped_df["M total"]) * 100).round(2).astype(str) + "%"
    grouped_df["W quotient"] = ((grouped_df["W gesamt"] / grouped_df["W total"]) * 100).round(2).astype(str) + "%"
    grouped_df["EW quotient"] = ((grouped_df["EW gesamt"] / grouped_df["EW total"]) * 100).round(2).astype(str) + "%"

    grouped_df = grouped_df[~grouped_df["Gemeinde"].isin(["Ausgewählte Gebiete zusammengefasst", "Sanierungsgebiet"])]

    grouped_df.to_excel(OUTPUT_FILENAME, index=False)
    print(f"Result saved to {OUTPUT_FILENAME}")

if __name__ == "__main__":
    parse_excel()
