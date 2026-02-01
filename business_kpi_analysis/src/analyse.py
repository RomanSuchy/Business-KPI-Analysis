import pandas as pd
import matplotlib.pyplot as plt

df = pd.read_excel("../data/kundendaten.xlsx")
print(df.head())

anzahl_kunden = df["Kunde_ID"].nunique()
durchschnitt_beitrag = df["Jahresbeitrag"].mean()
gesamtumsatz = df["Jahresbeitrag"].sum()
schadenquote = (df["Schaden"] == "Ja").mean() * 100

print("Anzahl Kunden:", anzahl_kunden)
print("Durchschnittlicher Beitrag:", round(durchschnitt_beitrag,2))
print("Gesamtumsatz:", gesamtumsatz)
print("Schadenquote:", round(schadenquote,2), "%")

umsatz_nach_vertrag = df.groupby("Vertragstyp")["Jahresbeitrag"].sum()
print(umsatz_nach_vertrag)

umsatz_nach_vertrag.plot(kind="bar")
plt.title("Umsatz nach Vertragstyp")
plt.xlabel("Vertragstyp")
plt.ylabel("Umsatz in €")
plt.tight_layout()
plt.show()

kpi_df = pd.DataFrame({
    "Kennzahl": [
        "Anzahl Kunden",
        "Durschnitt Beitrag",
        "Gesamtumsatz",
        "Schadenquote (%)"
    ],
    "Wert": [
        anzahl_kunden,
        round(durchschnitt_beitrag),
        gesamtumsatz,
        round(schadenquote, 2)
    ]
})

kpi_df.to_excel("../output/kpi_report.xlsx", index=False)