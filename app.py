import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Confronto Chiamate", layout="centered")
st.title("📞 Confronto tra File di Chiamate")

st.markdown("Carica due file Excel contenenti i dati delle chiamate per confrontarli.")

# Funzioni di supporto
def estrai_numero(alias):
    match = re.search(r'\b\d{3}\b', str(alias))
    return match.group(0) if match else None

def minuti_con_secondi(durata_str):
    try:
        minuti, secondi = map(int, str(durata_str).split(":"))
        return minuti + secondi / 60
    except:
        return 0

# Upload file
file1 = st.file_uploader("Carica il file del centralino", type=[".xlsx"])
file2 = st.file_uploader("Carica il file del gestionale", type=[".xlsx"])

if file1 is not None and file2 is not None:
    try:
        # File del centralino
        df1 = pd.read_excel(file1, header=0)
        df1 = df1.rename(columns={"Operatore": "alias", "Chiamate Risposte": "chiamate", "Minuti": "minuti"})
        df1["numero"] = df1["alias"].apply(estrai_numero)
        df1 = df1[["numero", "chiamate", "minuti"]]
        df1["chiamate"] = pd.to_numeric(df1["chiamate"], errors="coerce")
        df1["minuti"] = pd.to_numeric(df1["minuti"], errors="coerce")

        # File del gestionale
        df2 = pd.read_excel(file2, skiprows=2)
        df2.columns = ["alias", "chiamate", "durata", "media_durata"]
        df2["numero"] = df2["alias"].apply(estrai_numero)
        df2["minuti"] = df2["durata"].apply(minuti_con_secondi)
        df2 = df2[["numero", "chiamate", "minuti"]]

        # Confronto
        confronto = pd.merge(df1, df2, on="numero", how="outer", suffixes=("_file1", "_file2"))
        confronto["diff_chiamate"] = confronto["chiamate_file1"] - confronto["chiamate_file2"]
        confronto["diff_minuti"] = confronto["minuti_file1"] - confronto["minuti_file2"]

        st.success("Confronto completato! Scarica il file risultante qui sotto.")

        # Download del file risultato
        output = io.BytesIO()
        confronto.to_excel(output, index=False)
        st.download_button(
            label="📂 Scarica il file di confronto",
            data=output.getvalue(),
            file_name="confronto_chiamate_minuti.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Mostra anteprima tabella
        st.subheader("Anteprima del risultato")
        st.dataframe(confronto.head(10))

    except Exception as e:
        st.error(f"Errore durante l'elaborazione: {e}")
