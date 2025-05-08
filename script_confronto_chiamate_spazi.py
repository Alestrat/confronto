
import pandas as pd
import re

# === CONFIGURAZIONE ===
file1_path = "Controllo Mensile Aprile.xlsx"     # File 1: centralino
file2_path = "Estrazione Massiva Chiamate.xlsx"  # File 2: gestionale

# === FUNZIONI DI SUPPORTO ===

def estrai_numero(alias):
    match = re.search(r'\b\d{3}\b', str(alias))
    return match.group(0) if match else None

def minuti_con_secondi(durata_str):
    try:
        minuti, secondi = map(int, str(durata_str).split(":"))
        return minuti + secondi / 60
    except:
        return 0

# === CARICAMENTO E PULIZIA FILE 1 ===
df1 = pd.read_excel(file1_path)
df1["numero"] = df1["Operatore"].apply(estrai_numero)
df1 = df1[["numero", "Chiamate Risposte", "Minuti"]].copy()
df1.columns = ["numero", "chiamate", "minuti"]
df1["chiamate"] = pd.to_numeric(df1["chiamate"], errors="coerce")
df1["minuti"] = pd.to_numeric(df1["minuti"], errors="coerce")

# === CARICAMENTO E PULIZIA FILE 2 ===
df2 = pd.read_excel(file2_path, skiprows=2)
df2.columns = ["alias", "chiamate", "durata", "media_durata"]
df2["numero"] = df2["alias"].apply(estrai_numero)
df2["minuti"] = df2["durata"].apply(minuti_con_secondi)
df2 = df2[["numero", "chiamate", "minuti"]]

# === CONFRONTO ===
confronto = pd.merge(df1, df2, on="numero", how="outer", suffixes=("_file1", "_file2"))
confronto["diff_chiamate"] = confronto["chiamate_file1"] - confronto["chiamate_file2"]
confronto["diff_minuti"] = confronto["minuti_file1"] - confronto["minuti_file2"]

# === SALVATAGGIO RISULTATO ===
confronto.to_excel("confronto_chiamate_minuti_CORRETTO.xlsx", index=False)
print("✅ File di confronto generato: confronto_chiamate_minuti_CORRETTO.xlsx")
