import pandas as pd

file_input = 'righe_Ordini_ARCA.xlsx'
nome_foglio_pivot = 'Foglio1' 

def carica_dati_per_claude():
    try:
        # Carichiamo saltando le prime righe se la Pivot non parte dalla riga 1
        # Se i nomi delle colonne sono alla riga 2, usa skiprows=1
        df = pd.read_excel(file_input, sheet_name=nome_foglio_pivot, skiprows=1)

        # TRUCCO FONDAMENTALE: Riempie le celle vuote della Pivot (Forward Fill)
        # Se il Cliente è scritto solo sulla prima riga, lo copia in basso
        df = df.ffill()

        # Pulizia: Rimuovi righe dove la colonna principale è vuota o è un totale
        df = df.dropna(subset=[df.columns[1]]) 

        # Salvataggio
        dati_markdown = df.to_markdown(index=False)
        with open("dati_per_claude.txt", "w", encoding="utf-8") as f:
            f.write(dati_markdown)
        
        print("✅ Dati ottimizzati e salvati in 'dati_per_claude.txt'!")
        print(df.head()) # Vedi la differenza nel terminale!

    except Exception as e:
        print(f"❌ Errore: {e}")

if __name__ == "__main__":
    carica_dati_per_claude()
