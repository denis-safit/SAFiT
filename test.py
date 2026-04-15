import pandas as pd
df = pd.read_excel('righe_ordini_storico_con_date.xlsx', skiprows=2)
df.columns = [str(c).strip() for c in df.columns]
df['Codice Documento'] = df['Codice Documento'].ffill()
df['Qta Residua'] = pd.to_numeric(df['Qta Residua'], errors='coerce').fillna(0)
off_aperte = df[df['Codice Documento'].isin(['OFF','OFR']) & (df['Qta Residua'] > 0)]
print('Righe OFF/OFR aperte:', len(off_aperte))
print(off_aperte[['Codice Documento','Articolo C','Qta Residua']].head(10).to_string())
