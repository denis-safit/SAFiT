import pyodbc

path_access = r'E:\Safit\Produzione\Produzione\SAFIT\NGP - x Denis.accdb'
conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={path_access};'

try:
    conn = pyodbc.connect(conn_str)
    print("✅ Connessione ad Access riuscita!")
    conn.close()
except Exception as e:
    print(f"❌ Errore persistente: {e}")
