import pandas as pd
import streamlit as st
from io import BytesIO

def main():
    st.title("Analisi Ore Straordinarie")

    # Carica il file Excel
    uploaded_file = st.file_uploader("Carica il tuo file Excel", type="xlsx")
    
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)

        # Gestione delle intestazioni
        df = df.iloc[1:]
        df.columns = df.iloc[0]
        df = df.iloc[1:]
        df.reset_index(drop=True, inplace=True)

        # Crea le colonne Entrata e Uscita
        df['Entrata'] = pd.to_datetime(df['Data'] + ' ' + df['Orario entrata'], format='%d/%m/%Y %H:%M:%S')
        df['Uscita'] = pd.to_datetime(df['Data'] + ' ' + df['Orario uscita'], format='%d/%m/%Y %H:%M:%S')

        # Calcolare la durata in secondi
        df['Durata'] = (df['Uscita'] - df['Entrata']).dt.total_seconds()  # Totale in secondi

        # Impostare la durata a 0 se la causale è "Orario Ordinario"
        df.loc[df['Causale'] == 'Orario Ordinario', 'Durata'] = 0

        # Estrarre solo il giorno per raggruppare
        df['Giorno'] = df['Entrata'].dt.strftime('%d/%m/%Y')

        # Raggruppare per giorno e sommare i secondi
        df_raggruppato = df.groupby('Giorno')['Durata'].sum().reset_index()

        # Orario lavorativo standard
        orario_lavorativo_standard = pd.to_timedelta('07:12:00')

        # Convertire i secondi in timedelta
        df_raggruppato['Ore totali'] = pd.to_timedelta(df_raggruppato['Durata'], unit='s')

        # Calcolare straordinari e recupero
        def calcola_ore(row):
            differenza = row['Ore totali'] - orario_lavorativo_standard
            if differenza >= pd.Timedelta(0):  # Se positivo, è straordinario
                return differenza, pd.Timedelta(0)
            else:  # Se negativo, calcolare il recupero
                return pd.Timedelta(0), orario_lavorativo_standard - row['Ore totali']

        # Applicare la funzione riga per riga
        df_raggruppato[['Ore_straordinarie', 'Ore_recupero']] = df_raggruppato.apply(
        lambda row: pd.Series(calcola_ore(row)), axis=1)

        # Convertire i risultati in formato HH:MM:SS per leggibilità
        df_raggruppato['Ore totali'] = df_raggruppato['Ore totali'].apply(lambda x: str(x).split(' ')[-1])
        df_raggruppato['Ore_straordinarie'] = df_raggruppato['Ore_straordinarie'].apply(lambda x: str(x).split(' ')[-1])
        df_raggruppato['Ore_recupero'] = df_raggruppato['Ore_recupero'].apply(lambda x: str(x).split(' ')[-1])
        # Convert 'Giorno' column to datetime
        df_raggruppato['Giorno'] = pd.to_datetime(df_raggruppato['Giorno'], format='%d/%m/%Y')

        # Mapping manuale dei mesi in italiano
        mesi_italiani = {1: 'Gennaio', 2: 'Febbraio', 3: 'Marzo', 4: 'Aprile', 5: 'Maggio', 6: 'Giugno', 7: 'Luglio', 8: 'Agosto', 9: 'Settembre', 10: 'Ottobre', 11: 'Novembre', 12: 'Dicembre'}

        # Creare la colonna 'Mese_Anno' con il nome del mese in italiano
        df_raggruppato['Mese_Anno'] = pd.to_datetime(df_raggruppato['Giorno'], format='%d/%m/%Y').dt.month
        df_raggruppato['Mese_Anno'] = df_raggruppato['Mese_Anno'].map(mesi_italiani) + ' ' + pd.to_datetime(df_raggruppato['Giorno'], format='%d/%m/%Y').dt.year.astype(str)
        
        # Rimuovi la colonna temporanea
        riepilogo = df_raggruppato[['Mese_Anno', 'Ore_straordinarie', 'Ore_recupero']]

        # Assicurati che le colonne Ore_straordinarie e Ore_recupero siano in formato timedelta
        riepilogo['Ore_straordinarie'] = pd.to_timedelta(riepilogo['Ore_straordinarie'])
        riepilogo['Ore_recupero'] = pd.to_timedelta(riepilogo['Ore_recupero'])

        # Funzione per calcolare la differenza tra straordinari e recupero
        def calcola_ore_finali(row):
            if row['Ore_straordinarie'] > row['Ore_recupero']:
                return row['Ore_straordinarie'] - row['Ore_recupero']
            else:
                return row['Ore_recupero'] - row['Ore_straordinarie']

        # Calcolo della colonna Ore_finali
        riepilogo['Ore_finali'] = riepilogo.apply(calcola_ore_finali, axis=1)

# Raggruppa per Mese_Anno e somma le colonne
        riepilogo = riepilogo.groupby('Mese_Anno')[['Ore_straordinarie', 'Ore_recupero', 'Ore_finali']].sum().reset_index()

# Converti le colonne in formato leggibile HH:MM:SS
        riepilogo['Ore_straordinarie'] = riepilogo['Ore_straordinarie'].apply(lambda x: str(x).split(' ')[-1])
        riepilogo['Ore_recupero'] = riepilogo['Ore_recupero'].apply(lambda x: str(x).split(' ')[-1])
        riepilogo['Ore_finali'] = riepilogo['Ore_finali'].apply(lambda x: str(x).split(' ')[-1])
        
        # Mostra il riepilogo
        st.write("Riepilogo delle Ore Straordinarie:")
        st.dataframe(riepilogo)

        # Opzione per scaricare il riepilogo in formato Excel
        excel_file = create_excel_file(riepilogo)
        st.download_button(label="Scarica Riepilogo", data=excel_file, file_name='riepilogo_ore_straordinarie.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

#def format_timedelta(td):
 #   total_seconds = int(td.total_seconds())
 #   days, remainder = divmod(total_seconds, 86400)  # 86400 secondi in un giorno
 #   hours, remainder = divmod(remainder, 3600)
 #   minutes, seconds = divmod(remainder, 60)
    
    # Se i giorni sono maggiori di 0, aggiungi 24 alle ore
    #if days > 0:
    #    hours += days * 24
    
    #return f"{hours}:{minutes:02}:{seconds:02}"

def create_excel_file(df):
    # Crea un file Excel in memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Riepilogo')
    output.seek(0)
    return output.read()

if __name__ == "__main__":
    main()
