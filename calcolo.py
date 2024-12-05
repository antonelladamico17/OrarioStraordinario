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

        # Conversione delle date e orari
        df['Entrata'] = pd.to_datetime(df['Data'] + ' ' + df['Orario entrata'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
        df['Uscita'] = pd.to_datetime(df['Data'] + ' ' + df['Orario uscita'], format='%d/%m/%Y %H:%M:%S', errors='coerce')

        # Calcolo della durata
        df['Durata'] = (df['Uscita'] - df['Entrata']).dt.total_seconds()  # Totale in secondi
        df.loc[df['Causale'] == 'Orario Ordinario', 'Durata'] = 0

        # Estrarre il giorno per il raggruppamento
        df['Giorno'] = df['Entrata'].dt.strftime('%d/%m/%Y')

        # Raggruppare per giorno e sommare i secondi
        df_raggruppato = df.groupby('Giorno')['Durata'].sum().reset_index()

        # Orario lavorativo standard
        orario_lavorativo_standard = pd.to_timedelta('07:12:00')

        # Convertire i secondi in timedelta
        df_raggruppato['Ore totali'] = pd.to_timedelta(df_raggruppato['Durata'], unit='s')

        # Calcolo straordinari e recupero
        def calcola_ore(row):
            differenza = row['Ore totali'] - orario_lavorativo_standard
            if differenza >= pd.Timedelta(0):  # Straordinari
                return differenza, pd.Timedelta(0)
            else:  # Recupero
                return pd.Timedelta(0), abs(differenza)

        # Applicare la funzione calcola_ore
        df_raggruppato[['Ore_straordinarie', 'Ore_recupero']] = df_raggruppato.apply(
            lambda row: pd.Series(calcola_ore(row)), axis=1
        )

        # Conversione delle ore in formato HH:MM:SS
        for col in ['Ore totali', 'Ore_straordinarie', 'Ore_recupero']:
            df_raggruppato[col] = df_raggruppato[col].apply(lambda x: str(x).split(' ')[-1] if not pd.isna(x) else '')

        # Creare la colonna 'Mese_Anno' con il nome del mese in italiano
        df_raggruppato['Giorno'] = pd.to_datetime(df_raggruppato['Giorno'], format='%d/%m/%Y')
        mesi_italiani = {
            1: 'Gennaio', 2: 'Febbraio', 3: 'Marzo', 4: 'Aprile', 5: 'Maggio', 6: 'Giugno',
            7: 'Luglio', 8: 'Agosto', 9: 'Settembre', 10: 'Ottobre', 11: 'Novembre', 12: 'Dicembre'
        }
        df_raggruppato['Mese_Anno'] = df_raggruppato['Giorno'].dt.month.map(mesi_italiani) + ' ' + df_raggruppato['Giorno'].dt.year.astype(str)

        # Calcolo riepilogo mensile
        riepilogo = df_raggruppato.groupby('Mese_Anno')[['Ore_straordinarie', 'Ore_recupero']].sum().reset_index()

        # Convertire le colonne in timedelta
        riepilogo['Ore_straordinarie'] = pd.to_timedelta(riepilogo['Ore_straordinarie'], errors='coerce')
        riepilogo['Ore_recupero'] = pd.to_timedelta(riepilogo['Ore_recupero'], errors='coerce')

        # Calcolo Ore_finali
        riepilogo['Ore_finali'] = riepilogo.apply(
            lambda row: abs(row['Ore_straordinarie'] - row['Ore_recupero']), axis=1
        )

        # Conversione per leggibilità
        for col in ['Ore_straordinarie', 'Ore_recupero', 'Ore_finali']:
            riepilogo[col] = riepilogo[col].apply(lambda x: str(x).split(' ')[-1] if not pd.isna(x) else '')

        # Mostra il riepilogo
        st.write("Riepilogo delle Ore Straordinarie:")
        st.dataframe(riepilogo)

        # Opzione per scaricare il riepilogo in formato Excel
        excel_file = create_excel_file(riepilogo)
        st.download_button(
            label="Scarica Riepilogo",
            data=excel_file,
            file_name='riepilogo_ore_straordinarie.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

def create_excel_file(df):
    # Crea un file Excel in memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Riepilogo')
    output.seek(0)
    return output.read()

if __name__ == "__main__":
    main()
