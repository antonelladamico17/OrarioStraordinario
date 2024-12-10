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
        df = df.groupby('Giorno')['Durata'].sum().reset_index()

        # Orario lavorativo standard
        orario_lavorativo_standard = pd.to_timedelta('07:12:00').total_seconds()

        # Calcolo straordinari e recupero
        def calcola_ore(row):
            differenza = row['Durata'] - orario_lavorativo_standard
            if differenza >= 0:  # Straordinari
                return differenza, 0
            else:  # Recupero
                return 0, orario_lavorativo_standard - row['Durata']

        # Applicare la funzione calcola_ore
        df[['Ore straordinarie', 'Ore recupero']] = df.apply(lambda row: pd.Series(calcola_ore(row)), axis=1)


        # Creare la colonna 'Mese Anno' con il nome del mese in italiano
        df['Giorno'] = pd.to_datetime(df['Giorno'], format='%d/%m/%Y')
        mesi_italiani = {
            1: 'Gennaio', 2: 'Febbraio', 3: 'Marzo', 4: 'Aprile', 5: 'Maggio', 6: 'Giugno',
            7: 'Luglio', 8: 'Agosto', 9: 'Settembre', 10: 'Ottobre', 11: 'Novembre', 12: 'Dicembre'
        }
        df['Mese Anno'] = pd.to_datetime(df['Giorno'], format='%d/%m/%Y').dt.month
        df['Mese Anno'] = df['Mese Anno'].map(mesi_italiani) + ' ' + pd.to_datetime(df['Giorno'], format='%d/%m/%Y').dt.year.astype(str)

        # Calcolo riepilogo mensile
        df = df[['Mese Anno', 'Ore straordinarie', 'Ore recupero']]

        
        def calcola_ore_finali(row):
          return row['Ore straordinarie'] - row['Ore recupero']
        # Calcolo della colonna Ore_finali
        df['Ore finali'] = df.apply(calcola_ore_finali, axis=1)

        # Raggruppa per Mese Anno e somma le colonne
        riepilogo = df.groupby('Mese Anno')[['Ore straordinarie', 'Ore recupero', 'Ore finali']].sum().reset_index()
        riepilogo["Ore finali"] = riepilogo["Ore finali"] / 3600

        # Funzione per convertire i secondi in formato HH:MM:SS
        def convert_seconds(seconds):
    # Verifica se i secondi sono negativi
            is_negative = seconds < 0
            seconds = abs(seconds)  # Prendiamo il valore assoluto dei secondi per lavorare con il numero positivo
    
    # Calcolare ore
            hours = int(seconds)  # Otteniamo la parte intera come ore
    
    # Calcolare i minuti e secondi dai decimali
            minutes = int((seconds - hours) * 60)
            remaining_seconds = int(((seconds - hours) * 60 - minutes) * 60)
    
    # Creare la stringa nel formato HH:MM:SS
            time_str = f"{hours:02}:{minutes:02}:{remaining_seconds:02}"
    
    # Aggiungere il segno negativo se i secondi sono negativi
            if is_negative:
                time_str = "-" + time_str
    
            return time_str

        # Applicare la funzione alla colonna 'Ore_finali_format'
        riepilogo['Ore straordinarie'] = riepilogo['Ore straordinarie'].apply(convert_seconds)
        riepilogo['Ore recupero'] = riepilogo['Ore recupero'].apply(convert_seconds)
        riepilogo['Ore finali'] = riepilogo['Ore finali'].apply(convert_seconds)

        # Mostra il riepilogo
        st.write("Riepilogo delle Ore Straordinarie:")
        st.dataframe(riepilogo)

        # Opzione per scaricare il riepilogo in formato Excel
        excel_file = create_excel_file(riepilogo)
        st.download_button(
            label="Scarica Riepilogo",
            data=excel_file,
            file_name='riepilogo_Ore straordinarie.xlsx',
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
