# -*- coding: utf-8 -*-
"""calcolo.ipynb

Automatically generated by Colab.

Original file is located at
    https://colab.research.google.com/drive/1r_rCWBQ7H3rzKQiVSI8g_ZeqYuRCXQeJ
"""

import streamlit as st
import pandas as pd
from io import BytesIO

# Funzione principale della app
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

        # Rimuovi la colonna temporanea
        df = df[['Mese Anno', 'Ore straordinarie', 'Ore recupero']]

        def calcola_ore_finali(row):
            return row['Ore straordinarie'] - row['Ore recupero']
        # Calcolo della colonna Ore_finali
        df['Ore finali'] = df.apply(calcola_ore_finali, axis=1)

        # Raggruppa per Mese Anno e somma le colonne
        riepilogo = df.groupby('Mese Anno')[['Ore straordinarie', 'Ore recupero', 'Ore finali']].sum().reset_index()
        riepilogo["Ore straordinarie"] = riepilogo["Ore straordinarie"] / 3600
        riepilogo["Ore recupero"] = riepilogo["Ore recupero"] / 3600
        riepilogo["Ore finali"] = riepilogo["Ore finali"] / 3600

        # Ordine temporale
        mesi_italiani_reverse = {
            1: 'Gennaio', 2: 'Febbraio', 3: 'Marzo', 4: 'Aprile', 5: 'Maggio', 6: 'Giugno',
            7: 'Luglio', 8: 'Agosto', 9: 'Settembre', 10: 'Ottobre', 11: 'Novembre', 12: 'Dicembre'
        }

        # Funzione per convertire 'Mese_Anno' in formato 'YYYY-MM'
        def converti_mese_anno(mese_anno):
            mese, anno = mese_anno.split()
            mese = mese.strip()  # Rimuovere eventuali spazi prima e dopo il mese
            mese_num = None
            for num, nome in mesi_italiani_reverse.items():
                if mese == nome:
                    mese_num = num
                    break

            if mese_num:  # Se il mese è trovato
                return f"{anno}-{mese_num:02d}"
            else:
                return None  # Restituire None se il mese non è trovato

        # Applicare la funzione alla colonna 'Mese_Anno'
        riepilogo['Anno_Mese'] = riepilogo['Mese Anno'].apply(converti_mese_anno)
        riepilogo = riepilogo.dropna(subset=['Anno_Mese'])
        riepilogo['Anno_Mese'] = pd.to_datetime(riepilogo['Anno_Mese'], format='%Y-%m')
        riepilogo = riepilogo.sort_values(by='Anno_Mese')
        riepilogo = riepilogo.drop('Anno_Mese', axis=1)

        # Aggiunta input per permessi
        if "permessi_input" not in st.session_state:
            st.session_state["permessi_input"] = 0.0
        st.write("Inserisci i permessi mensili:")
        col1, col2, col3 = st.columns(3)
        selected_month = col1.selectbox("Seleziona il mese", list(mesi_italiani.values()))
        selected_year = col2.number_input("Inserisci l'anno", min_value=2000, max_value=2100, step=1, value=2023)
        ore_permesso = col3.number_input("Ore di permesso (in ore)", min_value=0.0, step=0.5, value=0.0, key="permessi_input")

        mese_anno_permesso = f"{selected_month} {int(selected_year)}"

        # Aggiunta colonna "Ore permesso"
        riepilogo['Ore permesso'] = 0.0

        if ore_permesso > 0:
            if mese_anno_permesso in riepilogo['Mese Anno'].values:
                riepilogo.loc[riepilogo['Mese Anno'] == mese_anno_permesso, 'Ore permesso'] = ore_permesso
                riepilogo.loc[riepilogo['Mese Anno'] == mese_anno_permesso, 'Ore finali'] -= ore_permesso
            st.session_state.permessi_input = 0.0  # Resetta il valore a 0 dopo averlo registrato

        # Calcolo dei cumulativi aggiornati
        cumulative_hours = 0
        cumulative_times = []
        for final_hours in riepilogo["Ore finali"]:
            cumulative_hours += final_hours
            cumulative_times.append(cumulative_hours)

        riepilogo["Cumulativo Ore"] = cumulative_times

        # Funzione per convertire i secondi in formato HH:MM:SS
        def convert_seconds(hours):
            total_seconds = int(hours * 3600)
            is_negative = total_seconds < 0
            total_seconds = abs(total_seconds)
            hh = total_seconds // 3600
            mm = (total_seconds % 3600) // 60
            ss = total_seconds % 60
            time_str = f"{hh:02}:{mm:02}:{ss:02}"
            return f"-{time_str}" if is_negative else time_str

        # Applicare la conversione
        riepilogo['Ore straordinarie'] = riepilogo['Ore straordinarie'].apply(convert_seconds)
        riepilogo['Ore recupero'] = riepilogo['Ore recupero'].apply(convert_seconds)
        riepilogo['Ore finali'] = riepilogo['Ore finali'].apply(convert_seconds)
        riepilogo['Cumulativo Ore'] = riepilogo['Cumulativo Ore'].apply(convert_seconds)
        riepilogo['Ore permesso'] = riepilogo['Ore permesso'].apply(convert_seconds)

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
