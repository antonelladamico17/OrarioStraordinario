import openpyxl
import pandas as pd
import streamlit as st


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
        df['Entrata'] = df['Data'].astype(str) + ' ' + df['Orario entrata']
        df['Uscita'] = df['Data'].astype(str) + ' ' + df['Orario uscita']
        df['Entrata'] = pd.to_datetime(df['Entrata'], format='%d/%m/%Y %H:%M:%S')
        df['Uscita'] = pd.to_datetime(df['Uscita'], format='%d/%m/%Y %H:%M:%S')

        # Calcolo delle ore straordinarie
        orario_lavorativo_standard = pd.to_timedelta('07:12:00')
        df['Differenza'] = df['Uscita'] - df['Entrata']
        df['Ore_straordinarie'] = df['Differenza'].where(df['Differenza'] > orario_lavorativo_standard, pd.Timedelta(0)) - orario_lavorativo_standard
        df['Ore_straordinarie'] = df['Ore_straordinarie'].where(df['Ore_straordinarie'] > pd.Timedelta(0), pd.Timedelta(0))

        # Creazione del riepilogo mensile
        df['Mese_Anno'] = df['Uscita'].dt.to_period('M')
        riepilogo = df.groupby('Mese_Anno')['Ore_straordinarie'].sum().reset_index()

        st.write("Riepilogo delle Ore Straordinarie:")
        st.dataframe(riepilogo)

        # Opzione per scaricare il riepilogo
        csv = riepilogo.to_csv(index=False).encode('utf-8')
        st.download_button(label="Scarica Riepilogo", data=csv, file_name='riepilogo_ore_straordinarie.csv', mime='text/csv')

if __name__ == "__main__":
    main()

