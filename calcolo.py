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
        df['Entrata'] = df['Data'].astype(str) + ' ' + df['Orario entrata']
        df['Uscita'] = df['Data'].astype(str) + ' ' + df['Orario uscita']
        df['Entrata'] = pd.to_datetime(df['Entrata'], format='%d/%m/%Y %H:%M:%S')
        df['Uscita'] = pd.to_datetime(df['Uscita'], format='%d/%m/%Y %H:%M:%S')

        # Calcolo delle ore straordinarie
        orario_lavorativo_standard = pd.to_timedelta('07:12:00')
        df['Differenza'] = df['Uscita'] - df['Entrata']
        df['Ore_straordinarie'] = df['Differenza'].where(df['Differenza'] > orario_lavorativo_standard, pd.Timedelta(0))
        df['Ore_straordinarie'] = df['Ore_straordinarie'] - orario_lavorativo_standard
        df['Ore_straordinarie'] = df['Ore_straordinarie'].where(df['Ore_straordinarie'] > pd.Timedelta(0), pd.Timedelta(0))

        # Formattazione delle ore straordinarie come stringhe
        df['Ore_straordinarie'] = df['Ore_straordinarie'].apply(format_timedelta)

        # Creazione del riepilogo mensile
        df['Mese_Anno'] = df['Uscita'].dt.to_period('M')
        riepilogo = df.groupby('Mese_Anno')['Ore_straordinarie'].first().reset_index()

        st.write("Riepilogo delle Ore Straordinarie:")
        st.dataframe(riepilogo)

        # Opzione per scaricare il riepilogo in formato Excel
        excel_file = create_excel_file(riepilogo)
        st.download_button(label="Scarica Riepilogo", data=excel_file, file_name='riepilogo_ore_straordinarie.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def format_timedelta(td):
    total_seconds = int(td.total_seconds())
    days, remainder = divmod(total_seconds, 86400)  # 86400 secondi in un giorno
    hours, remainder = divmod(remainder, 3600)
    minutes, seconds = divmod(remainder, 60)
    
    # Se i giorni sono maggiori di 0, aggiungi 24 alle ore
    if days > 0:
        hours += days * 24
    
    return f"{hours}:{minutes:02}:{seconds:02}"

def create_excel_file(df):
    # Crea un file Excel in memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Riepilogo')
    output.seek(0)
    return output.read()

if __name__ == "__main__":
    main()
