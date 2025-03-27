import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from datetime import datetime


def estrai_testo_da_pdf_testata(file_pdf):
    with pdfplumber.open(file_pdf) as pdf:
        first_page_text = pdf.pages[0].extract_text()
        match = re.search(r"F.A.S.I. DETTAGLIO RIMBORSI:\s*(.*?)\s* - ", first_page_text, re.IGNORECASE)
        if match:
            return match.group(1)
    return "no_testo" 

def estrai_data_da_pdf_testata(file_pdf):
     with pdfplumber.open(file_pdf) as pdf:
        first_page_text = pdf.pages[0].extract_text()
        match = re.search(r"Chiusura(.*?)del(.*?)Pagina", first_page_text, re.IGNORECASE)
        if match:
            return match.group(2)
        return "no_data"


def estrai_dati_da_pdf(lista_file_pdf):
    dati_completi = []

    for file_pdf in lista_file_pdf:
        s=estrai_testo_da_pdf_testata(file_pdf)
        dt=estrai_data_da_pdf_testata(file_pdf)
        with pdfplumber.open(file_pdf) as pdf:
           
            for page in pdf.pages:
                tables = page.extract_table()
                if tables:
                    for row in tables[1:]:
                        if len(row) >= 8:
                            a = row[3]
                            nf= row[4]
                            b = row[5]
                            c = row[6]
                            d = row[8]
                            if b and c and d:
                                dati_completi.append([s,dt,a,nf, b, c, d])

    colonne_selezionate = ["Società Testata","Data Testata","Nominativo Dirigente","Nominativo Familiare", "Data Fattura", "Numero Fattura", "Totale Rimborsato"]
    
    return pd.DataFrame(dati_completi, columns=colonne_selezionate)

# UI Streamlit
st.title("Estrazione Tabelle da PDF FasiOpen")

# Upload multiplo di file PDF
file_caricati = st.file_uploader("Carica i file PDF", type=["pdf"], accept_multiple_files=True)

if file_caricati:
    st.success(f"{len(file_caricati)} file caricati con successo!")
    
    # Estrai dati dai PDF
    df_finale = estrai_dati_da_pdf(file_caricati)

    if not df_finale.empty:
        st.write("### Anteprima Dati Estratti")
        st.dataframe(df_finale.head())

        # Creazione file Excel in memoria
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_finale.to_excel(writer, index=False, sheet_name="Rimborsi")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"estratto_fasi_open_{timestamp}.xlsx"
        # Download del file Excel
        st.download_button(
            label="📥 Scarica il file Excel",
            data=buffer.getvalue(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Nessuna tabella trovata nei PDF!")
