import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import zipfile
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

def estrai_pdf_da_zip(file_zip):
    pdf_files = []
    with zipfile.ZipFile(file_zip, "r") as zip_ref:
        for file_name in zip_ref.namelist():
            if file_name.lower().endswith(".pdf"):  
                with zip_ref.open(file_name) as pdf_file:
                    pdf_files.append(io.BytesIO(pdf_file.read()))  
    return pdf_files

def estrai_dati_da_pdf(lista_file_pdf):
    dati_completi = []

    for file_pdf in lista_file_pdf:
        s=estrai_testo_da_pdf_testata(file_pdf)
        dt=estrai_data_da_pdf_testata(file_pdf)
        with pdfplumber.open(file_pdf) as pdf:
           
            for page in pdf.pages:
                tables = page.extract_table()
                if tables:
                     for i, row in enumerate(tables[1:], 1):
                        st.info(row)
                        st.info(len(row))
                        if len(row) == 9:
                            st.info(row[3])
                            st.info(row[4])
                            st.info(row[5])
                            st.info(row[6])
                            st.info(row[9])
                            st.info("---------------------------------------------------------");
                            a = row[3]
                            nf= row[4]
                            b = row[5]
                            c = row[6]
                            d = row[9]
                            if (not(c==None) and d!= "0,00" and d!=""):
                                if ((a==None or a=='') and i>0):
                                    a=tables[i-1][3]
                                if ((nf==None or nf=='') and i>0):
                                    nf=tables[i-1][4]     
                                dati_completi.append([s,dt,a,nf, b, c, d])
                        else:
                            if len(row) == 10:
                                st.info(row[3])
                                st.info(row[4])
                                st.info(row[5])
                                st.info(row[6])
                                st.info(row[10])
                                st.info("---------------------------------------------------------");
                                a = row[3]
                                nf= row[4]
                                b = row[5]
                                c = row[6]
                                d = row[10]
                                if (not(c==None) and d!= "0,00" and d!=""):
                                    if ((a==None or a=='') and i>0):
                                        a=tables[i-1][3]
                                    if ((nf==None or nf=='') and i>0):
                                        nf=tables[i-1][4]     
                                    dati_completi.append([s,dt,a,nf, b, c, d])
                            else:
                                if len(row) == 11:
                                    st.info('3'+row[3])
                                    st.info('4'+row[4])
                                    st.info('5'+row[5])
                                    st.info('6'+row[6])
                                    st.info('11'+row[11])
                                    a = row[3]
                                    nf= row[4]
                                    b = row[5]
                                    c = row[6]
                                    d = row[11]
                                    st.info("---------------------------------------------------------");
                                    if (not(c==None) and d!= "0,00" and d!=""):
                                        if ((a==None or a=='') and i>0):
                                            a=tables[i-1][3]
                                        if ((nf==None or nf=='') and i>0):
                                            nf=tables[i-1][4]     
                                    dati_completi.append([s,dt,a,nf, b, c, d])
                                else:
                                    st.info("lunghezza row errata")

    colonne_selezionate = ["Società Testata","Data Testata","Nominativo Dirigente","Nominativo Familiare", "Data Fattura", "Numero Fattura", "Totale Rimborsato"]
    
    return pd.DataFrame(dati_completi, columns=colonne_selezionate)

st.title("Estrazione Tabelle da PDF FasiOpen")
st.info("Build 1.3.2.3 - 09/04/2025")

file_caricati = st.file_uploader("Carica i file PDF o ZIP", type=["pdf","zip"], accept_multiple_files=True)

if file_caricati:
    
    st.success(f"{len(file_caricati)} file caricati con successo!")
    lista_pdf = []
    
    for file in file_caricati:
        if file.name.lower().endswith(".pdf"):
            lista_pdf.append(file)
        elif file.name.lower().endswith(".zip"):
            lista_pdf.extend(estrai_pdf_da_zip(file))  

    if lista_pdf:
        df_finale = estrai_dati_da_pdf(lista_pdf)

        if not df_finale.empty:
            st.write("### Anteprima Dati Estratti")
            st.dataframe(df_finale.head())

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_finale.to_excel(writer, index=False, sheet_name="Rimborsi")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_name = f"estratto_fasi_open_{timestamp}.xlsx"

            st.download_button(
                label="📥 Scarica il file Excel",
                data=buffer.getvalue(),
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nessuna tabella trovata nei PDF!")
