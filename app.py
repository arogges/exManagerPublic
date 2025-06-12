import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import zipfile
from datetime import datetime
import traceback
import os


def estrai_testo_da_pdf_testata(file_pdf, file_name):
    try:
        with pdfplumber.open(file_pdf) as pdf:
            first_page_text = pdf.pages[0].extract_text()
            match = re.search(r"F.A.S.I. DETTAGLIO RIMBORSI:\s*(.*?)\s* - ", first_page_text, re.IGNORECASE)
            if match:
                return match.group(1)
        return "no_testo"
    except Exception as e:
        st.error(f"Errore nell'estrazione del testo dalla testata del file: {file_name}")
        st.error(f"Dettaglio errore: {str(e)}")
        return "errore_estrazione_testata"

def estrai_data_da_pdf_testata(file_pdf, file_name):
    try:
        with pdfplumber.open(file_pdf) as pdf:
            first_page_text = pdf.pages[0].extract_text()
            match = re.search(r"Chiusura(.*?)del(.*?)Pagina", first_page_text, re.IGNORECASE)
            if match:
                return match.group(2)
            return "no_data"
    except Exception as e:
        st.error(f"Errore nell'estrazione della data dalla testata del file: {file_name}")
        st.error(f"Dettaglio errore: {str(e)}")
        return "errore_estrazione_data"

def estrai_dati_formato_nuovo(file_pdf, file_name):
    """
    Estrae dati dal nuovo formato PDF (come 011-FasiOpen.pdf)
    Analizza la seconda pagina per estrarre i dati dalla tabella
    """
    try:
        with pdfplumber.open(file_pdf) as pdf:
            #st.info(f"🔍 DEBUG: Analizzando {file_name} - Numero pagine: {len(pdf.pages)}")
            
            # Estrai informazioni dalla prima pagina
            first_page_text = pdf.pages[0].extract_text()
            
            # DEBUG: Mostra parte del testo della prima pagina
            #if st.checkbox(f"🔍 DEBUG: Mostra testo prima pagina di {file_name}", key=f"debug1_{file_name}"):
            #    st.text(first_page_text[:500] + "...")
            
            # Estrai la società/fornitore (prima riga del documento)
            societa_match = re.search(r"^(.+?)\n", first_page_text)
            societa = societa_match.group(1).strip() if societa_match else "no_societa"
            
            # Estrai la data dal documento
            data_match = re.search(r"Roma,\s*(\d{2}/\d{2}/\d{4})", first_page_text)
            data_documento = data_match.group(1) if data_match else "no_data"
            
            #st.info(f"📊 DEBUG: Società: {societa}, Data: {data_documento}")
            
            # Analizza la seconda pagina per i dati della tabella
            if len(pdf.pages) < 2:
                st.warning(f"Il file {file_name} non ha una seconda pagina")
                return []
            
            second_page = pdf.pages[1]
            second_page_text = second_page.extract_text()
            
            # DEBUG: Mostra testo seconda pagina
            #if st.checkbox(f"🔍 DEBUG: Mostra testo seconda pagina di {file_name}", key=f"debug2_{file_name}"):
            #    st.text(second_page_text[:1000] + "...")
            
            tables = second_page.extract_table()
            
           # st.info(f"📋 DEBUG: Tabella trovata: {tables is not None}")
           # if tables:
            #    st.info(f"📋 DEBUG: Numero righe tabella: {len(tables)}")
                
                # DEBUG: Mostra le prime righe della tabella
             #   if st.checkbox(f"🔍 DEBUG: Mostra prime righe tabella di {file_name}", key=f"debug3_{file_name}"):
             #       for i, row in enumerate(tables[:5]):
             #           st.write(f"Riga {i}: {row}")
            
            dati_estratti = []
            
            if tables:
                # Trova l'indice della riga di intestazione
                header_row_idx = None
                for i, row in enumerate(tables):
                    if row and any(cell and ("Iscritto Principale" in str(cell) or "Main Client" in str(cell)) for cell in row):
                        header_row_idx = i
              #          st.info(f"📋 DEBUG: Header trovato alla riga {i}")
                        break
                
                if header_row_idx is None:
                    st.warning(f"Header non trovato in {file_name}. Cerco pattern alternativi...")
                    # Proviamo a cercare altre intestazioni
                    for i, row in enumerate(tables):
                        if row and any(cell and ("FasiOpen" in str(cell) or "Nominativo" in str(cell)) for cell in row if cell):
                            header_row_idx = i
               #             st.info(f"📋 DEBUG: Header alternativo trovato alla riga {i}")
                            break
                
                if header_row_idx is not None:
                #    st.info(f"📋 DEBUG: Processando righe da {header_row_idx + 1} a {len(tables)}")
                    
                    # Processa le righe dopo l'intestazione
                    current_nominativo = None
                    
                    for i in range(header_row_idx + 1, len(tables)):
                        row = tables[i]
                        if not row:
                            continue
                        
                 #       st.info(f"📋 DEBUG: Riga {i} - Lunghezza: {len(row)} - Contenuto: {row}")
                        
                        if len(row) < 6:
                          #  st.warning(f"Riga {i} troppo corta, saltata")
                            continue
                        
                        # Salta le righe "Totale"
                        if row and any(cell and "Totale" in str(cell) for cell in row if cell):
                           # st.info(f"Riga {i} è un totale, saltata")
                            continue
                        
                        # Estrai i dati dalla riga
                        cod_fasiopen = row[0] if len(row) > 0 and row[0] else ""
                        nominativo = row[1] if len(row) > 1 and row[1] else ""
                        nominativo_familiare = row[2] if len(row) > 2 and row[2] else ""
                        data_fattura = row[3] if len(row) > 3 and row[3] else ""
                        numero_fattura = row[4] if len(row) > 4 and row[4] else ""
                        importo_liquidato = row[6] if len(row) > 6 and row[6] else ""
                        
                  #      st.info(f"📋 DEBUG: Estratti - Nom: '{nominativo}', Data: '{data_fattura}', Num: '{numero_fattura}', Imp: '{importo_liquidato}'")
                        
                        # Se il nominativo è vuoto, usa quello precedente
                        if nominativo and nominativo.strip():
                            current_nominativo = nominativo.strip()
                        elif current_nominativo:
                            nominativo = current_nominativo
                        
                        # Aggiungi solo se abbiamo dati significativi
                        if (data_fattura and numero_fattura and importo_liquidato and 
                            importo_liquidato != "0,00" and importo_liquidato.strip() != ""):
                            record = [
                                societa,
                                data_documento,
                                nominativo,
                                nominativo_familiare,
                                data_fattura,
                                numero_fattura,
                                importo_liquidato
                            ]
                            dati_estratti.append(record)
                   #         st.success(f"✅ Riga {i} aggiunta: {record}")
                   #     else:
                   #         st.warning(f"Riga {i} non ha dati significativi o importo = 0,00")
                else:
                    st.error(f"Nessun header trovato nel file {file_name}")
            
            st.info(f"📊 DEBUG: Totale righe estratte da {file_name}: {len(dati_estratti)}")
            return dati_estratti
            
    except Exception as e:
        st.error(f"Errore nell'elaborazione del file nuovo formato '{file_name}'")
        st.error(f"Dettaglio errore: {str(e)}")
        st.error(f"Traceback: {traceback.format_exc()}")
        return []

def estrai_pdf_da_zip(file_zip):
    pdf_files = []
    pdf_names = []
    try:
        with zipfile.ZipFile(file_zip, "r") as zip_ref:
            for file_name in zip_ref.namelist():
                if file_name.lower().endswith(".pdf"):
                    try:
                        with zip_ref.open(file_name) as pdf_file:
                            pdf_content = io.BytesIO(pdf_file.read())
                            pdf_files.append(pdf_content)
                            pdf_names.append(file_name)
                    except Exception as e:
                        st.error(f"Errore nell'estrazione del file PDF '{file_name}' dallo ZIP")
                        st.error(f"Dettaglio errore: {str(e)}")
    except Exception as e:
        st.error(f"Errore nell'apertura del file ZIP: {file_zip.name}")
        st.error(f"Dettaglio errore: {str(e)}")
    
    return pdf_files, pdf_names

def estrai_dati_da_pdf(lista_file_pdf, lista_nomi_pdf=None):
    if lista_nomi_pdf is None:
        lista_nomi_pdf = [f"File_{i+1}" for i in range(len(lista_file_pdf))]
    
    dati_completi = []
    file_con_errori = []
    
    for idx, file_pdf in enumerate(lista_file_pdf):
        file_name = lista_nomi_pdf[idx]
        try:
            s = estrai_testo_da_pdf_testata(file_pdf, file_name)
            dt = estrai_data_da_pdf_testata(file_pdf, file_name)
            
            with pdfplumber.open(file_pdf) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    try:
                        tables = page.extract_table()
                        if tables:
                            for i, row in enumerate(tables[1:], 1):
                                try:
                                    if len(tables[i]) == 9:
                                        a = tables[i][3]
                                        nf = tables[i][4]
                                        b = tables[i][5]
                                        c = tables[i][6]
                                        d = tables[i][8]
                                        if (not(c==None) and c!="" and not(d==None) and d!= "0,00" and d!=""):
                                            if ((a==None or a=='') and i>0):
                                                a=tables[i-1][3]
                                            if ((nf==None or nf=='') and i>0):
                                                nf=tables[i-1][4]
                                            if ((a==None or a=='') and i>1):
                                                a=tables[i-2][3]
                                            if ((nf==None or nf=='') and i>1):
                                                nf=tables[i-2][4]
                                            dati_completi.append([s,dt,a,nf,b,c,d])
                                    
                                    elif len(tables[i]) == 10:
                                        a = tables[i][2]
                                        nf = tables[i][3]
                                        b = tables[i][4]
                                        c = tables[i][5]
                                        d = tables[i][9]
                                        if (c and c!="" and d and d!= "0,00" and d!=""):
                                            if ((a==None or a=='') and i>0):
                                                a=tables[i-1][2]
                                            if ((nf==None or nf=='') and i>0):
                                                nf=tables[i-1][3]
                                            if ((a==None or a=='') and i>1):
                                                a=tables[i-2][2]
                                            if ((nf==None or nf=='') and i>1):
                                                nf=tables[i-2][3]
                                            dati_completi.append([s,dt,a,nf,b,c,d])
                                    
                                    elif len(tables[i]) == 11:
                                        a = tables[i][3]
                                        nf = tables[i][4]
                                        b = tables[i][5]
                                        c = tables[i][6]
                                        d = tables[i][10]
                                        if (not(c==None) and c!="" and not(d==None) and d!= "0,00" and d!=""):
                                            if ((a==None or a=='') and i>0):
                                                a=tables[i-1][3]
                                            if ((nf==None or nf=='') and i>0):
                                                nf=tables[i-1][4]
                                            if ((a==None or a=='') and i>1):
                                                a=tables[i-2][3]
                                            if ((nf==None or nf=='') and i>1):
                                                nf=tables[i-2][4]
                                            dati_completi.append([s,dt,a,nf,b,c,d])
                                    else:
                                        st.warning(f"Formato tabella non riconosciuto nel file '{file_name}', pagina {page_num}, riga {i}: {len(tables[i])} colonne")
                                
                                except Exception as e:
                                    st.warning(f"Errore nell'elaborazione della riga {i} nella tabella del file '{file_name}', pagina {page_num}")
                                    st.warning(f"Dettaglio errore: {str(e)}")
                    
                    except Exception as e:
                        st.error(f"Errore nell'elaborazione della pagina {page_num} del file '{file_name}'")
                        st.error(f"Dettaglio errore: {str(e)}")
                        file_con_errori.append((file_name, f"Errore pagina {page_num}: {str(e)}"))
        
        except Exception as e:
            st.error(f"Errore nell'elaborazione del file '{file_name}'")
            st.error(f"Dettaglio errore: {str(e)}")
            st.error(f"Traceback completo: {traceback.format_exc()}")
            file_con_errori.append((file_name, str(e)))

    colonne_selezionate = ["Società Testata", "Data Testata", "Nominativo Dirigente", 
                          "Nominativo Familiare", "Data Fattura", "Numero Fattura", "Totale Rimborsato"]
    
    df = pd.DataFrame(dati_completi, columns=colonne_selezionate)
    return df, file_con_errori

def estrai_dati_nuovo_formato(lista_file_pdf, lista_nomi_pdf=None):
    """
    Estrae dati dai PDF del nuovo formato
    """
    if lista_nomi_pdf is None:
        lista_nomi_pdf = [f"File_{i+1}" for i in range(len(lista_file_pdf))]
    
    dati_completi = []
    file_con_errori = []
    
    for idx, file_pdf in enumerate(lista_file_pdf):
        file_name = lista_nomi_pdf[idx]
        try:
            dati_file = estrai_dati_formato_nuovo(file_pdf, file_name)
            dati_completi.extend(dati_file)
        except Exception as e:
            st.error(f"Errore nell'elaborazione del file nuovo formato '{file_name}'")
            st.error(f"Dettaglio errore: {str(e)}")
            file_con_errori.append((file_name, str(e)))
    
    colonne_selezionate = ["Società Testata", "Data Testata", "Nominativo Dirigente", 
                          "Nominativo Familiare", "Data Fattura", "Numero Fattura", "Totale Rimborsato"]
    
    df = pd.DataFrame(dati_completi, columns=colonne_selezionate)
    return df, file_con_errori

st.title("Estrazione Tabelle da PDF")
st.info("Build 1.5.3 - 12/06/2025 - Supporto doppio formato")

# Creo due sezioni separate per i due tipi di file
col1, col2 = st.columns(2)

with col1:
    st.subheader("📊 Formato Fasi")
    st.caption("File PDF/ZIP con formato tabelle Fasi")
    file_originali = st.file_uploader("Carica i file PDF o ZIP (Fasi)", 
                                     type=["pdf","zip"], 
                                     accept_multiple_files=True,
                                     key="originali")

with col2:
    st.subheader("📋 Formato FasiOpen")
    st.caption("File PDF/ZIP con formato tabelle FasiOpen")
    file_nuovi = st.file_uploader("Carica i file PDF o ZIP (FasiOpen)", 
                                 type=["pdf","zip"], 
                                 accept_multiple_files=True,
                                 key="nuovi")

# Elaborazione file originali
df_originali = pd.DataFrame()
errori_originali = []

if file_originali:
    st.success(f"📊 {len(file_originali)} file Fasi!")
    
    with st.spinner("Elaborazione file formato Fasi..."):
        lista_pdf_orig = []
        nomi_pdf_orig = []
        
        for file in file_originali:
            if file.name.lower().endswith(".pdf"):
                lista_pdf_orig.append(file)
                nomi_pdf_orig.append(file.name)
            elif file.name.lower().endswith(".zip"):
                pdf_estratti, nomi_estratti = estrai_pdf_da_zip(file)
                lista_pdf_orig.extend(pdf_estratti)
                nomi_pdf_orig.extend(nomi_estratti)

        if lista_pdf_orig:
            df_originali, errori_originali = estrai_dati_da_pdf(lista_pdf_orig, nomi_pdf_orig)

# Elaborazione file nuovi
df_nuovi = pd.DataFrame()
errori_nuovi = []

if file_nuovi:
    st.success(f"📋 {len(file_nuovi)} file FasiOpen!")
    
    with st.spinner("Elaborazione file formato FasiOpen..."):
        lista_pdf_nuovi = []
        nomi_pdf_nuovi = []
        
        for file in file_nuovi:
            if file.name.lower().endswith(".pdf"):
                lista_pdf_nuovi.append(file)
                nomi_pdf_nuovi.append(file.name)
            elif file.name.lower().endswith(".zip"):
                pdf_estratti, nomi_estratti = estrai_pdf_da_zip(file)
                lista_pdf_nuovi.extend(pdf_estratti)
                nomi_pdf_nuovi.extend(nomi_estratti)

        if lista_pdf_nuovi:
            df_nuovi, errori_nuovi = estrai_dati_nuovo_formato(lista_pdf_nuovi, nomi_pdf_nuovi)

# Combina i risultati se entrambi i tipi di file sono stati elaborati
if not df_originali.empty or not df_nuovi.empty:
    st.write("---")
    st.write("## 📊 Risultati Elaborazione")
    
    # Combina i DataFrame
    dataframes_da_combinare = []
    if not df_originali.empty:
        df_originali['Tipo_Formato'] = 'Originale'
        dataframes_da_combinare.append(df_originali)
    if not df_nuovi.empty:
        df_nuovi['Tipo_Formato'] = 'Nuovo'
        dataframes_da_combinare.append(df_nuovi)
    
    if dataframes_da_combinare:
        df_finale = pd.concat(dataframes_da_combinare, ignore_index=True)
        
        # Mostra statistiche separate
        col_stat1, col_stat2 = st.columns(2)
        
        with col_stat1:
            if not df_originali.empty:
                st.metric("Righe Formato Fasi", len(df_originali))
        
        with col_stat2:
            if not df_nuovi.empty:
                st.metric("Righe Formato FasiOpen", len(df_nuovi))
        
        st.write("### Statistiche Totali")
        st.write(f"- Numero totale di righe estratte: {len(df_finale)}")
        st.write(f"- Numero di società diverse: {df_finale['Società Testata'].nunique()}")
        
        # Mostra gli errori se ce ne sono
        tutti_errori = errori_originali + errori_nuovi
        if tutti_errori:
            st.warning(f"Elaborazione completata con {len(tutti_errori)} file problematici")
            
            with st.expander("Mostra dettagli degli errori"):
                for file_name, error in tutti_errori:
                    st.error(f"File: {file_name} - Errore: {error}")
        else:
            st.success("Elaborazione completata con successo per tutti i file!")
        
        # Anteprima dati
        st.write("### Anteprima Dati Estratti")
        st.dataframe(df_finale.head(10))
        
        # Prepara il file Excel con fogli separati
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            # Foglio principale con tutti i dati
            df_finale.to_excel(writer, index=False, sheet_name="Tutti_i_Rimborsi")
            
            # Fogli separati per formato
            if not df_originali.empty:
                df_originali.drop('Tipo_Formato', axis=1).to_excel(writer, index=False, sheet_name="Formato_Originale")
            
            if not df_nuovi.empty:
                df_nuovi.drop('Tipo_Formato', axis=1).to_excel(writer, index=False, sheet_name="Nuovo_Formato")
            
            # Foglio errori se presenti
            if tutti_errori:
                df_errori = pd.DataFrame(tutti_errori, columns=["Nome File", "Errore"])
                df_errori.to_excel(writer, index=False, sheet_name="Errori")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"estratto_fasi_open_completo_{timestamp}.xlsx"

        st.download_button(
            label="📥 Scarica il file Excel completo",
            data=buffer.getvalue(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Nessuna tabella estratta dai PDF! Controlla i messaggi di errore sopra.")

# Messaggio informativo se nessun file è stato caricato
if not file_originali and not file_nuovi:
    st.info("👆 Carica i file PDF o ZIP utilizzando i pulsanti sopra per iniziare l'elaborazione.")
    
    with st.expander("ℹ️ Informazioni sui formati supportati"):
        st.write("""
        **Formati**: File PDF Fasi con tabelle complesse multi-colonna (9, 10, 11 colonne) e File PDF FasiOpen.
        
        Lo script può elaborare entrambi i formati contemporaneamente e creerà un file Excel con fogli separati per ogni formato.
        """)


st.title("Riconciliazione Incassi")

# Caricamento file
incassi_file = st.file_uploader("Carica il file 'Incassi_da_allocare' (.xls o .xlsx)", type=["xls", "xlsx"])
dettagli_file = st.file_uploader("Carica il file 'Dettaglio_pagamenti' (.xls o .xlsx)", type=["xls", "xlsx"])

if incassi_file and dettagli_file:
    # Lettura dei file Excel
    df_incassi = pd.read_excel(incassi_file, dtype=str)
    df_dettagli = pd.read_excel(dettagli_file, dtype=str)

    st.subheader("Anteprima - Incassi da allocare")
    st.write(df_incassi.head())

    st.subheader("Anteprima - Dettaglio pagamenti")
    st.write(df_dettagli.head())

    # Estrazione SEQ da testo (tutte le celle)
    text_data = df_incassi.astype(str).values.ravel()
    seq_matches = []
    for cell in text_data:
        found = re.findall(r"(?i)seq\s*[:.]\s*(\d+)", cell)
        seq_matches.extend(found)

    seq_set = set(seq_matches)

    st.info(f"Trovati {len(seq_set)} SEQ unici nel file Incassi da Allocare")

    def estrai_seq_numerico(valore):
        if pd.isna(valore):
            return None
        match = re.search(r"\d+", str(valore))
        return match.group(0) if match else None

    # Filtro dei dettagli pagamenti
    if 'Seq' in df_dettagli.columns:
        df_dettagli['SEQ_PULITO'] = df_dettagli['Seq'].apply(estrai_seq_numerico)

        df_filtrato = df_dettagli[df_dettagli['SEQ_PULITO'].isin(seq_set)]

        st.subheader("Righe filtrate dal Dettaglio Pagamenti")

        # Download Excel
        output = BytesIO()
        df_filtrato.drop(columns=["SEQ_PULITO"]).to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="📥 Scarica risultati filtrati",
            data=output,
            file_name="dettagli_filtrati.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("❌ La colonna 'SEQ' non è presente nel file 'Dettaglio_pagamenti.xls'")
