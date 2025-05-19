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

st.title("Estrazione Tabelle da PDF FasiOpen")
st.info("Build 1.4.1 - 19/05/2025")

file_caricati = st.file_uploader("Carica i file PDF o ZIP", type=["pdf","zip"], accept_multiple_files=True)

if file_caricati:
    st.success(f"{len(file_caricati)} file caricati con successo!")
    
    with st.spinner("Elaborazione dei file in corso..."):
        lista_pdf = []
        nomi_pdf = []
        
        # Estrai tutti i PDF
        for file in file_caricati:
            if file.name.lower().endswith(".pdf"):
                lista_pdf.append(file)
                nomi_pdf.append(file.name)
            elif file.name.lower().endswith(".zip"):
                pdf_estratti, nomi_estratti = estrai_pdf_da_zip(file)
                lista_pdf.extend(pdf_estratti)
                nomi_pdf.extend(nomi_estratti)

        if lista_pdf:
            st.info(f"Totale file PDF da elaborare: {len(lista_pdf)}")
            
            # Mostra una barra di progresso
            progress_bar = st.progress(0)
            
            # Elabora i PDF in piccoli gruppi per aggiornare la barra di progresso
            chunk_size = max(1, len(lista_pdf) // 10)
            all_data = []
            all_errors = []
            
            for i in range(0, len(lista_pdf), chunk_size):
                end_idx = min(i + chunk_size, len(lista_pdf))
                chunk_pdf = lista_pdf[i:end_idx]
                chunk_names = nomi_pdf[i:end_idx]
                
                df_chunk, errori_chunk = estrai_dati_da_pdf(chunk_pdf, chunk_names)
                all_data.append(df_chunk)
                all_errors.extend(errori_chunk)
                
                # Aggiorna la barra di progresso
                progress_bar.progress(end_idx / len(lista_pdf))
            
            # Combina tutti i dati
            df_finale = pd.concat(all_data) if all_data else pd.DataFrame()
            
            # Nascondi la barra di progresso dopo il completamento
            progress_bar.empty()
            
            # Mostra un riepilogo degli errori
            if all_errors:
                st.warning(f"Elaborazione completata con {len(all_errors)} file problematici")
                
                with st.expander("Mostra dettagli degli errori"):
                    for file_name, error in all_errors:
                        st.error(f"File: {file_name} - Errore: {error}")
            else:
                st.success("Elaborazione completata con successo per tutti i file!")

            if not df_finale.empty:
                st.write("### Anteprima Dati Estratti")
                st.dataframe(df_finale.head())
                
                # Aggiungi statistiche
                st.write("### Statistiche")
                st.write(f"- Numero totale di righe estratte: {len(df_finale)}")
                st.write(f"- Numero di società diverse: {df_finale['Società Testata'].nunique()}")
                
                # Prepara il file Excel
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_finale.to_excel(writer, index=False, sheet_name="Rimborsi")
                    
                    # Aggiungi un foglio con gli errori se ce ne sono
                    if all_errors:
                        df_errori = pd.DataFrame(all_errors, columns=["Nome File", "Errore"])
                        df_errori.to_excel(writer, index=False, sheet_name="Errori")

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                file_name = f"estratto_fasi_open_{timestamp}.xlsx"

                st.download_button(
                    label="📥 Scarica il file Excel",
                    data=buffer.getvalue(),
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Nessuna tabella estratta dai PDF! Controlla i messaggi di errore sopra.")
        else:
            st.warning("Nessun file PDF trovato nei file caricati.")
