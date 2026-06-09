import streamlit as st
import pdfplumber
import fitz  # PyMuPDF - per estrazione testo con spazi corretti
import pandas as pd
import io
import re
import zipfile
from datetime import datetime
import traceback
import os
from io import BytesIO
from pathlib import Path

# Configurazione pagina - DEVE essere la prima chiamata Streamlit
st.set_page_config(
    page_title="Fondi Manager",
    page_icon="🏥",
    layout="wide"
)


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
    Usa PyMuPDF (fitz) per estrazione testo con spazi corretti
    Analizza la seconda pagina per estrarre i dati dalla tabella
    Restituisce una tupla (dati_estratti, importo_distinta)
    """
    try:
        if hasattr(file_pdf, 'seek'):
            file_pdf.seek(0)
        file_bytes = file_pdf.read()

        doc = fitz.open(stream=file_bytes, filetype="pdf")

        # Estrai informazioni dalla prima pagina con PyMuPDF
        first_page_text = doc[0].get_text()

        # Estrai la società destinataria (in alto a destra nella prima pagina)
        societa = "no_societa"

        # Pattern per trovare nomi di società: parole maiuscole seguite da SRL/SPA/etc
        societa_pattern = re.findall(r'([A-Z][A-Z0-9\s\.\'\-]+(?:S\.?R\.?L\.?|SRL|S\.?P\.?A\.?|SPA|S\.?N\.?C\.?|SNC|S\.?A\.?S\.?|SAS))', first_page_text)

        for match in societa_pattern:
            match_clean = match.strip()
            if "FasiOpen" in match_clean or "Fondo" in match_clean or "Assistenza" in match_clean:
                continue
            societa_raw = re.split(r'\s*-\s*GRUPPO', match_clean)[0].strip()
            if societa_raw:
                societa = societa_raw
                break

        # Fallback: cerca nomi società noti senza suffisso legale (es. "DP DENT")
        if societa == "no_societa":
            if re.search(r'\bDP\s*DENT\b', first_page_text, re.IGNORECASE):
                societa = "DP DENT"

        # Estrai la data valuta dal testo centrale (pattern "con valuta DD/MM/YYYY")
        testo_normalizzato = first_page_text.replace('\n', ' ')
        data_documento = "no_data"
        valuta_match = re.search(r"con\s*valuta\s*(\d{2}/\d{2}/\d{4})", testo_normalizzato, re.IGNORECASE)
        if valuta_match:
            data_documento = valuta_match.group(1)

        # Analizza la seconda pagina per i dati della tabella
        if len(doc) < 2:
            st.warning(f"Il file {file_name} non ha una seconda pagina")
            doc.close()
            return [], 0.0

        second_page = doc[1]
        tabs = second_page.find_tables()

        dati_estratti = []
        totale_fatture_file = 0.0

        if tabs.tables:
            tables = tabs[0].extract()  # lista di liste, come pdfplumber

            # Trova l'indice della riga di intestazione
            header_row_idx = None
            for i, row in enumerate(tables):
                if row and any(cell and ("Iscritto Principale" in str(cell) or "Main Client" in str(cell)) for cell in row):
                    header_row_idx = i
                    break

            if header_row_idx is None:
                for i, row in enumerate(tables):
                    if row and any(cell and ("FasiOpen" in str(cell) or "Nominativo" in str(cell)) for cell in row if cell):
                        header_row_idx = i
                        break

            if header_row_idx is not None:
                current_nominativo = None

                for i in range(header_row_idx + 1, len(tables)):
                    row = tables[i]
                    if not row:
                        continue

                    if len(row) < 6:
                        continue

                    # Salta le righe "Totale"
                    if row and any(cell and "Totale" in str(cell) for cell in row if cell):
                        continue

                    # Estrai i dati dalla riga
                    cod_fasiopen = row[0] if len(row) > 0 and row[0] else ""
                    nominativo_raw = row[1] if len(row) > 1 and row[1] else ""  # Iscritto Principale
                    nominativo_familiare_raw = row[2] if len(row) > 2 and row[2] else ""  # Nominativo Familiare
                    data_fattura = row[3] if len(row) > 3 and row[3] else ""
                    numero_fattura = row[4] if len(row) > 4 and row[4] else ""
                    importo_fattura = row[5] if len(row) > 5 and row[5] else ""  # Importo Fattura
                    importo_liquidato = row[6] if len(row) > 6 and row[6] else ""  # Totale Rimborsato

                    # Separa cognome e nome con spazio (rimuovi newline)
                    nominativo = nominativo_raw.replace('\n', ' ').strip() if nominativo_raw else ""
                    nominativo_familiare = nominativo_familiare_raw.replace('\n', ' ').strip() if nominativo_familiare_raw else ""

                    # Se il nominativo è vuoto, usa quello precedente
                    if nominativo and nominativo.strip():
                        current_nominativo = nominativo.strip()
                    elif current_nominativo:
                        nominativo = current_nominativo

                    # Usa il familiare se presente, altrimenti l'iscritto principale
                    paziente = nominativo_familiare if nominativo_familiare else nominativo

                    # Includi solo se fattura e rimborso sono diversi da zero
                    if (data_fattura and numero_fattura and
                        importo_fattura and importo_fattura.strip() and importo_fattura != "0,00" and
                        importo_liquidato and importo_liquidato.strip() and importo_liquidato != "0,00"):
                        record = [
                            societa,
                            data_documento,
                            paziente,
                            data_fattura,
                            numero_fattura,
                            importo_liquidato
                        ]
                        dati_estratti.append(record)
                        try:
                            importo_str = importo_liquidato.replace(' ', '').replace('.', '').replace(',', '.')
                            totale_fatture_file += float(importo_str)
                        except:
                            pass
            else:
                st.error(f"Nessun header trovato nel file {file_name}")

        doc.close()
        return dati_estratti, totale_fatture_file

    except Exception as e:
        st.error(f"Errore nell'elaborazione del file nuovo formato '{file_name}'")
        st.error(f"Dettaglio errore: {str(e)}")
        st.error(f"Traceback: {traceback.format_exc()}")
        return [], 0.0

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
            # Reset posizione file prima di ogni lettura
            if hasattr(file_pdf, 'seek'):
                file_pdf.seek(0)
            s = estrai_testo_da_pdf_testata(file_pdf, file_name)

            if hasattr(file_pdf, 'seek'):
                file_pdf.seek(0)
            dt = estrai_data_da_pdf_testata(file_pdf, file_name)

            # Lista per raccogliere i dati di questo file e calcolare l'importo distinta
            dati_file = []
            totale_fatture_file = 0.0

            # Reset posizione file prima della lettura principale
            if hasattr(file_pdf, 'seek'):
                file_pdf.seek(0)

           # st.info(f"DEBUG FASI: Elaborazione file {file_name}, testata={s}, data={dt}")

            with pdfplumber.open(file_pdf) as pdf:
                #st.info(f"DEBUG FASI: File {file_name} ha {len(pdf.pages)} pagine")
                for page_num, page in enumerate(pdf.pages, 1):
                    try:
                        tables = page.extract_table()
                        #st.info(f"DEBUG FASI: Pagina {page_num}, tabella trovata: {tables is not None}, righe: {len(tables) if tables else 0}")
                        if tables:
                            # Debug: mostra struttura prime righe
                            #for debug_i, debug_row in enumerate(tables[:3]):
                                #st.info(f"DEBUG FASI: Riga {debug_i} ha {len(debug_row)} colonne: {debug_row}")

                            for i, row in enumerate(tables[1:], 1):
                                try:
                                    #st.info(f"DEBUG FASI: Riga {i} - colonne: {len(tables[i])}")
                                    if len(tables[i]) == 9:
                                        a = tables[i][3]  # Nominativo Dirigente
                                        nf = tables[i][4]  # Nominativo Familiare
                                        b = tables[i][5]  # Data Fattura
                                        c = tables[i][6]  # Numero Fattura
                                        tot_fatt = tables[i][7]  # Totale Fattura
                                        d = tables[i][8]  # Totale Rimborsato
                                        # Includi solo se fattura e rimborso sono diversi da zero
                                        if (c and c.strip() and
                                            tot_fatt and tot_fatt.strip() and tot_fatt != "0,00" and
                                            d and d.strip() and d != "0,00"):
                                            if ((a==None or a=='') and i>0):
                                                a=tables[i-1][3]
                                            if ((a==None or a=='') and i>1):
                                                a=tables[i-2][3]
                                            paziente = nf if (nf and nf.strip()) else a
                                            dati_file.append([s, dt, paziente, b, c, d])
                                            try:
                                                totale_fatture_file += float(d.replace(' ', '').replace('.', '').replace(',', '.'))
                                            except:
                                                pass

                                    elif len(tables[i]) == 10:
                                        a = tables[i][2]  # Nominativo Dirigente
                                        nf = tables[i][3]  # Nominativo Familiare
                                        b = tables[i][4]  # Data Fattura
                                        c = tables[i][5]  # Numero Fattura
                                        tot_fatt = tables[i][8]  # Totale Fattura
                                        d = tables[i][9]  # Totale Rimborsato
                                        # Includi solo se fattura e rimborso sono diversi da zero
                                        if (c and c.strip() and
                                            tot_fatt and tot_fatt.strip() and tot_fatt != "0,00" and
                                            d and d.strip() and d != "0,00"):
                                            if ((a==None or a=='') and i>0):
                                                a=tables[i-1][2]
                                            if ((a==None or a=='') and i>1):
                                                a=tables[i-2][2]
                                            paziente = nf if (nf and nf.strip()) else a
                                            dati_file.append([s, dt, paziente, b, c, d])
                                            try:
                                                totale_fatture_file += float(d.replace(' ', '').replace('.', '').replace(',', '.'))
                                            except:
                                                pass

                                    elif len(tables[i]) == 11:
                                        a = tables[i][3]  # Nominativo Dirigente
                                        nf = tables[i][4]  # Nominativo Familiare
                                        b = tables[i][5]  # Data Fattura
                                        c = tables[i][6]  # Numero Fattura
                                        tot_fatt = tables[i][8]  # Totale Fattura (indice corretto)
                                        d = tables[i][10]  # Totale Rimborsato
                                        # Includi solo se fattura e rimborso sono diversi da zero
                                        if (c and c.strip() and
                                            tot_fatt and tot_fatt.strip() and tot_fatt != "0,00" and
                                            d and d.strip() and d != "0,00"):
                                            if ((a==None or a=='') and i>0):
                                                a=tables[i-1][3]
                                            if ((a==None or a=='') and i>1):
                                                a=tables[i-2][3]
                                            paziente = nf if (nf and nf.strip()) else a
                                            dati_file.append([s, dt, paziente, b, c, d])
                                            try:
                                                totale_fatture_file += float(d.replace(' ', '').replace('.', '').replace(',', '.'))
                                            except:
                                                pass
                                    else:
                                        st.warning(f"Formato tabella non riconosciuto nel file '{file_name}', pagina {page_num}, riga {i}: {len(tables[i])} colonne")

                                except Exception as e:
                                    st.warning(f"Errore nell'elaborazione della riga {i} nella tabella del file '{file_name}', pagina {page_num}")
                                    st.warning(f"Dettaglio errore: {str(e)}")

                    except Exception as e:
                        st.error(f"Errore nell'elaborazione della pagina {page_num} del file '{file_name}'")
                        st.error(f"Dettaglio errore: {str(e)}")
                        file_con_errori.append((file_name, f"Errore pagina {page_num}: {str(e)}"))

            # Formatta l'importo distinta per questo file
            importo_distinta_str = f"{totale_fatture_file:.2f}".replace('.', ',') if totale_fatture_file > 0 else ""

            # Aggiungi l'importo distinta a ogni riga del file
            for riga in dati_file:
                # Inserisci importo distinta dopo società e data: [s, importo, dt, paziente, b, c, d]
                dati_completi.append([riga[0], importo_distinta_str, riga[1], riga[2], riga[3], riga[4], riga[5], file_name])

        except Exception as e:
            st.error(f"Errore nell'elaborazione del file '{file_name}'")
            st.error(f"Dettaglio errore: {str(e)}")
            st.error(f"Traceback completo: {traceback.format_exc()}")
            file_con_errori.append((file_name, str(e)))

    colonne_selezionate = ["Società Testata", "Importo Distinta", "Data Testata", "Nominativo Dirigente",
                           "Data Fattura", "Numero Fattura", "Totale Rimborsato", "Nome File PDF"]

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
            dati_file, totale_fatture_file = estrai_dati_formato_nuovo(file_pdf, file_name)

            # Formatta l'importo distinta per questo file
            importo_distinta_str = f"{totale_fatture_file:.2f}".replace('.', ',') if totale_fatture_file > 0 else ""

            # Aggiungi l'importo distinta a ogni riga del file
            # dati_file contiene: [societa, data_documento, paziente, data_fattura, numero_fattura, importo_liquidato]
            for riga in dati_file:
                # Inserisci importo distinta: [societa, importo, data, paziente, data_fatt, num_fatt, tot_rimb]
                dati_completi.append([riga[0], importo_distinta_str, riga[1], riga[2], riga[3], riga[4], riga[5], file_name])
        except Exception as e:
            st.error(f"Errore nell'elaborazione del file nuovo formato '{file_name}'")
            st.error(f"Dettaglio errore: {str(e)}")
            file_con_errori.append((file_name, str(e)))

    colonne_selezionate = ["Società Testata", "Importo Distinta", "Data Testata",
                          "Nominativo Dirigente", "Data Fattura", "Numero Fattura", "Totale Rimborsato", "Nome File PDF"]

    df = pd.DataFrame(dati_completi, columns=colonne_selezionate)

    return df, file_con_errori


def estrai_riferimento_ri3(n_assegno):
    """Extract TT ticket numbers and invoice numbers from the N. assegno RI3 section."""
    if pd.isna(n_assegno):
        return [], []
    text = str(n_assegno)

    ri3_match = re.search(r':RI3:(.*?)(?::SEC:|$)', text, re.DOTALL | re.IGNORECASE)
    ri3_text = ri3_match.group(1).strip() if ri3_match else text

    tt_numbers = [t.upper() for t in re.findall(r'TT-\d+', ri3_text, re.IGNORECASE)]

    fattura_numbers = []

    # Pattern: FATTURA: NNN/NNN or FATTURAN NNN/NNN (explicit label)
    explicit = re.findall(r'FATTURA\w*\s*:?\s*(\d[\d/]*)', ri3_text, re.IGNORECASE)
    fattura_numbers.extend(explicit)

    # Pattern: token immediately after TT-XXXXXXXX (implicit invoice reference)
    after_tt = re.sub(r'.*?TT-\d+\s*', '', ri3_text, count=1).strip()
    if after_tt:
        m = re.match(r'(\d[\d/]*)', after_tt)
        if m:
            val = m.group(1)
            if val not in fattura_numbers:
                fattura_numbers.append(val)

    fattura_numbers = list(dict.fromkeys(fattura_numbers))
    fattura_numbers = [f for f in fattura_numbers if f and not re.match(r'^0{4,}', f)]

    return tt_numbers, fattura_numbers


def riconcilia_incassi_aon(df_fatture, df_incassi):
    """
    Reconcile AON incassi with fatture.
    Adds columns: Rif. Estratto, Fatture Riconciliate, Metodo Match.
    Matching priority: TT ticket number first, then Num Fattura.
    """
    tt_index = {}
    num_index = {}

    for _, row in df_fatture.iterrows():
        tt = str(row.get('Master Ticket CRM', '')).strip().upper()
        num = str(row.get('Num Fattura', '')).strip()
        paziente = str(row.get('PAZIENTE', '')).strip()
        importo = str(row.get('Importo Liquidato', '')).strip()
        clinica = str(row.get('Nome Ric EE Fattura', '')).strip()

        entry = {'num_fattura': num, 'paziente': paziente, 'importo': importo, 'clinica': clinica}

        if tt and tt.lower() != 'nan':
            tt_index.setdefault(tt, []).append(entry)

        if num and num.lower() != 'nan':
            norm = num.replace('.', '/').replace(' ', '')
            num_index.setdefault(norm, []).append({**entry, 'tt': tt})

    fatture_col = []
    pazienti_col = []
    metodo_col = []
    rif_col = []

    for _, row in df_incassi.iterrows():
        n_ass = row.get('N. assegno', '')
        tt_list, fatt_list = estrai_riferimento_ri3(n_ass)

        rif_parts = list(tt_list)
        for f in fatt_list:
            if f not in rif_parts:
                rif_parts.append(f)
        rif_col.append('; '.join(rif_parts[:3]) if rif_parts else '')

        matched = []
        method = 'Non trovato'

        for tt in tt_list:
            if tt in tt_index:
                matched.extend(tt_index[tt])
                method = 'TT'

        if not matched:
            for fatt_num in fatt_list:
                norm = fatt_num.replace('.', '/').replace(' ', '')
                if norm in num_index:
                    matched.extend(num_index[norm])
                    method = 'Num Fattura'

        if matched:
            seen = set()
            unique = []
            for m in matched:
                key = m['num_fattura']
                if key not in seen:
                    seen.add(key)
                    unique.append(m)

            fatture_col.append(' | '.join(m['num_fattura'] for m in unique))
            pazienti_col.append(' | '.join(
                m['paziente'] for m in unique
                if m['paziente'] and m['paziente'].lower() != 'nan'
            ))
        else:
            fatture_col.append('')
            pazienti_col.append('')

        metodo_col.append(method)

    df_result = df_incassi.copy()
    df_result['Rif. Estratto'] = rif_col
    df_result['Assistito'] = pazienti_col
    df_result['Fatture Riconciliate'] = fatture_col
    df_result['Metodo Match'] = metodo_col

    return df_result


st.title("Estrazione Tabelle da PDF")
st.info("Build 1.8.0 - 09/06/2026 - Aggiunta sezione Riconciliazione Incassi AON")

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

        st.info(f"DEBUG: {len(lista_pdf_orig)} PDF da elaborare")
        if lista_pdf_orig:
            df_originali, errori_originali = estrai_dati_da_pdf(lista_pdf_orig, nomi_pdf_orig)
            st.info(f"DEBUG: DataFrame risultante ha {len(df_originali)} righe")
            if errori_originali:
                st.warning(f"DEBUG: {len(errori_originali)} errori: {errori_originali}")

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

        # Riordina le colonne nell'ordine desiderato
        colonne_ordinate = ['Società Testata', 'Importo Distinta', 'Data Testata',
                           'Nominativo Dirigente', 'Data Fattura', 'Numero Fattura',
                           'Totale Rimborsato', 'Nome File PDF', 'Tipo_Formato']
        # Seleziona solo le colonne che esistono nel DataFrame
        colonne_finali = [col for col in colonne_ordinate if col in df_finale.columns]
        df_finale = df_finale[colonne_finali]

        # Riordina anche i DataFrame originali per i fogli separati
        if not df_originali.empty:
            colonne_orig = [col for col in colonne_ordinate if col in df_originali.columns]
            df_originali = df_originali[colonne_orig]
        if not df_nuovi.empty:
            colonne_nuovi = [col for col in colonne_ordinate if col in df_nuovi.columns]
            df_nuovi = df_nuovi[colonne_nuovi]

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
        
        # Genera la data di elaborazione
        data_elaborazione = datetime.now().strftime("%Y%m%d_%H%M%S")

        # Crea file Excel separato per FASI (formato originale)
        if not df_originali.empty:
            buffer_fasi = io.BytesIO()
            with pd.ExcelWriter(buffer_fasi, engine="openpyxl") as writer:
                df_orig_export = df_originali.drop('Tipo_Formato', axis=1, errors='ignore')
                # Aggiungi colonne Fattura Originale e Clinica dallo split di Numero Fattura per "\"
                df_orig_export['Fattura Originale'] = df_orig_export['Numero Fattura'].apply(
                    lambda x: str(x).split('/')[0].strip() if pd.notna(x) and '/' in str(x) else str(x) if pd.notna(x) else '')
                df_orig_export['Clinica'] = df_orig_export['Numero Fattura'].apply(
                    lambda x: str(x).split('/')[1].strip() if pd.notna(x) and '/' in str(x) else '')
                df_orig_export.to_excel(writer, index=False, sheet_name="Rimborsi_FASI")

                # Aggiungi foglio errori se presenti errori per FASI
                if errori_originali:
                    df_errori_fasi = pd.DataFrame(errori_originali, columns=["Nome File", "Errore"])
                    df_errori_fasi.to_excel(writer, index=False, sheet_name="Errori")

            file_name_fasi = f"FASI_{data_elaborazione}.xlsx"
            # Rimuovi caratteri di ritorno a capo dai campi di testo per evitare rottura del formato CSV
            df_orig_csv = df_orig_export.copy()
            df_orig_csv = df_orig_csv.apply(lambda col: col.map(lambda x: str(x).replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ') if pd.notna(x) else x))
            csv_fasi = df_orig_csv.to_csv(index=False, header=False, sep=';').encode('utf-8')
            file_name_fasi_csv = f"FASI_{data_elaborazione}.csv"
            col_fasi_xl, col_fasi_csv = st.columns(2)
            with col_fasi_xl:
                st.download_button(
                    label="📥 Scarica Excel FASI",
                    data=buffer_fasi.getvalue(),
                    file_name=file_name_fasi,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_fasi"
                )
            with col_fasi_csv:
                st.download_button(
                    label="📥 Scarica CSV FASI",
                    data=csv_fasi,
                    file_name=file_name_fasi_csv,
                    mime="text/csv",
                    key="download_fasi_csv"
                )

        # Crea file Excel separato per FASIOPEN (nuovo formato)
        if not df_nuovi.empty:
            buffer_fasiopen = io.BytesIO()
            with pd.ExcelWriter(buffer_fasiopen, engine="openpyxl") as writer:
                df_nuovi_export = df_nuovi.drop('Tipo_Formato', axis=1, errors='ignore')
                # Aggiungi colonne Fattura Originale e Clinica dallo split di Numero Fattura per "\"
                df_nuovi_export['Fattura Originale'] = df_nuovi_export['Numero Fattura'].apply(
                    lambda x: str(x).split('/')[0].strip() if pd.notna(x) and '/' in str(x) else str(x) if pd.notna(x) else '')
                df_nuovi_export['Clinica'] = df_nuovi_export['Numero Fattura'].apply(
                    lambda x: str(x).split('/')[1].strip() if pd.notna(x) and '/' in str(x) else '')
                df_nuovi_export.to_excel(writer, index=False, sheet_name="Rimborsi_FASIOPEN")

                # Aggiungi foglio errori se presenti errori per FASIOPEN
                if errori_nuovi:
                    df_errori_fasiopen = pd.DataFrame(errori_nuovi, columns=["Nome File", "Errore"])
                    df_errori_fasiopen.to_excel(writer, index=False, sheet_name="Errori")

            file_name_fasiopen = f"FASIOPEN_{data_elaborazione}.xlsx"
            # Rimuovi caratteri di ritorno a capo dai campi di testo per evitare rottura del formato CSV
            df_nuovi_csv = df_nuovi_export.copy()
            df_nuovi_csv = df_nuovi_csv.apply(lambda col: col.map(lambda x: str(x).replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ') if pd.notna(x) else x))
            csv_fasiopen = df_nuovi_csv.to_csv(index=False, header=False, sep=';').encode('utf-8')
            file_name_fasiopen_csv = f"FASIOPEN_{data_elaborazione}.csv"
            col_fo_xl, col_fo_csv = st.columns(2)
            with col_fo_xl:
                st.download_button(
                    label="📥 Scarica Excel FASIOPEN",
                    data=buffer_fasiopen.getvalue(),
                    file_name=file_name_fasiopen,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_fasiopen"
                )
            with col_fo_csv:
                st.download_button(
                    label="📥 Scarica CSV FASIOPEN",
                    data=csv_fasiopen,
                    file_name=file_name_fasiopen_csv,
                    mime="text/csv",
                    key="download_fasiopen_csv"
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

    # Trova la colonna "data operazione" nel file incassi (case-insensitive)
    col_data_op = None
    for col in df_incassi.columns:
        if 'data operazione' in str(col).strip().lower():
            col_data_op = col
            break

    # Estrazione SEQ da testo riga per riga, con associazione alla data operazione
    seq_matches = []
    seq_to_data_pagamento = {}
    for idx, row in df_incassi.iterrows():
        if col_data_op and pd.notna(row.get(col_data_op)):
            val = str(row[col_data_op]).strip()
            try:
                # Excel salva le date come numeri seriali: convertiamo in data
                serial = float(val)
                data_op = (pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(serial))).strftime('%d/%m/%Y')
            except ValueError:
                # Già una stringa data leggibile
                data_op = val
        else:
            data_op = ""
        for cell in row.values:
            found = re.findall(r"(?i)seq\s*[:.]\s*(\d+)", str(cell))
            for seq in found:
                seq_matches.append(seq)
                if seq not in seq_to_data_pagamento and data_op:
                    seq_to_data_pagamento[seq] = data_op

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

        df_filtrato = df_dettagli[df_dettagli['SEQ_PULITO'].isin(seq_set)].copy()

        # Aggiungi colonna "Data Pagamento" dalla "data operazione" del file incassi
        df_filtrato['Data Pagamento'] = df_filtrato['SEQ_PULITO'].map(seq_to_data_pagamento).fillna('')

        st.subheader("Righe filtrate da Dettaglio Pagamenti")
        st.write(df_filtrato.drop(columns=["SEQ_PULITO"]))

        # Download file
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
        st.error("❌ La colonna 'Seq' non è presente nel file 'Dettaglio_pagamenti'")


def validate_clinica_file(df):
    """
    Valida che il dataframe contenga le colonne attese per i dati delle cliniche.
    """
    expected_columns = [
        "CLINICA", "AM", "DM", "TIPOLOGIA", "TIPOLOGIA INCASSO", 
        "DATA OPERAZIONE", "DATA VALUTA", "IMPORTO", "DESCRIZIONE", 
        "NUMERO INT CODE", "PREVENTIVO", "CODICE FINANZIAMENTO", 
        "RAGIONE SOCIALE FINANZIARIA"
    ]
    
    # Verifica presenza colonna CLINICA
    if "CLINICA" not in df.columns:
        return False, "Colonna 'CLINICA' non trovata!"
    
    return True, "File valido"

def split_excel_by_clinica(df):
    """
    Divide il dataframe in base ai valori della colonna CLINICA.
    
    Returns:
        dict: Dictionary con chiave=nome_clinica, valore=dataframe_clinica
    """
    cliniche_data = {}
    
    # Raggruppa per clinica
    grouped = df.groupby("CLINICA")
    
    for clinica_name, group in grouped:
        # Pulisci il nome della clinica per renderlo sicuro come nome file
        safe_name = str(clinica_name).replace("/", "_").replace("\\", "_").replace(":", "_").replace("*", "_").replace("?", "_").replace("\"", "_").replace("<", "_").replace(">", "_").replace("|", "_")
        cliniche_data[safe_name] = group
    
    return cliniche_data

def create_zip_file(cliniche_data):
    """
    Crea un file ZIP contenente tutti i file Excel delle cliniche.
    """
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for clinica_name, df_clinica in cliniche_data.items():
            # Crea il file Excel in memoria
            excel_buffer = io.BytesIO()
            df_clinica.to_excel(excel_buffer, index=False, engine='openpyxl')
            excel_buffer.seek(0)
            
            # Aggiungi al ZIP
            zip_file.writestr(f"{clinica_name}.xlsx", excel_buffer.read())
    
    zip_buffer.seek(0)
    return zip_buffer

def main():
    # Titolo dell'app
    st.title("🏥 Divisore Excel per Cliniche")
    st.markdown("---")
    
    # Descrizione
    st.markdown("""
    ### 📋 Questa app divide un file Excel in base ai valori della colonna CLINICA
    
    **Funzionalità:**
    - Carica un file Excel con dati delle cliniche
    - Divide automaticamente i dati per ogni clinica
    - Scarica tutti i file in un archivio ZIP
    """)
    
    # Sidebar per informazioni
    with st.sidebar:
        st.header("ℹ️ Informazioni")
        st.markdown("""
        **Colonne richieste:**
        - CLINICA (obbligatoria)
        - AM, DM, TIPOLOGIA
        - TIPOLOGIA INCASSO
        - DATA OPERAZIONE
        - DATA VALUTA
        - IMPORTO, DESCRIZIONE
        - NUMERO INT CODE
        - PREVENTIVO
        - CODICE FINANZIAMENTO
        - RAGIONE SOCIALE FINANZIARIA
        """)
    
    # Upload del file
    uploaded_file = st.file_uploader(
        "📁 Carica il file Excel",
        type=['xlsx', 'xls'],
        help="Seleziona un file Excel con i dati delle cliniche"
    )
    
    if uploaded_file is not None:
        try:
            # Leggi il file Excel
            with st.spinner("📖 Lettura del file Excel..."):
                df = pd.read_excel(uploaded_file)
            
            # Mostra informazioni sul file
            st.success(f"✅ File caricato con successo! ({len(df)} righe, {len(df.columns)} colonne)")
            
            # Validazione
            is_valid, message = validate_clinica_file(df)
            
            if not is_valid:
                st.error(f"❌ {message}")
                st.stop()
            
            # Mostra anteprima del file
            with st.expander("👀 Anteprima dati (prime 5 righe)"):
                st.dataframe(df.head())
            
            # Analisi delle cliniche
            st.markdown("### 🏥 Analisi Cliniche")
            
            unique_clinics = df["CLINICA"].unique()
            clinic_counts = df["CLINICA"].value_counts()
            
            # Statistiche principali
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Totale Cliniche", len(unique_clinics))
            
            with col2:
                st.metric("Totale Operazioni", len(df))
            
            with col3:
                avg_operations = len(df) / len(unique_clinics)
                st.metric("Media Operazioni/Clinica", f"{avg_operations:.1f}")
            
            # Tabella dettagliata delle cliniche
            st.markdown("#### 📊 Dettaglio per Clinica")
            
            clinic_summary = pd.DataFrame({
                'Clinica': clinic_counts.index,
                'Numero Operazioni': clinic_counts.values,
                'Percentuale': (clinic_counts.values / len(df) * 100).round(1)
            }).reset_index(drop=True)
            
            st.dataframe(clinic_summary, use_container_width=True)
            
            # Grafico a barre
            st.markdown("#### 📈 Grafico Operazioni per Clinica")
            st.bar_chart(clinic_counts)
            
            # Pulsante per elaborare
            st.markdown("---")
            
            if st.button("🔄 Dividi File per Cliniche", type="primary", use_container_width=True):
                
                with st.spinner("⚙️ Elaborazione in corso..."):
                    # Dividi i dati
                    cliniche_data = split_excel_by_clinica(df)
                    
                    # Crea il file ZIP
                    zip_buffer = create_zip_file(cliniche_data)
                
                st.success("✅ Elaborazione completata!")
                
                # Mostra risultati
                st.markdown("### 📁 File Creati")
                
                result_data = []
                for clinica_name, df_clinica in cliniche_data.items():
                    result_data.append({
                        'File': f"{clinica_name}.xlsx",
                        'Righe': len(df_clinica),
                        'Clinica': clinica_name
                    })
                
                result_df = pd.DataFrame(result_data)
                st.dataframe(result_df, use_container_width=True)
                
                # Pulsante download
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                zip_filename = f"cliniche_divise_{timestamp}.zip"
                
                st.download_button(
                    label="📥 Scarica tutti i file (ZIP)",
                    data=zip_buffer,
                    file_name=zip_filename,
                    mime="application/zip",
                    use_container_width=True
                )
                
                st.markdown("---")
                st.info(f"📋 Creati {len(cliniche_data)} file Excel, uno per ogni clinica")
        
        except Exception as e:
            st.error(f"❌ Errore durante l'elaborazione: {str(e)}")
            st.markdown("**Possibili cause:**")
            st.markdown("- File Excel corrotto o formato non supportato")
            st.markdown("- Colonne mancanti o nomi diversi")
            st.markdown("- Problemi di encoding del file")
    
    else:
        # Messaggio quando non c'è file caricato
        st.info("👆 Carica un file Excel per iniziare")
        
        # Esempio di formato file atteso
        with st.expander("💡 Formato file atteso"):
            st.markdown("""
            Il file Excel deve contenere almeno queste colonne:
            
            | CLINICA | AM | DM | TIPOLOGIA | TIPOLOGIA INCASSO | DATA OPERAZIONE | ... |
            |---------|----|----|-----------|-------------------|-----------------|-----|
            | Clinica A | ... | ... | ... | ... | 2024-01-15 | ... |
            | Clinica B | ... | ... | ... | ... | 2024-01-16 | ... |
            | Clinica A | ... | ... | ... | ... | 2024-01-17 | ... |
            """)

if __name__ == "__main__":
    main()


# ============================================================
# Riconciliazione Incassi AON
# ============================================================

st.write("---")
st.title("Riconciliazione Incassi AON")
st.info(
    "Carica il file delle fatture emesse verso AON e il file degli incassi ricevuti "
    "per identificare automaticamente a quali fatture corrispondono i pagamenti. "
    "La corrispondenza viene cercata prima tramite numero TT (ticket CRM), "
    "poi tramite numero fattura estratto dalla descrizione del pagamento."
)

col_aon_f, col_aon_i = st.columns(2)

with col_aon_f:
    st.subheader("📄 File Fatture AON")
    st.caption("File con le fatture emesse (es. 'Aprile 2026.xlsx')")
    fatture_aon_file = st.file_uploader(
        "Carica file fatture", type=["xls", "xlsx"], key="fatture_aon_uploader"
    )

with col_aon_i:
    st.subheader("💰 File Incassi AON")
    st.caption("File con gli incassi ricevuti (es. 'Incassi ricevuti Aprile 2026.xlsx')")
    incassi_aon_file = st.file_uploader(
        "Carica file incassi", type=["xls", "xlsx"], key="incassi_aon_uploader"
    )

if fatture_aon_file and incassi_aon_file:
    df_fatture_aon = pd.read_excel(fatture_aon_file, dtype=str)
    df_incassi_aon = pd.read_excel(incassi_aon_file, dtype=str)

    col_req_fatture = ['Master Ticket CRM', 'Num Fattura', 'PAZIENTE', 'Importo Liquidato']
    mancanti_fatture = [c for c in col_req_fatture if c not in df_fatture_aon.columns]
    mancanti_incassi = [] if 'N. assegno' in df_incassi_aon.columns else ["'N. assegno'"]

    if mancanti_fatture:
        st.error(f"Colonne mancanti nel file fatture: {mancanti_fatture}")
    elif mancanti_incassi:
        st.error("Colonna 'N. assegno' non trovata nel file incassi.")
    else:
        with st.spinner("Elaborazione riconciliazione in corso..."):
            df_riconciliato = riconcilia_incassi_aon(df_fatture_aon, df_incassi_aon)

        n_tt = (df_riconciliato['Metodo Match'] == 'TT').sum()
        n_fatt = (df_riconciliato['Metodo Match'] == 'Num Fattura').sum()
        n_none = (df_riconciliato['Metodo Match'] == 'Non trovato').sum()

        st.write("### Risultati Riconciliazione")
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.metric("Riconciliati via TT", n_tt)
        with col_m2:
            st.metric("Riconciliati via N. Fattura", n_fatt)
        with col_m3:
            st.metric("Non riconciliati", n_none)

        preview_cols = ['Importo controvalore', 'N. assegno', 'Rif. Estratto',
                        'Assistito', 'Fatture Riconciliate', 'Metodo Match']
        if 'Società' in df_riconciliato.columns:
            preview_cols = ['Società'] + preview_cols
        available_preview = [c for c in preview_cols if c in df_riconciliato.columns]

        st.write("### Anteprima (prime 20 righe)")
        st.dataframe(df_riconciliato[available_preview].head(20))

        if n_none > 0:
            with st.expander(f"Mostra {n_none} pagamenti non riconciliati"):
                mask_none = df_riconciliato['Metodo Match'] == 'Non trovato'
                st.dataframe(df_riconciliato[mask_none][available_preview])

        output_aon = BytesIO()
        with pd.ExcelWriter(output_aon, engine='openpyxl') as writer:
            df_riconciliato.to_excel(writer, index=False, sheet_name='Incassi_Riconciliati')
            df_non_ric = df_riconciliato[df_riconciliato['Metodo Match'] == 'Non trovato']
            if not df_non_ric.empty:
                df_non_ric[available_preview].to_excel(
                    writer, index=False, sheet_name='Non_Riconciliati'
                )
        output_aon.seek(0)

        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        st.download_button(
            label="📥 Scarica file riconciliato (Excel)",
            data=output_aon,
            file_name=f"Incassi_AON_Riconciliati_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_riconciliazione_aon"
        )




