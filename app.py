import streamlit as st
import pandas as pd
import io
from datetime import datetime
import xlsxwriter

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(
    page_title="Kartell - Elaboratore CSV",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- FUNZIONE DI AUTENTICAZIONE ---
def check_password():
    """Restituisce `True` se l'utente ha inserito la password corretta."""

    def password_entered():
        """Controlla se la password inserita è corretta."""
        if st.session_state["password"] == "Kartell2024": # Password predefinita, modificabile
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Rumuovi la password dallo stato per sicurezza
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # Mostra l'interfaccia di login
        st.markdown("""
            <style>
            .login-container {
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                height: 70vh;
                background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            }
            .login-card {
                background: white;
                padding: 3rem;
                border-radius: 20px;
                box-shadow: 0 15px 35px rgba(0,0,0,0.1);
                width: 400px;
                text-align: center;
            }
            .stTextInput>div>div>input {
                border-radius: 10px;
                padding: 10px;
                border: 1px solid #dee2e6;
            }
            </style>
        """, unsafe_allow_html=True)

        cols = st.columns([1, 2, 1])
        with cols[1]:
            st.markdown('<div class="login-container">', unsafe_allow_html=True)
            st.markdown("""
                <div class="login-card">
                    <h1 style='color: #007bff; margin-bottom: 20px;'>Benvenuto</h1>
                    <p style='color: #6c757d; margin-bottom: 30px;'>Inserisci la password per accedere al portale Kartell</p>
                </div>
            """, unsafe_allow_html=True)
            
            st.text_input("Password", type="password", on_change=password_entered, key="password")
            
            if "password_correct" in st.session_state and not st.session_state["password_correct"]:
                st.error("😕 Password errata. Riprova.")
            
            st.markdown('</div>', unsafe_allow_html=True)
        return False
    elif not st.session_state["password_correct"]:
        # Se l'utente ha inserito una psw errata precedentemente
        st.text_input("Password", type="password", on_change=password_entered, key="password")
        st.error("😕 Password errata. Riprova.")
        return False
    else:
        # Password corretta
        return True

if not check_password():
    st.stop()

# --- CUSTOM CSS PER DESIGN PREMIUM ---
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        width: 100%;
        border-radius: 10px;
        height: 3em;
        background-color: #007bff;
        color: white;
        font-weight: bold;
        border: none;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #0056b3;
        color: white;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    .upload-box {
        border: 2px dashed #007bff;
        border-radius: 15px;
        padding: 2.5rem;
        background-color: white;
        text-align: center;
        margin-bottom: 25px;
    }
    .header-style {
        background: linear-gradient(135deg, #007bff 0%, #0056b3 100%);
        padding: 40px;
        border-radius: 15px;
        color: white;
        text-align: center;
        margin-bottom: 30px;
        box-shadow: 0 10px 20px rgba(0,0,0,0.1);
    }
    .footer {
        text-align: center;
        color: #6c757d;
        margin-top: 50px;
        padding: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- LOGICA DI ELABORAZIONE ---

def get_metadata(df):
    """
    Estrae i metadati per il naming e il raggruppamento.
    """
    metadata = {
        "week": "W?",
        "magazzino": "Sconosciuto",
        "is_storno": False,
        "group_id": "ND"
    }
    
    if df.empty:
        return metadata

    # 1. Settimana da Data_Fattura (indice 7 o nome colonna)
    col_data = "Data_Fattura" if "Data_Fattura" in df.columns else df.columns[7] if len(df.columns) > 7 else None
    if col_data:
        try:
            # Convertiamo in datetime se non lo è già
            first_date = pd.to_datetime(df[col_data].iloc[0])
            if not pd.isna(first_date):
                metadata["week"] = f"W{first_date.isocalendar()[1]}"
        except:
            pass

    # 2. Magazzino (Fk_Magazzino)
    # Cerchiamo la colonna Fk_Magazzino
    col_mag = next((c for c in df.columns if "Magazzino" in c), None)
    if col_mag:
        mag_val = str(df[col_mag].iloc[0]).strip()
        mag_map = {
            "KARTELL_NUOVO": "Kerry",
            "KARTELL_FORNITORE_KART00": "00",
            "KARTELL_FORNITORE_KARTUS": "US",
            "KARTELL_FORNITORE_KARTAE": "Dubai",
            "KARTELL_FORNITORE_KSPPAR": "Ricambi"
        }
        metadata["magazzino"] = mag_map.get(mag_val, mag_val)
        metadata["group_id"] = mag_val

    # 3. Causale (Fk_Causale_Contabile) per N-c
    col_cau = next((c for c in df.columns if "Causale" in c), None)
    if col_cau:
        cau_val = str(df[col_cau].iloc[0]).upper().strip()
        if cau_val in ["STORNOCORRISPETTIVO", "STORNOFATTURACLIENTE"]:
            metadata["is_storno"] = True
            
    return metadata

def transform_data(uploaded_file):
    try:
        df = pd.read_csv(
            uploaded_file, 
            sep=';', 
            dtype={0: str},
            encoding='latin1'
        )
        # Pulizia nomi colonne
        df.columns = [c.strip() for c in df.columns]
        
        # 1. GESTIONE DATE (YYYYMMDD)
        date_indices = [7, 28] 
        cols_to_fix = [df.columns[i] for i in date_indices if i < len(df.columns)]
        for c in df.columns:
            if "DATA" in c.upper() and c not in cols_to_fix:
                cols_to_fix.append(c)

        for col in cols_to_fix:
            df[col] = pd.to_datetime(df[col].astype(str).str.replace('.0', '', regex=False), format='%Y%m%d', errors='coerce')
        
        # 2. GESTIONE COLONNE NUMERICHE
        num_cols = [
            "Tasso_di_Cambio", "Quantita", "Totale_Merce_In_Valuta_ii", 
            "Sconto_Fattura_Valuta", "Totale_Merce_In_Valuta_ie", 
            "Sconto_Prodotti_Valuta", "Totale_Sconti_Valuta", 
            "Totale_Merce_EUR_ii", "Totale_Merce_EUR_ie", 
            "Sconto_Fattura_EUR", "Sconto_Prodotti_EUR", 
            "Totale_Sconti_EUR", "Totale_Sconto_Pct", 
            "Totale_Merce_Netto_Sconti_In_Valuta", "Totale_Merce_Netto_Sconti_EUR",
            "Fk_Ordine_Cliente", "Fk_Dettaglio_Ordine_Fornitore", "Ambito_Nazionalita"
        ]
        
        for col in num_cols:
            if col in df.columns:
                # Sostituiamo la virgola con il punto e convertiamo in numerico
                df[col] = df[col].astype(str).str.replace(',', '.', regex=False).str.strip()
                df[col] = pd.to_numeric(df[col], errors='coerce')
            
        return df
    except Exception as e:
        st.error(f"Errore nella lettura del file {uploaded_file.name}: {e}")
        return None

def generate_excel(processed_data):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Formati
    header_format = workbook.add_format({'bold': True, 'bg_color': '#007bff', 'font_color': 'white', 'border': 1, 'align': 'center'})
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'align': 'center'})
    text_format = workbook.add_format({'num_format': '@', 'align': 'left'})
    curr_format = workbook.add_format({'num_format': '#,##0.00 €', 'align': 'right'})
    summary_header = workbook.add_format({'bold': True, 'bg_color': '#FFEB3B', 'border': 1, 'align': 'center'})
    
    for item in processed_data:
        df = item["df"]
        sheet_name = item["original_name"][:31].replace('.csv', '')
        worksheet = workbook.add_worksheet(sheet_name)
        
        # Scrittura dati CSV
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
        for row_num, row_data in enumerate(df.values):
            for col_num, cell_value in enumerate(row_data):
                fmt = None
                if col_num == 0: fmt = text_format
                elif isinstance(cell_value, pd.Timestamp): fmt = date_format
                
                if pd.isna(cell_value):
                    worksheet.write(row_num + 1, col_num, "", fmt)
                else:
                    worksheet.write(row_num + 1, col_num, cell_value, fmt)

        # --- AGGIUNTA TABELLA RIEPILOGO (Macro Logic) ---
        def index_to_letter(idx):
            letter = ""
            while idx >= 0:
                letter = chr(65 + (idx % 26)) + letter
                idx = (idx // 26) - 1
            return letter

        def get_col_letter(col_name):
            try:
                idx = list(df.columns).index(col_name)
                return index_to_letter(idx)
            except: return None

        col_eur = get_col_letter("Totale_Merce_Netto_Sconti_EUR") or "Z"
        col_valuta = get_col_letter("Totale_Merce_Netto_Sconti_In_Valuta") or "Y"
        col_amb = get_col_letter("Ambito_Nazionalita") or "AO"
        col_naz = get_col_letter("Nazione") or "I"
        
        # Percentuale Fee basata sul magazzino
        mag = item["metadata"]["magazzino"]
        fee_pcts = {
            "Kerry": 0.22,
            "00": 0.185,
            "Dubai": 0.15,
            "US": 0.16,
            "Ricambi": 0.16
        }
        fee_val = fee_pcts.get(mag, 0.22)
        
        last_col_idx = len(df.columns)
        summary_start_col = last_col_idx + 1
        
        # Definizione intestazioni
        headers = [
            "Totale fatturato in Euro", 
            "Ns Fee in Euro", 
            "Ns Fee vendite Extra-UE in Euro",
            "Totale fatturato in Valuta",
            "Ns Fee in Valuta"
        ]
        
        for i, h in enumerate(headers):
            worksheet.write(0, summary_start_col + i, h, summary_header)

        # Lettere colonne riepilogo
        c_tot_eur = index_to_letter(summary_start_col)
        c_tot_val = index_to_letter(summary_start_col + 3)
        
        # 1. Totale Fatturato in Euro
        worksheet.write_formula(1, summary_start_col, f"=SUM({col_eur}:{col_eur})", curr_format)
        
        # 2. Ns Fee in Euro (basato su magazzino)
        worksheet.write_formula(1, summary_start_col + 1, f"={c_tot_eur}2*{fee_val}", curr_format)
        
        # 3. Ns Fee Extra-UE (Solo se NON US o Dubai)
        if mag not in ["US", "Dubai"]:
            formula_extra = (
                f"=SUMIF({col_amb}:{col_amb}, 2, {col_eur}:{col_eur})*0.025 "
                f"- SUMIFS({col_eur}:{col_eur}, {col_amb}:{col_amb}, 2, {col_naz}:{col_naz}, \"MC\")*0.025 "
                f"- SUMIFS({col_eur}:{col_eur}, {col_amb}:{col_amb}, 2, {col_naz}:{col_naz}, \"FR\")*0.025"
            )
            worksheet.write_formula(1, summary_start_col + 2, formula_extra, curr_format)
        else:
            worksheet.write(1, summary_start_col + 2, "N/A", text_format)

        # 4. Totale Fatturato in Valuta (Solo se NON Kerry, 00 o Ricambi)
        if mag not in ["Kerry", "00", "Ricambi"]:
            worksheet.write_formula(1, summary_start_col + 3, f"=SUM({col_valuta}:{col_valuta})", curr_format)
            # 5. Ns Fee in Valuta
            worksheet.write_formula(1, summary_start_col + 4, f"={c_tot_val}2*{fee_val}", curr_format)
        else:
            worksheet.write(1, summary_start_col + 3, "N/A", text_format)
            worksheet.write(1, summary_start_col + 4, "N/A", text_format)

        # Autofit
        for i, col in enumerate(df.columns):
            worksheet.set_column(i, i, 15)
        worksheet.set_column(summary_start_col, summary_start_col + 4, 30)

    workbook.close()
    return output.getvalue()

# --- INTERFACCIA UTENTE ---

st.markdown('<div class="header-style"><h1>📊 Elaboratore Fatturato Kartell</h1><p>Analisi automatica Gruppi e Naming Dinamico</p></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Carica un file CSV per l'elaborazione settimanale", 
    type=["csv"], 
    accept_multiple_files=False
)

if uploaded_file:
    df = transform_data(uploaded_file)
    if df is not None:
        meta = get_metadata(df)
        
        st.write(f"### 📄 Analisi File: {uploaded_file.name}")
        c1, c2, c3 = st.columns(3)
        c1.metric("Settimana", meta["week"])
        c2.metric("Magazzino", meta["magazzino"])
        c3.metric("Tipo", "Storno (N-c)" if meta["is_storno"] else "Standard")
        
        with st.expander("Anteprima dati caricati"):
            st.dataframe(df.head(10))

        st.markdown("---")
        
        # Determiniamo il nome del file finale
        storno_suffix = " N-c" if meta["is_storno"] else ""
        final_filename = f"Conteggi vendite {meta['week']} {meta['magazzino']}{storno_suffix}.xlsx"
        
        st.info(f"📁 Il file verrà scaricato come: **{final_filename}**")
        
        if st.button("🚀 GENERA E SCARICA EXCEL"):
            with st.spinner("Generazione in corso..."):
                # Prepariamo i dati (singolo file in lista per compatibilità con funzione)
                gen_data = [{"df": df, "original_name": uploaded_file.name, "metadata": meta}]
                excel_binary = generate_excel(gen_data)
                
                st.success("✅ Elaborazione completata!")
                st.download_button(
                    label="📥 Clicca qui per scaricare",
                    data=excel_binary,
                    file_name=final_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

else:
    st.markdown("""
        <div class="upload-box">
            <h3>Logica di Naming Automatica:</h3>
            <p>Il file finale seguirà lo schema:</p>
            <code>Conteggi vendite [Settimana] [Magazzino] [N-c]</code>
            <ul style="text-align: left; margin-top: 10px;">
                <li><b>W:</b> Ricavato dalla colonna Data_Fattura</li>
                <li><b>Magazzino:</b> Tradotto da KARTELL_ (Kerry, 00, US, Dubai, Ricambi)</li>
                <li><b>Storno (N-c):</b> Aggiunto automaticamente per Causali di storno</li>
            </ul>
        </div>
    """, unsafe_allow_html=True)

st.markdown('<div class="footer">Replica logica business Kartell per fatturato settimanale</div>', unsafe_allow_html=True)

