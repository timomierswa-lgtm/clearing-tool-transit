import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import locale

# Versuche deutsches Locale für den Monatsnamen zu setzen (falls vom Server unterstützt)
try:
    locale.setlocale(locale.LC_TIME, 'de_DE.UTF-8')
except:
    pass

st.set_page_config(page_title="Bestandsumlagerung Splitter", layout="wide")
st.title("📦 Bestandsumlagerung - Datei Splitter")
st.write("Lade hier das Wörterbuch, die Ursprungsdatei und die Vorlage hoch. Das Tool splittet die Dateien pro Source BP, befüllt die Daten intelligent und erstellt ein ZIP-Archiv für SAP.")

# 1. Datei-Uploads
col1, col2, col3 = st.columns(3)
with col1:
    dict_file = st.file_uploader("1. Wörterbuch (.xlsx)", type=["xlsx"])
with col2:
    source_file = st.file_uploader("2. Ursprungsdatei (.xlsx)", type=["xlsx"])
with col3:
    template_file = st.file_uploader("3. Vorlage (.xlsx)", type=["xlsx"])

if dict_file and source_file and template_file:
    try:
        # 2. Dateien einlesen
        df_dict = pd.read_excel(dict_file)
        df_source = pd.read_excel(source_file)
        df_template_empty = pd.read_excel(template_file)
        
        # Spaltennamen zur Sicherheit von Leerzeichen befreien
        df_dict.columns = df_dict.columns.str.strip()
        df_source.columns = df_source.columns.str.strip()
        df_template_empty.columns = df_template_empty.columns.str.strip()

        st.success("Alle Dateien erfolgreich eingelesen!")
        
        if st.button("Dateien generieren & Splitten"):
            
            # --- DATEN VORBEREITUNG ---
            
            # BP Code sicher als String ohne ".0" formatieren
            if 'TransitSourceDocumentBPCode' in df_source.columns:
                df_source['TransitSourceDocumentBPCode'] = df_source['TransitSourceDocumentBPCode'].fillna(0).astype(int).astype(str)
                # Entferne '0' BPs (falls leere Zeilen vorhanden waren)
                df_source = df_source[df_source['TransitSourceDocumentBPCode'] != '0']
                
            if 'TransitSourceDocumentBPCode' in df_dict.columns:
                df_dict['TransitSourceDocumentBPCode'] = df_dict['TransitSourceDocumentBPCode'].fillna(0).astype(int).astype(str)
            
            # Dictionary und Source mergen 
            merged_df = pd.merge(
                df_source, 
                df_dict, 
                left_on=['TransitSourceDocumentBPCode', 'TransitSourceDocumentType'], 
                right_on=['TransitSourceDocumentBPCode', 'TransitSourceDocumentType'], 
                how='left'
            )
            
            # Fehlende Mappings prüfen
            fehlende_mappings = merged_df['ToBin'].isna().sum()
            if fehlende_mappings > 0:
                st.warning(f"Achtung: Für {fehlende_mappings} Zeilen konnte kein ToBin im Wörterbuch gefunden werden.")

            # Datum und Monat vorbereiten
            heute_str = datetime.now().strftime('%Y-%m-%d')
            # Fallback falls deutsches Locale auf dem Server nicht klappt
            monate = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"]
            aktueller_monat = monate[datetime.now().month - 1]
            doc_comment_text = f"Bin Clearing {aktueller_monat}"

            # --- SPLITTEN & ZIP ERSTELLEN ---
            zip_buffer = io.BytesIO()
            bps = merged_df['TransitSourceDocumentBPCode'].dropna().unique()
            
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for bp in bps:
                    # Daten für diesen spezifischen BP filtern
                    bp_data = merged_df[merged_df['TransitSourceDocumentBPCode'] == bp].copy()
                    
                    # Leeren DataFrame im Format der Vorlage erstellen
                    df_out = pd.DataFrame(columns=df_template_empty.columns)
                    
                    # Werte mappen (mit .get(), um Fehler zu vermeiden falls eine Spalte doch fehlt)
                    df_out['BP'] = bp
                    df_out['DocDate'] = heute_str
                    df_out['TaxDate'] = heute_str
                    
                    # From Bin und From Wh
                    from_bins = bp_data.get('BinCode', '')
                    df_out['From Bin'] = from_bins
                    df_out['From Wh'] = from_bins.astype(str).str.split('-').str[0]
                    
                    # To Bin und To Wh
                    to_bins = bp_data.get('ToBin', '')
                    df_out['To Bin'] = to_bins
                    df_out['To Wh'] = to_bins.astype(str).apply(lambda x: x.split('-')[0] if pd.notna(x) else '')
                    
                    # Restliche Felder
                    df_out['EAN'] = bp_data.get('EAN', '')
                    df_out['Qty'] = bp_data.get('WHSinStock', '')
                    df_out['Unit Price'] = bp_data.get('MovAVGpriceEUR', '')
                    df_out['Doc Comments'] = doc_comment_text
                    
                    # Excel-Datei im Arbeitsspeicher erstellen
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        df_out.to_excel(writer, index=False, sheet_name='Clearing')
                    
                    # Datei zum ZIP-Archiv hinzufügen
                    file_name = f"Clearing_Umlagerung_{bp}.xlsx"
                    zf.writestr(file_name, excel_buffer.getvalue())
            
            # --- DOWNLOAD BEREITSTELLEN ---
            st.success(f"Erfolgreich {len(bps)} Datei(en) generiert!")
            st.download_button(
                label="📥 ZIP-Archiv mit allen Dateien herunterladen",
                data=zip_buffer.getvalue(),
                file_name=f"SAP_Clearing_Dateien_{heute_str}.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Es gab einen Fehler beim Verarbeiten: {e}")
        st.info("Bitte überprüfe, ob die Spaltennamen in den hochgeladenen Dateien korrekt sind.")