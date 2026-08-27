import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import os

st.set_page_config(page_title="Bestandsumlagerung Splitter", layout="wide")
st.title("📦 Bestandsumlagerung - Datei Splitter")
st.write("Lade hier deine Ursprungsdatei hoch. Das Tool splittet die Dateien pro Source BP und erstellt ein ZIP-Archiv für SAP.")

# Dateipfade im GitHub-Repository
DICT_PATH = "Wörterbuch.xlsx"
TEMPLATE_PATH = "CLEARING DATEI MAYEL VORLAGE.xlsx"

# Prüfen, ob die Hintergrunddateien in GitHub existieren
if not os.path.exists(DICT_PATH) or not os.path.exists(TEMPLATE_PATH):
    st.error(f"Fehler: Bitte lade die Dateien '{DICT_PATH}' und '{TEMPLATE_PATH}' in dein GitHub-Repository hoch!")
else:
    # 1. Wörterbuch im Hintergrund laden und zur Kontrolle anzeigen
    df_dict = pd.read_excel(DICT_PATH)
    df_dict.columns = df_dict.columns.str.strip()
    
    with st.expander("📖 Aktuelles Wörterbuch anzeigen", expanded=False):
        st.dataframe(df_dict, use_container_width=True)
        st.info("Tipp: Um dieses Wörterbuch anzupassen, lade einfach eine aktualisierte 'Wörterbuch.xlsx' in dein GitHub-Repository hoch.")

    # 2. Upload-Feld NUR noch für die Ursprungsdatei
    st.markdown("### Schritt 1: Datei hochladen")
    source_file = st.file_uploader("Ursprungsdatei (.xlsx)", type=["xlsx"])

    if source_file:
        try:
            # 3. Dateien einlesen
            df_source = pd.read_excel(source_file)
            df_template_empty = pd.read_excel(TEMPLATE_PATH)
            
            df_source.columns = df_source.columns.str.strip()
            df_template_empty.columns = df_template_empty.columns.str.strip()

            st.success("Ursprungsdatei erfolgreich eingelesen!")
            
            st.markdown("### Schritt 2: Verarbeiten")
            if st.button("Dateien generieren & Splitten"):
                
                # --- DATEN VORBEREITUNG ---
                if 'TransitSourceDocumentBPCode' in df_source.columns:
                    df_source['TransitSourceDocumentBPCode'] = df_source['TransitSourceDocumentBPCode'].fillna(0).astype(int).astype(str)
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
                monate = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"]
                aktueller_monat = monate[datetime.now().month - 1]
                doc_comment_text = f"Bin Clearing {aktueller_monat}"

                # --- SPLITTEN & ZIP ERSTELLEN ---
                zip_buffer = io.BytesIO()
                bps = merged_df['TransitSourceDocumentBPCode'].dropna().unique()
                
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for bp in bps:
                        bp_data = merged_df[merged_df['TransitSourceDocumentBPCode'] == bp].copy()
                        df_out = pd.DataFrame(columns=df_template_empty.columns)
                        
                        df_out['BP'] = bp
                        df_out['DocDate'] = heute_str
                        df_out['TaxDate'] = heute_str
                        
                        from_bins = bp_data.get('BinCode', '')
                        df_out['From Bin'] = from_bins
                        df_out['From Wh'] = from_bins.astype(str).str.split('-').str[0]
                        
                        to_bins = bp_data.get('ToBin', '')
                        df_out['To Bin'] = to_bins
                        df_out['To Wh'] = to_bins.astype(str).apply(lambda x: x.split('-')[0] if pd.notna(x) else '')
                        
                        df_out['EAN'] = bp_data.get('EAN', '')
                        df_out['Qty'] = bp_data.get('WHSinStock', '')
                        df_out['Unit Price'] = bp_data.get('MovAVGpriceEUR', '')
                        df_out['Doc Comments'] = doc_comment_text
                        
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            df_out.to_excel(writer, index=False, sheet_name='Clearing')
                        
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
            st.info("Bitte überprüfe, ob die Spaltennamen in der hochgeladenen Datei korrekt sind.")
