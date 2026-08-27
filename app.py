import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import os

st.set_page_config(page_title="Bestandsumlagerung Splitter", layout="wide")
st.title("📦 Bestandsumlagerung - Datei Splitter")
st.write("Lade hier deine Ursprungsdatei hoch. Passe bei Bedarf das Wörterbuch an.")

# Dateipfade im GitHub-Repository
DICT_PATH = "Wörterbuch.xlsx"
TEMPLATE_PATH = "CLEARING DATEI MAYEL VORLAGE.xlsx"

if not os.path.exists(DICT_PATH) or not os.path.exists(TEMPLATE_PATH):
    st.error(f"Fehler: Bitte lade die Dateien '{DICT_PATH}' und '{TEMPLATE_PATH}' in dein GitHub-Repository hoch!")
else:
    # --- 1. WÖRTERBUCH INTERAKTIV ---
    st.markdown("### 📖 Wörterbuch anpassen")
    st.write("Klicke in die Tabelle, um Werte zu ändern. Um neue BPs hinzuzufügen, scrolle nach unten und klicke in die leere Zeile.")
    
    # Wörterbuch laden
    df_dict_original = pd.read_excel(DICT_PATH)
    df_dict_original.columns = df_dict_original.columns.str.strip()
    
    # Interaktiver Editor (erlaubt das Hinzufügen und Löschen von Zeilen)
    edited_dict = st.data_editor(df_dict_original, num_rows="dynamic", use_container_width=True)

    # --- 2. URSPRUNGSDATEI UPLOAD ---
    st.markdown("### 📤 Schritt 1: Datei hochladen")
    source_file = st.file_uploader("Ursprungsdatei (.xlsx) hier ablegen", type=["xlsx"])

    if source_file:
        try:
            # 3. Dateien einlesen (EAN zwingend als Text/String einlesen!)
            df_source = pd.read_excel(source_file, dtype={'EAN': str})
            df_template_empty = pd.read_excel(TEMPLATE_PATH)
            
            df_source.columns = df_source.columns.str.strip()
            df_template_empty.columns = df_template_empty.columns.str.strip()

            # EAN bereinigen (falls Pandas '.0' am Ende anhängt)
            if 'EAN' in df_source.columns:
                df_source['EAN'] = df_source['EAN'].astype(str).str.replace(r'\.0$', '', regex=True)
                df_source['EAN'] = df_source['EAN'].replace('nan', '')

            st.success("Ursprungsdatei erfolgreich eingelesen!")
            
            st.markdown("### ⚙️ Schritt 2: Verarbeiten & Herunterladen")
            if st.button("Dateien generieren & Splitten"):
                
                # --- DATEN VORBEREITUNG ---
                if 'TransitSourceDocumentBPCode' in df_source.columns:
                    df_source['TransitSourceDocumentBPCode'] = df_source['TransitSourceDocumentBPCode'].fillna(0).astype(int).astype(str)
                    df_source = df_source[df_source['TransitSourceDocumentBPCode'] != '0']
                    
                if 'TransitSourceDocumentBPCode' in edited_dict.columns:
                    edited_dict['TransitSourceDocumentBPCode'] = edited_dict['TransitSourceDocumentBPCode'].fillna(0).astype(int).astype(str)
                
                # Dictionary und Source mergen (mit dem bearbeiteten Wörterbuch!)
                merged_df = pd.merge(
                    df_source, 
                    edited_dict, 
                    left_on=['TransitSourceDocumentBPCode', 'TransitSourceDocumentType'], 
                    right_on=['TransitSourceDocumentBPCode', 'TransitSourceDocumentType'], 
                    how='left'
                )
                
                # Datum und Monat vorbereiten (Immer der Tag, an dem die Datei erstellt wird)
                heute_str = datetime.now().strftime('%Y-%m-%d')
                monate = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"]
                aktueller_monat = monate[datetime.now().month - 1]
                doc_comment_text = f"Bin Clearing {aktueller_monat}"

                # --- SPLITTEN & ZIP ERSTELLEN ---
                zip_buffer = io.BytesIO()
                
                # In gemappte (Erfolg) und ungemappte (Fehler) aufteilen
                mapped_df = merged_df[merged_df['ToBin'].notna()]
                unmapped_df = merged_df[merged_df['ToBin'].isna()]
                
                bps = mapped_df['TransitSourceDocumentBPCode'].unique()
                
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    
                    # 1. Erfolgreich gemappte Dateien erstellen
                    for bp in bps:
                        # Index zurücksetzen, um Zuweisungsfehler zu vermeiden!
                        bp_data = mapped_df[mapped_df['TransitSourceDocumentBPCode'] == bp].copy().reset_index(drop=True)
                        anzahl_zeilen = len(bp_data)
                        
                        # Neue Tabelle mit der exakten Anzahl an Zeilen erstellen
                        df_out = pd.DataFrame(index=range(anzahl_zeilen), columns=df_template_empty.columns)
                        
                        # Feste Werte für JEDE Zeile eintragen
                        df_out['BP'] = [bp] * anzahl_zeilen
                        df_out['DocDate'] = [heute_str] * anzahl_zeilen
                        df_out['TaxDate'] = [heute_str] * anzahl_zeilen
                        df_out['Doc Comments'] = [doc_comment_text] * anzahl_zeilen
                        
                        # Dynamische Werte mappen
                        from_bins = bp_data.get('BinCode', pd.Series([''] * anzahl_zeilen))
                        df_out['From Bin'] = from_bins.values
                        df_out['From Wh'] = from_bins.astype(str).str.split('-').str[0].values
                        
                        to_bins = bp_data.get('ToBin', pd.Series([''] * anzahl_zeilen))
                        df_out['To Bin'] = to_bins.values
                        df_out['To Wh'] = to_bins.astype(str).str.split('-').str[0].values
                        
                        df_out['EAN'] = bp_data.get('EAN', pd.Series([''] * anzahl_zeilen)).values
                        df_out['Qty'] = bp_data.get('WHSinStock', pd.Series([''] * anzahl_zeilen)).values
                        df_out['Unit Price'] = bp_data.get('MovAVGpriceEUR', pd.Series([''] * anzahl_zeilen)).values
                        
                        # In Excel umwandeln
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            df_out.to_excel(writer, index=False, sheet_name='Clearing')
                        
                        file_name = f"Clearing_Umlagerung_{bp}.xlsx"
                        zf.writestr(file_name, excel_buffer.getvalue())
                    
                    # 2. Nicht zugeordnete Zeilen exportieren (falls vorhanden)
                    if not unmapped_df.empty:
                        error_buffer = io.BytesIO()
                        with pd.ExcelWriter(error_buffer, engine='openpyxl') as writer:
                            unmapped_df.to_excel(writer, index=False, sheet_name='Nicht_Zugeordnet')
                        zf.writestr("Nicht_Zugeordnet_Fehler.xlsx", error_buffer.getvalue())
                        st.warning(f"Achtung: {len(unmapped_df)} Zeilen konnten nicht zugeordnet werden. Diese findest du in der Datei 'Nicht_Zugeordnet_Fehler.xlsx' im ZIP-Archiv.")

                # --- DOWNLOAD BEREITSTELLEN ---
                st.success(f"Erfolgreich {len(bps)} fertige Datei(en) generiert!")
                st.download_button(
                    label="📥 ZIP-Archiv herunterladen",
                    data=zip_buffer.getvalue(),
                    file_name=f"SAP_Clearing_Dateien_{heute_str}.zip",
                    mime="application/zip"
                )

        except Exception as e:
            st.error(f"Es gab einen Fehler beim Verarbeiten: {e}")
            st.info("Bitte überprüfe, ob die Spaltennamen in der hochgeladenen Datei korrekt sind.")
