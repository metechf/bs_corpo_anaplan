import streamlit as st
import pandas as pd
from io import BytesIO
import re
import zipfile
import os
import tempfile
import subprocess
import sys
import unicodedata
# Import du module mapping
try:
    from mapping import AnaplanMapper
    MAPPING_AVAILABLE = True
except ImportError:
    MAPPING_AVAILABLE = False

# ---------------------------------------------------------------------------
# Global print filter to masquer les logs DEBUG
# ---------------------------------------------------------------------------
import builtins as _builtins

def _print_no_debug(*args, **kwargs):
    """Wrapper de print qui ignore toute ligne contenant le mot DEBUG."""
    if args and isinstance(args[0], str) and "DEBUG" in args[0]:
        return  # on n'affiche pas les logs de debug
    return _builtins.print(*args, **kwargs)

# Remplacer la fonction print globale par la version filtrée
print = _print_no_debug

# ---------------------------------------------------------------------------
# Classe principale pour gérer tous les modules Anaplan
# ---------------------------------------------------------------------------

class AnaplanMainApp:
    """Classe principale qui combine tous les générateurs Anaplan dans une interface unifiée."""
    
    def __init__(self):
        self.quarter_to_months = {"Q1": [1, 2, 3], "Q2": [4, 5, 6], "Q3": [7, 8, 9], "Q4": [10, 11, 12]}
        # Initialiser les variables de session
        if 'generated_files' not in st.session_state:
            st.session_state.generated_files = {}
        if 'mapped_files' not in st.session_state:
            st.session_state.mapped_files = {}
        if 'generation_completed' not in st.session_state:
            st.session_state.generation_completed = False
        if 'mapping_completed' not in st.session_state:
            st.session_state.mapping_completed = False
        if 'sanity_check_completed' not in st.session_state:
            st.session_state.sanity_check_completed = False
        if 'sanity_check_files' not in st.session_state:
            st.session_state.sanity_check_files = []
        # Dictionnaire français/anglais mois -> numéro (accents supprimés)
        self.month_name_to_num = {
            "JANVIER": 1, "JANUARY": 1,
            "FEVRIER": 2, "FÉVRIER": 2, "FEBRUARY": 2,
            "MARS": 3, "MARCH": 3,
            "AVRIL": 4, "APRIL": 4,
            "MAI": 5, "MAY": 5,
            "JUIN": 6, "JUNE": 6,
            "JUILLET": 7, "JULY": 7,
            "AOUT": 8, "AOÛT": 8, "AUGUST": 8,
            "SEPTEMBRE": 9, "SEPTEMBER": 9,
            "OCTOBRE": 10, "OCTOBER": 10,
            "NOVEMBRE": 11, "NOVEMBER": 11,
            "DECEMBRE": 12, "DÉCEMBRE": 12, "DECEMBER": 12,
        }
    
    # ---------------------------------------------------------------------------
    # Fonctions utilitaires communes
    # ---------------------------------------------------------------------------
    
    @staticmethod
    def tokenizer_contract(val):
        """Tokenise un contract pour le matching par similarité."""
        if pd.isna(val) or val == "":
            return set()
        val = str(val).lower()
        val = val.replace(",", ".")
        tokens = re.split(r"[_\s\.,']", val)
        tokens = [t for t in tokens if t]
        return set(tokens)
    
    @staticmethod
    def match_par_tokens(volume_tokens, fulfill_df_tokens, seuil=6):
        """Trouve le meilleur match Country basé sur la similarité des tokens de Contract."""
        best_match = ""
        max_communs = 0
        for _, row in fulfill_df_tokens.iterrows():
            communs = volume_tokens.intersection(row["Contract_tokens"])
            if len(communs) > max_communs and len(communs) >= seuil:
                best_match = row["Country"]
                max_communs = len(communs)
        return best_match
    
    @staticmethod
    def match_par_tokens_price(volume_tokens, fulfill_df_tokens, seuil=6):
        """Trouve le meilleur match Price basé sur la similarité des tokens de Contract."""
        best_match = ""
        max_communs = 0
        for _, row in fulfill_df_tokens.iterrows():
            communs = volume_tokens.intersection(row["Contract_tokens"])
            if len(communs) > max_communs and len(communs) >= seuil:
                best_match = row["Price"]
                max_communs = len(communs)
        return best_match
    
    @staticmethod
    def get_prix_from_contract_price(contract, contract_price_df, mois):
        """Récupère le prix FOB depuis #ContractPrice pour un contrat et un mois donnés."""
        if contract_price_df.empty or not contract:
            return ""
        
        # Chercher la ligne correspondant au contrat
        matching_rows = contract_price_df[contract_price_df.get("Contract", "") == contract]
        if matching_rows.empty:
            return ""
        
        price_row = matching_rows.iloc[0]
        
        # Mapping mois vers colonnes de prix (gère Oct/Nov/Déc mensuels avec fallback Q4)
        # Pour 7-9 on conserve les colonnes mensuelles explicites déjà utilisées
        month_to_price_cols = {
            7: ["Price[Juillet]"],
            8: ["Price[Aout]", "Price[Août]"],
            9: ["Price[Septembre]"],
            10: ["Price[Octobre]", "Price[Q4]"],
            11: ["Price[Novembre]", "Price[Q4]"],
            12: ["Price[Decembre]", "Price[Décembre]", "Price[Q4]"]
        }
        
        candidate_cols = month_to_price_cols.get(mois, [])
        for col_name in candidate_cols:
            if col_name in price_row.index:
                prix_value = price_row[col_name]
                if pd.notna(prix_value) and str(prix_value).strip() != "":
                    try:
                        return float(prix_value)
                    except:
                        return ""
        return ""
    
    @staticmethod
    def extraire_table_par_nom(feuille: pd.DataFrame, nom_table: str) -> pd.DataFrame:
        """Extrait une table dont le nom commence par #<nom_table> dans une feuille brute."""
        table_start = None
        table_end = None
        for i, val in enumerate(feuille.iloc[:, 0]):
            if isinstance(val, str) and val.strip().startswith("#"):
                nom = val.strip().lstrip("#").strip()
                if nom_table.lower() in nom.lower():
                    table_start = i + 1
                elif table_start is not None:
                    table_end = i
                    break
        if table_start is None:
            return pd.DataFrame()
        if table_end is None:
            table_end = len(feuille)
        table = feuille.iloc[table_start:table_end].copy()
        table.dropna(how="all", inplace=True)
        table.columns = table.iloc[0]
        table = table[1:]
        table.dropna(how="all", inplace=True)
        return table.reset_index(drop=True)

    @staticmethod
    def find_sheet(xls: pd.ExcelFile, *keywords):
        """Retourne le premier nom de feuille contenant toutes les expressions fournies."""
        kws = [k.lower() for k in keywords]
        for name in xls.sheet_names:
            lname = name.lower()
            if all(k in lname for k in kws):
                return name
        return None
    
    # ---------------------------------------------------------------------------
    # Fonctions TCD (Tableaux Croisés Dynamiques)
    # ---------------------------------------------------------------------------
    
    def create_tcd_synthese(self, df_plat, volume_col="VOLUME (T)", site_col="Site/Entité", 
                           qualite_col="Qualité", operation_col=None, type_transaction_col=None):
        """Crée un TCD à partir du fichier plat avec mois détaillés + trimestres agrégés (VOLUMES SEULEMENT)."""
        if df_plat.empty:
            return pd.DataFrame()
        
        try:
            # Définir les colonnes d'index selon le type de fichier
            index_cols = [site_col, qualite_col]
            if operation_col and operation_col in df_plat.columns:
                index_cols.append(operation_col)
            if type_transaction_col and type_transaction_col in df_plat.columns:
                index_cols.append(type_transaction_col)
            
            # Vérifier que les colonnes existent
            missing_cols = [col for col in index_cols + [volume_col, "Mois"] if col not in df_plat.columns]
            if missing_cols:
                print(f"⚠️ Colonnes manquantes pour TCD Synthèse : {missing_cols}")
                return pd.DataFrame()
            
            # Créer le TCD avec les mois
            tcd_data = []
            
            # Grouper par les colonnes d'index
            grouped = df_plat.groupby(index_cols + ["Mois"])[volume_col].sum().reset_index()
            
            # Pivoter pour avoir les mois en colonnes
            pivot_mois = grouped.pivot_table(
                index=index_cols,
                columns="Mois", 
                values=volume_col,
                fill_value=0,
                aggfunc='sum'
            ).reset_index()
            
            # S'assurer que TOUS les mois (1-12) sont présents
            for i in range(1, 13):
                col_name = f"Volume Mois {i}"
                if i in pivot_mois.columns:
                    # Renommer la colonne existante
                    pivot_mois = pivot_mois.rename(columns={i: col_name})
                else:
                    # Ajouter la colonne manquante avec des 0
                    pivot_mois[col_name] = 0
            
            # Réorganiser les colonnes dans l'ordre : index + mois 1-12 + trimestres
            ordered_cols = index_cols.copy()
            for i in range(1, 13):
                ordered_cols.append(f"Volume Mois {i}")
            
            # Ajouter les trimestres
            for q, mois_list in self.quarter_to_months.items():
                vol_cols = [f"Volume Mois {m}" for m in mois_list]
                pivot_mois[f"Volume {q}"] = pivot_mois[vol_cols].sum(axis=1)
                ordered_cols.append(f"Volume {q}")
            
            # Réorganiser les colonnes selon l'ordre voulu
            pivot_mois = pivot_mois[ordered_cols]
            
            return pivot_mois
            
        except Exception as e:
            print(f"❌ Erreur TCD Synthèse : {e}")
            return pd.DataFrame()
    
    def create_tcd_sources(self, sources_data_list, source_names=None):
        """Copie les tables sources avec leur structure ORIGINALE (#TableName + données)."""
        if not sources_data_list:
            return pd.DataFrame()
        
        try:
            combined_data = []
            max_cols = 0
            
            # Déterminer le nombre maximum de colonnes pour aligner toutes les tables
            for df_source in sources_data_list:
                if not df_source.empty:
                    max_cols = max(max_cols, len(df_source.columns))
            
            # Si aucune donnée, retourner DataFrame vide
            if max_cols == 0:
                return pd.DataFrame()
            
            # Créer les noms de colonnes génériques pour couvrir le maximum
            generic_columns = [f"Col_{i}" for i in range(max_cols)]
            
            for i, df_source in enumerate(sources_data_list):
                if df_source.empty:
                    continue
                    
                # Nom de la table (avec # comme dans les fichiers originaux)
                source_name = source_names[i] if source_names and i < len(source_names) else f"Source_{i+1}"
                
                # 1. Ligne avec #NomTable (comme dans les fichiers inputs)
                table_header_data = [f"#{source_name}"] + [""] * (max_cols - 1)
                table_header = pd.DataFrame([table_header_data], columns=generic_columns)
                combined_data.append(table_header)
                
                # 2. Ligne d'en-têtes des colonnes de la table
                headers_data = list(df_source.columns) + [""] * (max_cols - len(df_source.columns))
                headers_row = pd.DataFrame([headers_data], columns=generic_columns)
                combined_data.append(headers_row)
                
                # 3. Données de la table (étendre pour correspondre au max_cols)
                df_extended = df_source.copy()
                for j in range(len(df_source.columns), max_cols):
                    df_extended[f"Col_{j}"] = ""
                
                # Renommer les colonnes pour correspondre au schéma générique
                df_extended.columns = generic_columns
                combined_data.append(df_extended)
                
                # 4. Ligne vide entre les tables
                empty_row_data = [""] * max_cols
                empty_row = pd.DataFrame([empty_row_data], columns=generic_columns)
                combined_data.append(empty_row)
            
            if combined_data:
                result = pd.concat(combined_data, ignore_index=True)
                # Laisser les colonnes sans noms pour avoir une première ligne vide dans Excel
                result.columns = ["" for _ in range(len(result.columns))]
                return result
            
            return pd.DataFrame()
            
        except Exception as e:
            print(f"❌ Erreur TCD Sources : {e}")
            return pd.DataFrame()
    
    @staticmethod
    def is_valid_consumption_type(column_name: str):
        """Détermine si une colonne représente un vrai type de consommation spécifique et non un processus chimique."""
        column_lower = column_name.lower().strip()
        
        # Mots-clés de PROCESSUS CHIMIQUES à EXCLURE
        process_keywords = {
            'concentration', 'fusion', 'phosphorique', 'sulfurique', 
            'granulation', 'banalization', 'clarification', 'purification',
            'treatment', 'process', 'production', 'manufacturing'
        }
        
        # Mots-clés de CONSOMMATION SPÉCIFIQUE à INCLURE
        consumption_keywords = {
            'energy', 'energie', 'electrical', 'electrique', 
            'water', 'eau', 'steam', 'vapeur', 'vapor',
            'gas', 'gaz', 'fuel', 'combustible', 'coal', 'charbon',
            'oil', 'huile', 'electricity', 'electricite',
            'thermal', 'thermique', 'cooling', 'refroidissement',
            'compressed air', 'air comprime', 'nitrogen', 'azote',
            'oxygen', 'oxygene', 'chemical', 'chimique',
            'raw material', 'matiere premiere', 'catalyst', 'catalyseur'
        }
        
        # Si la colonne contient des mots-clés de processus, on l'exclut
        if any(keyword in column_lower for keyword in process_keywords):
            return False
        
        # Si la colonne contient des mots-clés de consommation, on l'inclut
        if any(keyword in column_lower for keyword in consumption_keywords):
            return True
        
        # Par défaut, on inclut (pour les cas non prévus)
        return True
    
    # ---------------------------------------------------------------------------
    # Module 1: PPV Production
    # ---------------------------------------------------------------------------
    
    def generate_ppv_production(self, uploaded_file, exercice: str, date_version):
        """Génère le DataFrame et l'Excel pour PPV Production."""
        xls = pd.ExcelFile(uploaded_file)
        annee = pd.to_datetime(date_version).year
        all_rows = []

        # Extraction ----------------------------------------------------------------
        extraction_sheet = next((s for s in xls.sheet_names if "extraction" in s.lower()), None)
        extraction_df = pd.DataFrame()
        if extraction_sheet:
            extraction_df = self.extraire_table_par_nom(xls.parse(extraction_sheet, header=None), "ExtractionVolume")

        for _, row in extraction_df.iterrows():
            all_rows.extend(
                self._process_volume_row(
                    row,
                    prefix="Volume",
                    site_field="Mine",
                    qual_field="RawRock",
                    operation_label="Extraction",
                    exercice=exercice,
                    date_version=date_version,
                    annee=annee,
                )
            )

        # Traitements physiques ------------------------------------------------------
        physical_sheet = next((s for s in xls.sheet_names if "physical" in s.lower() or "physique" in s.lower()), None)
        physical_df = pd.DataFrame()
        if physical_sheet:
            physical_df = self.extraire_table_par_nom(xls.parse(physical_sheet, header=None), "VolumeOutputProduct")

        for _, row in physical_df.iterrows():
            all_rows.extend(
                self._process_volume_row(
                    row,
                    prefix="VolumeOutputProduct",
                    site_field="Entity",
                    qual_field="Output",
                    operation_label="Traitements Physiques",
                    exercice=exercice,
                    date_version=date_version,
                    annee=annee,
                )
            )

        # Traitements chimiques ------------------------------------------------------
        chemical_sheet = next((s for s in xls.sheet_names if "chemical" in s.lower() or "chimique" in s.lower()), None)
        chemical_df = pd.DataFrame()
        if chemical_sheet:
            chemical_df = self.extraire_table_par_nom(xls.parse(chemical_sheet, header=None), "Production")

        for _, row in chemical_df.iterrows():
            all_rows.extend(
                self._process_volume_row(
                    row,
                    prefix="Production",
                    site_field="Entity",
                    qual_field="Product",
                    operation_label="Traitements Chimiques",
                    exercice=exercice,
                    date_version=date_version,
                    annee=annee,
                )
            )

        if not all_rows:
            st.error("Aucune donnée trouvée pour PPV Production.")
            return pd.DataFrame(), b""

        df_final = pd.DataFrame(all_rows)
        df_final.insert(0, "#ID", [f"#{i+1}" for i in range(len(df_final))])

        # Export Excel avec feuilles séparées pour chaque table source
        print("📊 Génération fichier Excel PPV Production...")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # Feuille 1: Fichier plat (avec formatage coloré)
            df_final.to_excel(writer, index=False, sheet_name="Fichier plat PPV Production")
            workbook = writer.book
            worksheet = writer.sheets["Fichier plat PPV Production"]
            format_extraction = workbook.add_format({"bg_color": "#E8F4FD"})
            format_physical = workbook.add_format({"bg_color": "#E8F5E8"})
            format_chemical = workbook.add_format({"bg_color": "#FFF8E1"})
            for row_idx, op in enumerate(df_final["Opération"], start=1):
                fmt = format_extraction if op == "Extraction" else format_physical if op == "Traitements Physiques" else format_chemical
                worksheet.set_row(row_idx, cell_format=fmt)
            
            # Feuilles des tables sources (avec structure originale)
            if not extraction_df.empty:
                extraction_df.to_excel(writer, index=False, sheet_name="ExtractionVolume")
                print(f"✅ Table source : ExtractionVolume ({len(extraction_df)} lignes)")
            
            if not physical_df.empty:
                physical_df.to_excel(writer, index=False, sheet_name="VolumeOutputProduct")
                print(f"✅ Table source : VolumeOutputProduct ({len(physical_df)} lignes)")
            
            if not chemical_df.empty:
                chemical_df.to_excel(writer, index=False, sheet_name="Production")
                print(f"✅ Table source : Production ({len(chemical_df)} lignes)")
            
            # NEW: Coloration des onglets selon le type d'opération (Bloc)
            tab_color_map = {
                "ExtractionVolume": "#E8F4FD",      # Extraction – Bleu clair
                "VolumeOutputProduct": "#E8F5E8",   # Physiques – Vert clair
                "Production": "#FFF8E1",            # Chimiques – Jaune clair
            }
            for sheet_name, color in tab_color_map.items():
                if sheet_name in writer.sheets:
                    writer.sheets[sheet_name].set_tab_color(color)
        
        return df_final, output.getvalue()

    def generate_ppv_production_v2(self, uploaded_file, exercice: str, date_version):
        """Génère le DataFrame et l'Excel pour PPV Production."""
        xls = pd.ExcelFile(uploaded_file)
        annee = pd.to_datetime(date_version).year
        all_rows = []

        # Extraction ----------------------------------------------------------------
        extraction_sheet = next((s for s in xls.sheet_names if "extraction" in s.lower()), None)
        extraction_df = pd.DataFrame()
        if extraction_sheet:
            extraction_df = self.extraire_table_par_nom(xls.parse(extraction_sheet, header=None), "ExtractionVolume")

        for _, row in extraction_df.iterrows():
            all_rows.extend(
                self._process_volume_row(
                    row,
                    prefix="Volume",
                    site_field="Mine",
                    qual_field="RawRock",
                    operation_label="Extraction",
                    exercice=exercice,
                    date_version=date_version,
                    annee=annee,
                )
            )

        # Traitements physiques ------------------------------------------------------
        physical_sheet = next((s for s in xls.sheet_names if "physical" in s.lower() or "physique" in s.lower()), None)
        physical_df = pd.DataFrame()
        if physical_sheet:
            physical_df = self.extraire_table_par_nom(xls.parse(physical_sheet, header=None), "VolumeOutputProduct")

        for _, row in physical_df.iterrows():
            all_rows.extend(
                self._process_volume_row(
                    row,
                    prefix="VolumeOutputProduct",
                    site_field="Entity",
                    qual_field="Output",
                    operation_label="Traitements Physiques",
                    exercice=exercice,
                    date_version=date_version,
                    annee=annee,
                )
            )

        # Traitements chimiques ------------------------------------------------------
        chemical_sheet = next((s for s in xls.sheet_names if "chemical" in s.lower() or "chimique" in s.lower()), None)
        chemical_df = pd.DataFrame()
        if chemical_sheet:
            chemical_df = self.extraire_table_par_nom(xls.parse(chemical_sheet, header=None), "Production")

        for _, row in chemical_df.iterrows():
            all_rows.extend(
                self._process_volume_row(
                    row,
                    prefix="Production",
                    site_field="Entity",
                    qual_field="Product",
                    operation_label="Traitements Chimiques",
                    exercice=exercice,
                    date_version=date_version,
                    annee=annee,
                )
            )

        if not all_rows:
            st.error("Aucune donnée trouvée pour PPV Production.")
            return pd.DataFrame(), b""

        df_final = pd.DataFrame(all_rows)
        df_final.insert(0, "#ID", [f"#{i + 1}" for i in range(len(df_final))])

        # Export Excel avec feuilles séparées pour chaque table source
        print("📊 Génération fichier Excel PPV Production...")

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # Feuille 1: Fichier plat (avec formatage coloré)
            df_final.to_excel(writer, index=False, sheet_name="Fichier plat PPV Production")
            workbook = writer.book
            worksheet = writer.sheets["Fichier plat PPV Production"]
            format_extraction = workbook.add_format({"bg_color": "#E8F4FD"})
            format_physical = workbook.add_format({"bg_color": "#E8F5E8"})
            format_chemical = workbook.add_format({"bg_color": "#FFF8E1"})
            for row_idx, op in enumerate(df_final["Type Operation"], start=1):
                fmt = format_extraction if op == "Extraction" else format_physical if op == "Traitements Physiques" else format_chemical
                worksheet.set_row(row_idx, cell_format=fmt)

            # Feuilles des tables sources (avec structure originale)
            if not extraction_df.empty:
                extraction_df.to_excel(writer, index=False, sheet_name="ExtractionVolume")
                print(f"✅ Table source : ExtractionVolume ({len(extraction_df)} lignes)")

            if not physical_df.empty:
                physical_df.to_excel(writer, index=False, sheet_name="VolumeOutputProduct")
                print(f"✅ Table source : VolumeOutputProduct ({len(physical_df)} lignes)")

            if not chemical_df.empty:
                chemical_df.to_excel(writer, index=False, sheet_name="Production")
                print(f"✅ Table source : Production ({len(chemical_df)} lignes)")

            # NEW: Coloration des onglets selon le type d'opération (Bloc)
            tab_color_map = {
                "ExtractionVolume": "#E8F4FD",  # Extraction – Bleu clair
                "VolumeOutputProduct": "#E8F5E8",  # Physiques – Vert clair
                "Production": "#FFF8E1",  # Chimiques – Jaune clair
            }
            for sheet_name, color in tab_color_map.items():
                if sheet_name in writer.sheets:
                    writer.sheets[sheet_name].set_tab_color(color)

        return df_final, output.getvalue()

    def generate_ppv_production_v3(self, ppv_file, uploaded_file, exercice: str, date_version):
        """Génère le DataFrame et l'Excel pour PPV Production."""
        xls = pd.ExcelFile(uploaded_file)
        xls_ppv = pd.ExcelFile(ppv_file)
        annee = pd.to_datetime(date_version).year
        all_rows = []

        # Extraction ----------------------------------------------------------------
        # NOUVELLE LOGIQUE : Lire depuis OIK, OIB, OIG
        extraction_sheets = ["OIK", "OIB", "OIG"]
        extraction_data_list = []

        for sheet_name in extraction_sheets:
            if sheet_name not in xls_ppv.sheet_names:
                print(f"⚠️ Feuille {sheet_name} non trouvée, ignorée")
                continue

            try:
                # Lire la feuille brute
                raw_sheet = xls_ppv.parse(sheet_name, header=None)

                # 1. NOUVEAU : Trouver la ligne avec "Volumes extraits" pour identifier les colonnes de périodes
                periode_row_idx = None
                periode_columns = {}  # {colonne_index: nom_periode}

                for idx, row in raw_sheet.iterrows():
                    # Chercher "Volumes extraits (tonnages equivalents SM)"
                    if any("volumes extraits" in str(cell).lower() and "tonnages" in str(cell).lower()
                           for cell in row if pd.notna(cell)):
                        # La ligne suivante ou +1/+2 lignes contient les périodes
                        # Vérifier les 3 prochaines lignes
                        for offset in range(1, 4):
                            if idx + offset < len(raw_sheet):
                                potential_periode_row = raw_sheet.iloc[idx + offset]
                                # Vérifier si cette ligne contient des noms de mois ou trimestres
                                has_periode = any(
                                    str(cell).lower().strip() in [
                                        "janvier", "février", "fevrier", "mars", "avril", "mai", "juin",
                                        "juillet", "août", "aout", "septembre", "octobre", "novembre",
                                        "décembre", "decembre", "q1", "q2", "q3", "q4"
                                    ] for cell in potential_periode_row if pd.notna(cell)
                                )
                                if has_periode:
                                    periode_row_idx = idx + offset
                                    print(
                                        f"✅ Ligne des périodes trouvée dans {sheet_name} à la ligne {periode_row_idx}")

                                    # Extraire le mapping colonne → période
                                    month_mapping = {
                                        "janvier": 1, "février": 2, "fevrier": 2, "mars": 3,
                                        "avril": 4, "mai": 5, "juin": 6,
                                        "juillet": 7, "août": 8, "aout": 8, "septembre": 9,
                                        "octobre": 10, "novembre": 11, "décembre": 12, "decembre": 12
                                    }

                                    for col_idx, cell_value in enumerate(potential_periode_row):
                                        if pd.notna(cell_value):
                                            cell_lower = str(cell_value).lower().strip()

                                            # Vérifier si c'est un mois
                                            for mois_name, mois_num in month_mapping.items():
                                                if mois_name in cell_lower:
                                                    periode_columns[col_idx] = ("mois", mois_num)
                                                    print(f"   📅 Colonne {col_idx} = Mois {mois_num} ({cell_value})")
                                                    break

                                            # Vérifier si c'est un trimestre
                                            cell_upper = str(cell_value).upper().strip()
                                            if cell_upper in ["Q1", "Q2", "Q3", "Q4"]:
                                                periode_columns[col_idx] = ("trimestre", cell_upper)
                                                print(f"   📅 Colonne {col_idx} = Trimestre {cell_upper}")

                                    break
                        break

                if not periode_columns:
                    print(f"⚠️ Aucune colonne de période détectée dans {sheet_name}")
                    continue

                # 2. Trouver la section "Detail par qualite"
                detail_start_idx = None
                detail_end_idx = None

                for idx, row in raw_sheet.iterrows():
                    # Chercher "Detail par qualite"
                    if any("detail par qualite" in str(cell).lower() for cell in row if pd.notna(cell)):
                        detail_start_idx = idx
                        print(f"✅ 'Detail par qualite' trouvé dans {sheet_name} à la ligne {idx}")
                        continue

                    # Chercher la fin (ligne "Lavage" ou "dont reprise")
                    if detail_start_idx is not None and detail_end_idx is None:
                        if any("lavage" in str(cell).lower() or "dont reprise" in str(cell).lower()
                               for cell in row if pd.notna(cell)):
                            detail_end_idx = idx
                            print(f"🛑 Fin de section détectée dans {sheet_name} à la ligne {idx}")
                            break

                if detail_start_idx is None:
                    print(f"⚠️ Section 'Detail par qualite' non trouvée dans {sheet_name}")
                    continue

                if detail_end_idx is None:
                    detail_end_idx = len(raw_sheet)

                # 3. Extraire la section "Detail par qualite"
                detail_section = raw_sheet.iloc[detail_start_idx:detail_end_idx].copy()

                # La première ligne après "Detail par qualite" contient les headers
                headers_idx = 0
                for idx in range(len(detail_section)):
                    row = detail_section.iloc[idx]
                    if any("secteur" in str(cell).lower() for cell in row if pd.notna(cell)):
                        headers_idx = idx
                        break

                # Identifier la colonne "Secteur" et "Qualite"
                headers_row = detail_section.iloc[headers_idx]
                secteur_col_idx = None
                qualite_col_idx = None

                for col_idx, cell_value in enumerate(headers_row):
                    if pd.notna(cell_value):
                        cell_lower = str(cell_value).lower().strip()
                        if "secteur" in cell_lower:
                            secteur_col_idx = col_idx
                        elif "qualite" in cell_lower or "qualité" in cell_lower:
                            qualite_col_idx = col_idx

                print(f"📋 Secteur colonne: {secteur_col_idx}, Qualite colonne: {qualite_col_idx}")

                # 4. Traiter les données ligne par ligne
                data_start_idx = headers_idx + 1
                current_secteur = None

                for row_idx in range(data_start_idx, len(detail_section)):
                    row = detail_section.iloc[row_idx]

                    # Vérifier si ligne vide complète
                    if row.isna().all():
                        continue

                    # Récupérer Secteur et Qualité
                    secteur_val = row.iloc[secteur_col_idx] if secteur_col_idx is not None else None
                    qualite_val = row.iloc[qualite_col_idx] if qualite_col_idx is not None else None

                    # Mettre à jour le secteur courant si non vide (forward-fill)
                    if pd.notna(secteur_val) and str(secteur_val).strip() != "":
                        current_secteur = str(secteur_val).strip()

                    # Ignorer les lignes sans qualité
                    if pd.isna(qualite_val) or str(qualite_val).strip() == "":
                        continue

                    # Utiliser le secteur courant
                    secteur_to_use = current_secteur if current_secteur else ""
                    qualite_to_use = str(qualite_val).strip()

                    print(f"🔍 Traitement : Secteur='{secteur_to_use}' | Qualite='{qualite_to_use}'")

                    # 5. Extraire les volumes en utilisant le mapping des périodes
                    volumes_traites = []  # Pour éviter les doublons de mois

                    # Parcourir les colonnes identifiées comme périodes
                    for col_idx, (periode_type, periode_value) in periode_columns.items():
                        if col_idx >= len(row):
                            continue

                        cell_value = row.iloc[col_idx]

                        if pd.isna(cell_value) or cell_value == "" or cell_value == "-":
                            continue

                        try:
                            volume_value = float(cell_value)

                            if volume_value == 0:
                                continue

                            if periode_type == "mois":
                                # Volume mensuel direct
                                mois_num = periode_value
                                volume_mensuel = volume_value * 1000  # Conversion Kt → T
                                volumes_traites.append(mois_num)

                                all_rows.append({
                                    "Exercice": exercice,
                                    "Date de la Version": str(date_version),
                                    "Année": annee,
                                    "Mois": mois_num,
                                    "Site/Entité": secteur_to_use,
                                    "Qualité": qualite_to_use,
                                    "Partenaire Groupe": "",
                                    "Type Operation": "Extraction",
                                    "Operation": "",
                                    "VOLUME (T)": volume_mensuel,
                                    "_source": sheet_name
                                })
                                print(f"   ✅ Mois {mois_num} : {volume_mensuel} T")

                            elif periode_type == "trimestre":
                                # Répartir le volume trimestriel sur 3 mois
                                trimestre = periode_value
                                volume_mensuel = (volume_value * 1000) / 3  # Conversion Kt → T et division par 3

                                months = self.quarter_to_months[trimestre]
                                for mois in months:
                                    if mois not in volumes_traites:  # Éviter doublons
                                        all_rows.append({
                                            "Exercice": exercice,
                                            "Date de la Version": str(date_version),
                                            "Année": annee,
                                            "Mois": mois,
                                            "Site/Entité": secteur_to_use,
                                            "Qualité": qualite_to_use,
                                            "Partenaire Groupe": "",
                                            "Type Operation": "Extraction",
                                            "Operation": "",
                                            "VOLUME (T)": volume_mensuel,
                                            "_source": sheet_name
                                        })
                                        print(f"   ✅ {trimestre} → Mois {mois} : {volume_mensuel} T")

                        except (ValueError, TypeError) as e:
                            print(f"   ⚠️ Erreur conversion colonne {col_idx} : {e}")

                # Sauvegarder la section pour export
                extraction_data_list.append(detail_section.iloc[headers_idx:])
                print(f"✅ {sheet_name} traité avec succès")

            except Exception as e:
                print(f"❌ Erreur traitement {sheet_name} : {e}")
                import traceback
                traceback.print_exc()

        # Traitements physiques ------------------------------------------------------
        physical_sheet = next((s for s in xls.sheet_names if "physical" in s.lower() or "physique" in s.lower()), None)
        physical_df = pd.DataFrame()
        if physical_sheet:
            physical_df = self.extraire_table_par_nom(xls.parse(physical_sheet, header=None), "VolumeOutputProduct")

        for _, row in physical_df.iterrows():
            all_rows.extend(
                self._process_volume_row(
                    row,
                    prefix="VolumeOutputProduct",
                    site_field="Entity",
                    qual_field="Output",
                    operation_label="Traitements Physiques",
                    exercice=exercice,
                    date_version=date_version,
                    annee=annee,
                )
            )

        # Traitements chimiques ------------------------------------------------------
        chemical_sheet = next((s for s in xls.sheet_names if "chemical" in s.lower() or "chimique" in s.lower()), None)
        chemical_df = pd.DataFrame()
        if chemical_sheet:
            chemical_df = self.extraire_table_par_nom(xls.parse(chemical_sheet, header=None), "Production")

        for _, row in chemical_df.iterrows():
            all_rows.extend(
                self._process_volume_row(
                    row,
                    prefix="Production",
                    site_field="Entity",
                    qual_field="Product",
                    operation_label="Traitements Chimiques",
                    exercice=exercice,
                    date_version=date_version,
                    annee=annee,
                )
            )

        if not all_rows:
            st.error("Aucune donnée trouvée pour PPV Production.")
            return pd.DataFrame(), b""

        df_final = pd.DataFrame(all_rows)
        df_final.insert(0, "#ID", [f"#{i + 1}" for i in range(len(df_final))])

        # Export Excel avec feuilles séparées pour chaque table source
        print("📊 Génération fichier Excel PPV Production...")

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # Feuille 1: Fichier plat (avec formatage coloré)
            df_final.to_excel(writer, index=False, sheet_name="Fichier plat PPV Production")
            workbook = writer.book
            worksheet = writer.sheets["Fichier plat PPV Production"]
            format_extraction = workbook.add_format({"bg_color": "#E8F4FD"})
            format_physical = workbook.add_format({"bg_color": "#E8F5E8"})
            format_chemical = workbook.add_format({"bg_color": "#FFF8E1"})
            for row_idx, op in enumerate(df_final["Type Operation"], start=1):
                fmt = format_extraction if op == "Extraction" else format_physical if op == "Traitements Physiques" else format_chemical
                worksheet.set_row(row_idx, cell_format=fmt)

            # Feuilles des tables sources
            for i, extraction_source_df in enumerate(extraction_data_list):
                if not extraction_source_df.empty:
                    sheet_name = f"Extraction_{extraction_sheets[i]}"
                    extraction_source_df.to_excel(writer, index=False, sheet_name=sheet_name)
                    print(f"✅ Table source : {sheet_name} ({len(extraction_source_df)} lignes)")

            if not physical_df.empty:
                physical_df.to_excel(writer, index=False, sheet_name="VolumeOutputProduct")
                print(f"✅ Table source : VolumeOutputProduct ({len(physical_df)} lignes)")

            if not chemical_df.empty:
                chemical_df.to_excel(writer, index=False, sheet_name="Production")
                print(f"✅ Table source : Production ({len(chemical_df)} lignes)")

            # Coloration des onglets selon le type d'opération
            tab_color_map = {
                "Extraction_OIK": "#E8F4FD",
                "Extraction_OIB": "#E8F4FD",
                "Extraction_OIG": "#E8F4FD",
                "VolumeOutputProduct": "#E8F5E8",
                "Production": "#FFF8E1",
            }
            for sheet_name, color in tab_color_map.items():
                if sheet_name in writer.sheets:
                    writer.sheets[sheet_name].set_tab_color(color)

        return df_final, output.getvalue()

    def generate_ppv_production_v4(self, ppv_file, uploaded_file, input_file, exercice: str, date_version):
        """Génère le DataFrame et l'Excel pour PPV Production."""
        xls = pd.ExcelFile(uploaded_file)
        xls_ppv = pd.ExcelFile(ppv_file)
        xls_input = pd.ExcelFile(input_file)
        annee = pd.to_datetime(date_version).year
        all_rows = []

        # Extraction ----------------------------------------------------------------
        # NOUVELLE LOGIQUE : Lire depuis OIK, OIB, OIG
        extraction_sheets = ["OIK", "OIB", "OIG"]
        extraction_data_list = []

        for sheet_name in extraction_sheets:
            if sheet_name not in xls_ppv.sheet_names:
                print(f"⚠️ Feuille {sheet_name} non trouvée, ignorée")
                continue

            try:
                # Lire la feuille brute
                raw_sheet = xls_ppv.parse(sheet_name, header=None)

                # 1. NOUVEAU : Trouver la ligne avec "Volumes extraits" pour identifier les colonnes de périodes
                periode_row_idx = None
                periode_columns = {}  # {colonne_index: nom_periode}

                for idx, row in raw_sheet.iterrows():
                    # Chercher "Volumes extraits (tonnages equivalents SM)"
                    if any("volumes extraits" in str(cell).lower() and "tonnages" in str(cell).lower()
                           for cell in row if pd.notna(cell)):
                        # La ligne suivante ou +1/+2 lignes contient les périodes
                        # Vérifier les 3 prochaines lignes
                        for offset in range(1, 4):
                            if idx + offset < len(raw_sheet):
                                potential_periode_row = raw_sheet.iloc[idx + offset]
                                # Vérifier si cette ligne contient des noms de mois ou trimestres
                                has_periode = any(
                                    str(cell).lower().strip() in [
                                        "janvier", "février", "fevrier", "mars", "avril", "mai", "juin",
                                        "juillet", "août", "aout", "septembre", "octobre", "novembre",
                                        "décembre", "decembre", "q1", "q2", "q3", "q4"
                                    ] for cell in potential_periode_row if pd.notna(cell)
                                )
                                if has_periode:
                                    periode_row_idx = idx + offset
                                    print(
                                        f"✅ Ligne des périodes trouvée dans {sheet_name} à la ligne {periode_row_idx}")

                                    # Extraire le mapping colonne → période
                                    month_mapping = {
                                        "janvier": 1, "février": 2, "fevrier": 2, "mars": 3,
                                        "avril": 4, "mai": 5, "juin": 6,
                                        "juillet": 7, "août": 8, "aout": 8, "septembre": 9,
                                        "octobre": 10, "novembre": 11, "décembre": 12, "decembre": 12
                                    }

                                    for col_idx, cell_value in enumerate(potential_periode_row):
                                        if pd.notna(cell_value):
                                            cell_lower = str(cell_value).lower().strip()

                                            # Vérifier si c'est un mois
                                            for mois_name, mois_num in month_mapping.items():
                                                if mois_name in cell_lower:
                                                    periode_columns[col_idx] = ("mois", mois_num)
                                                    print(f"   📅 Colonne {col_idx} = Mois {mois_num} ({cell_value})")
                                                    break

                                            # Vérifier si c'est un trimestre
                                            cell_upper = str(cell_value).upper().strip()
                                            if cell_upper in ["Q1", "Q2", "Q3", "Q4"]:
                                                periode_columns[col_idx] = ("trimestre", cell_upper)
                                                print(f"   📅 Colonne {col_idx} = Trimestre {cell_upper}")

                                    break
                        break

                if not periode_columns:
                    print(f"⚠️ Aucune colonne de période détectée dans {sheet_name}")
                    continue

                # 2. Trouver la section "Detail par qualite"
                detail_start_idx = None
                detail_end_idx = None

                for idx, row in raw_sheet.iterrows():
                    # Chercher "Detail par qualite"
                    if any("detail par qualite" in str(cell).lower() for cell in row if pd.notna(cell)):
                        detail_start_idx = idx
                        print(f"✅ 'Detail par qualite' trouvé dans {sheet_name} à la ligne {idx}")
                        continue

                    # Chercher la fin (ligne "Lavage" ou "dont reprise")
                    if detail_start_idx is not None and detail_end_idx is None:
                        if any("lavage" in str(cell).lower() or "dont reprise" in str(cell).lower()
                               for cell in row if pd.notna(cell)):
                            detail_end_idx = idx
                            print(f"🛑 Fin de section détectée dans {sheet_name} à la ligne {idx}")
                            break

                if detail_start_idx is None:
                    print(f"⚠️ Section 'Detail par qualite' non trouvée dans {sheet_name}")
                    continue

                if detail_end_idx is None:
                    detail_end_idx = len(raw_sheet)

                # 3. Extraire la section "Detail par qualite"
                detail_section = raw_sheet.iloc[detail_start_idx:detail_end_idx].copy()

                # La première ligne après "Detail par qualite" contient les headers
                headers_idx = 0
                for idx in range(len(detail_section)):
                    row = detail_section.iloc[idx]
                    if any("secteur" in str(cell).lower() for cell in row if pd.notna(cell)):
                        headers_idx = idx
                        break

                # Identifier la colonne "Secteur" et "Qualite"
                headers_row = detail_section.iloc[headers_idx]
                secteur_col_idx = None
                qualite_col_idx = None

                for col_idx, cell_value in enumerate(headers_row):
                    if pd.notna(cell_value):
                        cell_lower = str(cell_value).lower().strip()
                        if "secteur" in cell_lower:
                            secteur_col_idx = col_idx
                        elif "qualite" in cell_lower or "qualité" in cell_lower:
                            qualite_col_idx = col_idx

                print(f"📋 Secteur colonne: {secteur_col_idx}, Qualite colonne: {qualite_col_idx}")

                # 4. Traiter les données ligne par ligne
                data_start_idx = headers_idx + 1
                current_secteur = None

                for row_idx in range(data_start_idx, len(detail_section)):
                    row = detail_section.iloc[row_idx]

                    # Vérifier si ligne vide complète
                    if row.isna().all():
                        continue

                    # Récupérer Secteur et Qualité
                    secteur_val = row.iloc[secteur_col_idx] if secteur_col_idx is not None else None
                    qualite_val = row.iloc[qualite_col_idx] if qualite_col_idx is not None else None

                    # Mettre à jour le secteur courant si non vide (forward-fill)
                    if pd.notna(secteur_val) and str(secteur_val).strip() != "":
                        current_secteur = str(secteur_val).strip()

                    # Ignorer les lignes sans qualité
                    if pd.isna(qualite_val) or str(qualite_val).strip() == "":
                        continue

                    # Utiliser le secteur courant
                    secteur_to_use = current_secteur if current_secteur else ""
                    qualite_to_use = str(qualite_val).strip()

                    print(f"🔍 Traitement : Secteur='{secteur_to_use}' | Qualite='{qualite_to_use}'")

                    # 5. Extraire les volumes en utilisant le mapping des périodes
                    volumes_traites = []  # Pour éviter les doublons de mois

                    # Parcourir les colonnes identifiées comme périodes
                    for col_idx, (periode_type, periode_value) in periode_columns.items():
                        if col_idx >= len(row):
                            continue

                        cell_value = row.iloc[col_idx]

                        if pd.isna(cell_value) or cell_value == "" or cell_value == "-":
                            continue

                        try:
                            volume_value = float(cell_value)

                            if volume_value == 0:
                                continue

                            if periode_type == "mois":
                                # Volume mensuel direct
                                mois_num = periode_value
                                volume_mensuel = volume_value * 1000  # Conversion Kt → T
                                volumes_traites.append(mois_num)

                                all_rows.append({
                                    "Exercice": exercice,
                                    "Date de la Version": str(date_version),
                                    "Année": annee,
                                    "Mois": mois_num,
                                    "Site/Entité": secteur_to_use,
                                    "Qualité": qualite_to_use,
                                    "Partenaire Groupe": "",
                                    "Type Operation": "Extraction",
                                    "Operation": "",
                                    "VOLUME (T)": volume_mensuel,
                                    "_source": sheet_name
                                })
                                print(f"   ✅ Mois {mois_num} : {volume_mensuel} T")

                            elif periode_type == "trimestre":
                                # Répartir le volume trimestriel sur 3 mois
                                trimestre = periode_value
                                volume_mensuel = (volume_value * 1000) / 3  # Conversion Kt → T et division par 3

                                months = self.quarter_to_months[trimestre]
                                for mois in months:
                                    if mois not in volumes_traites:  # Éviter doublons
                                        all_rows.append({
                                            "Exercice": exercice,
                                            "Date de la Version": str(date_version),
                                            "Année": annee,
                                            "Mois": mois,
                                            "Site/Entité": secteur_to_use,
                                            "Qualité": qualite_to_use,
                                            "Partenaire Groupe": "",
                                            "Type Operation": "Extraction",
                                            "Operation": "",
                                            "VOLUME (T)": volume_mensuel,
                                            "_source": sheet_name
                                        })
                                        print(f"   ✅ {trimestre} → Mois {mois} : {volume_mensuel} T")

                        except (ValueError, TypeError) as e:
                            print(f"   ⚠️ Erreur conversion colonne {col_idx} : {e}")

                # Sauvegarder la section pour export
                extraction_data_list.append(detail_section.iloc[headers_idx:])
                print(f"✅ {sheet_name} traité avec succès")

            except Exception as e:
                print(f"❌ Erreur traitement {sheet_name} : {e}")
                import traceback
                traceback.print_exc()

        # Traitements physiques ------------------------------------------------------
        print("\n" + "=" * 80)
        print("🔧 TRAITEMENT PHYSIQUE - Début du processus")
        print("=" * 80)

        physical_sheet = next((s for s in xls.sheet_names if "physical" in s.lower() or "physique" in s.lower()), None)
        physical_df = pd.DataFrame()

        if physical_sheet:
            # Étape 1 : Lecture des tables
            print("\n📖 Étape 1 : Lecture des tables")
            physical_df = self.extraire_table_par_nom(xls.parse(physical_sheet, header=None), "VolumeOutputProduct")

            # Exclusion des produits pour traitements physiques
            try:
                exclusion_df = self.extraire_table_par_nom(xls.parse(physical_sheet, header=None),
                                                           "Production_Exclusion")
                if not exclusion_df.empty and 'Product' in exclusion_df.columns:
                    excluded_products = exclusion_df['Product'].dropna().tolist()
                    physical_df = physical_df[~physical_df['Output'].isin(excluded_products)]
                    print(f" Produits exclus (Physical) : {excluded_products}")
            except:
                pass

            mix_blending_df = self.extraire_table_par_nom(xls.parse(physical_sheet, header=None), "MixForBlending")

            print(f"   ✅ VolumeOutputProduct : {len(physical_df)} lignes")
            print(f"   ✅ MixForBlending : {len(mix_blending_df)} lignes")

            # Lecture de la TreatmentMatrix depuis input_file
            treatment_matrix_df = pd.DataFrame()

            try:
                treatment_matrix_df = self.extraire_table_par_nom(
                    xls_input.parse("PhysicalTreatments", header=None),
                    "TreatmentMatrix / TraitementOIK / HS"
                )
                if not treatment_matrix_df.empty:
                    print(f"   La table #TreatmentMatrix / TraitementOIK / HS trouvée dans PhysicalTreatments avec : {len(treatment_matrix_df)} lignes")
            except:
                print(f"   La table #TreatmentMatrix / TraitementOIK / HS non trouvée dans PhysicalTreatments")

            if treatment_matrix_df.empty:
                print("   ⚠️ TreatmentMatrix non trouvée, traitement sans Yield")

            # Vérification des tables nécessaires
            if not physical_df.empty and not mix_blending_df.empty:
                print("\n🔗 Étape 2 : Merge 1 - VolumeOutputProduct ⬅ MixForBlending")

                # Avant le merge, dédupliquer MixForBlending
                mix_blending_df_dedup = mix_blending_df.drop_duplicates(
                    subset=['Entity', 'Input'],
                    keep='first'
                )

                # Merge 1 : LEFT JOIN sur Entity + Output/Input
                merged_df = pd.merge(
                    physical_df,
                    mix_blending_df_dedup,
                    left_on=['Entity', 'Output'],
                    right_on=['Entity', 'Input'],
                    how='left',
                    suffixes=('_volumeoutputproduct', '_mixforblending')
                )

                print(f"   ✅ Résultat Merge 1 : {len(merged_df)} lignes")

                # Identifier les colonnes de volumes
                volume_cols = [col for col in physical_df.columns if str(col).startswith('VolumeOutputProduct[')]
                print(f"   📊 Colonnes de volumes détectées : {volume_cols}")

                # Étape 3 : Traitement selon Output_right (Output_mixforblending)
                print("\n⚙️ Étape 3 : Gestion du cas Output_right et application du Yield")

                for idx, row in merged_df.iterrows():
                    output_right = row.get('Output_mixforblending', None)
                    entity = row.get('Entity', row.get('Entity_volumeoutputproduct', ''))
                    output_left = row.get('Output', row.get('Output_volumeoutputproduct', ''))

                    # Cas 1 : Output_right est NAN → pas de Yield, on garde les volumes originaux
                    if pd.isna(output_right) or output_right == '':
                        print(
                            f"   ℹ️ Ligne {idx} : Output_right=NAN → Volume original conservé (Entity={entity}, Output={output_left})")
                        yield_to_apply = 1.0

                    # Cas 2 : Output_right existe → chercher le Yield dans TreatmentMatrix
                    else:
                        print(f"   🔍 Ligne {idx} : Output_right='{output_right}' → Recherche du Yield")

                        if not treatment_matrix_df.empty:
                            # Merge 2 : Chercher le Yield
                            matching_yield = treatment_matrix_df[
                                (treatment_matrix_df['Entity'] == entity) &
                                (treatment_matrix_df['Input'] == output_right)
                                ]

                            if not matching_yield.empty and 'Yield' in matching_yield.columns:
                                yield_value = matching_yield.iloc[0]['Yield']

                                if pd.notna(yield_value):
                                    try:
                                        yield_to_apply = float(yield_value)
                                        print(f"      ✅ Yield trouvé : {yield_to_apply}")
                                    except (ValueError, TypeError):
                                        print(f"      ⚠️ Yield invalide ({yield_value}), utilisation de Yield=1.0")
                                        yield_to_apply = 1.0
                                else:
                                    print(f"      ⚠️ Yield NAN pour Entity={entity}, Input={output_right} → Yield=1.0")
                                    yield_to_apply = 1.0
                            else:
                                print(
                                    f"      ⚠️ Aucun match trouvé dans TreatmentMatrix pour Entity={entity}, Input={output_right} → Yield=1.0")
                                yield_to_apply = 1.0
                        else:
                            print(f"      ⚠️ TreatmentMatrix vide → Yield=1.0")
                            yield_to_apply = 1.0

                    # Application du Yield sur TOUTES les colonnes de volumes
                    if yield_to_apply != 1.0:
                        print(f"      📐 Application du Yield {yield_to_apply} sur les volumes")

                    for vol_col in volume_cols:
                        original_value = row[vol_col]
                        if pd.notna(original_value) and original_value != 0:
                            try:
                                new_value = float(original_value) * yield_to_apply
                                merged_df.at[idx, vol_col] = new_value

                                if yield_to_apply != 1.0:
                                    print(f"         {vol_col}: {original_value} × {yield_to_apply} = {new_value}")
                            except (ValueError, TypeError):
                                print(f"         ⚠️ Impossible de convertir {vol_col}={original_value}")

                # Nettoyer les colonnes pour revenir au format attendu par _process_volume_row
                # On garde uniquement les colonnes originales de VolumeOutputProduct
                columns_to_keep = ['Entity', 'Treatment', 'Output_volumeoutputproduct'] + volume_cols
                physical_df_processed = merged_df[columns_to_keep].copy()

                # # Renommer si nécessaire (enlever les suffixes)
                # physical_df_processed.columns = [
                #     col.replace('_volumeoutputproduct', '') for col in physical_df_processed.columns
                # ]

                # Renommer Output_volumeoutputproduct → Output
                physical_df_processed.rename(columns={
                    'Output_volumeoutputproduct': 'Output'
                }, inplace=True)

                print(f"\n✅ DataFrame final pour traitement : {len(physical_df_processed)} lignes")

            else:
                print("   ⚠️ MixForBlending vide, traitement standard sans Yield")
                physical_df_processed = physical_df.copy()
        else:
            print("⚠️ Feuille Physical/Physique non trouvée")
            physical_df_processed = pd.DataFrame()

        # Traitement des lignes avec _process_volume_row (après application du Yield)
        print("\n🔄 Traitement des volumes avec _process_volume_row")
        for idx, row in physical_df_processed.iterrows():
            processed_rows = self._process_volume_row(
                row,
                prefix="VolumeOutputProduct",
                site_field="Entity",
                qual_field="Output",
                operation_label="Traitements Physiques",
                exercice=exercice,
                date_version=date_version,
                annee=annee,
            )
            all_rows.extend(processed_rows)
            print(f"   ✅ Ligne {idx} : {len(processed_rows)} lignes générées")

        print("=" * 80)
        print("✅ TRAITEMENT PHYSIQUE - Terminé")
        print("=" * 80 + "\n")

        # Traitements chimiques ------------------------------------------------------
        chemical_sheet = next((s for s in xls.sheet_names if "chemical" in s.lower() or "chimique" in s.lower()), None)
        chemical_df = pd.DataFrame()
        if chemical_sheet:
            chemical_df = self.extraire_table_par_nom(xls.parse(chemical_sheet, header=None), "Production")

            # Exclusion des produits pour traitments chimiques
            try:
                exclusion_df = self.extraire_table_par_nom(xls.parse(chemical_sheet, header=None), "Production_Exclusion")
                if not exclusion_df.empty and 'Product' in exclusion_df.columns:
                    excluded_products = exclusion_df['Product'].dropna().tolist()
                    chemical_df = chemical_df[~chemical_df['Product'].isin(excluded_products)]
                    print(f" Produits exclus (Chemical) : {excluded_products}")
            except:
                pass

        for _, row in chemical_df.iterrows():
            all_rows.extend(
                self._process_volume_row(
                    row,
                    prefix="Production",
                    site_field="Entity",
                    qual_field="Product",
                    operation_label="Traitements Chimiques",
                    exercice=exercice,
                    date_version=date_version,
                    annee=annee,
                )
            )

        # Expédition ----------------------------------------------------------------
        print("\n" + "=" * 80)
        print("📦 EXPÉDITION - Début du processus")
        print("=" * 80)

        connexions_sheet = "Connexions"
        expedition_df = pd.DataFrame()

        if connexions_sheet in xls.sheet_names:
            try:
                connexions_raw = xls.parse(connexions_sheet, header=None)
                flow_df = self.extraire_table_par_nom(connexions_raw, "Flow")

                # Filtrer sur Connexion contenant "Expedition"
                if not flow_df.empty and 'Connexion' in flow_df.columns:
                    expedition_df = flow_df[
                        flow_df['Connexion'].str.contains('Expedition', case=False, na=False)
                    ].copy()

                    print(f"✅ Table #Flow : {len(flow_df)} lignes")
                    print(f"✅ Filtre 'Expedition' : {len(expedition_df)} lignes retenues")

                    # Identifier colonnes de volumes Flow[...]
                    flow_volume_cols = [col for col in expedition_df.columns if col.startswith('Flow[')]
                    print(f"📊 Colonnes Flow détectées : {flow_volume_cols}")

                    # Mapping des mois/trimestres
                    month_mapping = {
                        "janvier": 1, "février": 2, "fevrier": 2, "mars": 3,
                        "avril": 4, "mai": 5, "juin": 6,
                        "juillet": 7, "août": 8, "aout": 8, "septembre": 9,
                        "octobre": 10, "novembre": 11, "décembre": 12, "decembre": 12
                    }

                    for idx, row in expedition_df.iterrows():
                        origin = row.get('Origin', '')
                        product = row.get('Product', '')
                        destination = row.get('Destination', '')
                        year = row.get('Year', annee)

                        volumes_traites = []

                        # 1. Colonnes mensuelles
                        for col in flow_volume_cols:
                            col_lower = str(col).lower().strip()

                            # Extraire le nom entre crochets
                            if '[' in col and ']' in col:
                                period_name = col.split('[')[1].split(']')[0].lower().strip()

                                # Vérifier si c'est un mois
                                mois_num = None
                                for mois_name, mois_n in month_mapping.items():
                                    if mois_name in period_name:
                                        mois_num = mois_n
                                        break

                                if mois_num and pd.notna(row[col]):
                                    try:
                                        volume_value = float(row[col])
                                        if volume_value != 0:
                                            volume_mensuel = volume_value * 1000  # Kt → T
                                            volumes_traites.append(mois_num)

                                            all_rows.append({
                                                "Exercice": exercice,
                                                "Date de la Version": str(date_version),
                                                "Année": year,
                                                "Mois": mois_num,
                                                "Site/Entité": origin,
                                                "Qualité": product,
                                                "Partenaire Groupe": destination,
                                                "Type Operation": "Expedition",
                                                "Operation": "",
                                                "VOLUME (T)": volume_mensuel,
                                                "_source": "Flow_Expedition"
                                            })
                                    except (ValueError, TypeError):
                                        pass

                        # 2. Colonnes trimestrielles (pour mois non couverts)
                        for col in flow_volume_cols:
                            col_upper = str(col).upper().strip()

                            if 'Q1' in col_upper or 'Q2' in col_upper or 'Q3' in col_upper or 'Q4' in col_upper:
                                # Extraire Q1/Q2/Q3/Q4
                                for q in ["Q1", "Q2", "Q3", "Q4"]:
                                    if q in col_upper and pd.notna(row[col]):
                                        try:
                                            volume_value = float(row[col])
                                            if volume_value != 0:
                                                volume_mensuel = (volume_value * 1000) / 3  # Kt → T / 3

                                                months = self.quarter_to_months[q]
                                                for mois in months:
                                                    if mois not in volumes_traites:
                                                        all_rows.append({
                                                            "Exercice": exercice,
                                                            "Date de la Version": str(date_version),
                                                            "Année": year,
                                                            "Mois": mois,
                                                            "Site/Entité": origin,
                                                            "Qualité": product,
                                                            "Partenaire Groupe": destination,
                                                            "Type Operation": "Expedition",
                                                            "Operation": "",
                                                            "VOLUME (T)": volume_mensuel,
                                                            "_source": "Flow_Expedition"
                                                        })
                                        except (ValueError, TypeError):
                                            pass
                                        break

                    print(f"✅ Expédition traitée : lignes ajoutées à all_rows")
                else:
                    print("⚠️ Colonne 'Connexion' non trouvée dans #Flow")

            except Exception as e:
                print(f"❌ Erreur traitement Expédition : {e}")
                import traceback
                traceback.print_exc()
        else:
            print("⚠️ Feuille 'Connexions' non trouvée")

        print("=" * 80)
        print("✅ EXPÉDITION - Terminé")
        print("=" * 80 + "\n")

        if not all_rows:
            st.error("Aucune donnée trouvée pour PPV Production.")
            return pd.DataFrame(), b""

        df_final = pd.DataFrame(all_rows)
        df_final.insert(0, "#ID", [f"#{i + 1}" for i in range(len(df_final))])

        # Export Excel avec feuilles séparées pour chaque table source
        print("📊 Génération fichier Excel PPV Production...")

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # Feuille 1: Fichier plat (avec formatage coloré)
            df_final.to_excel(writer, index=False, sheet_name="Fichier plat PPV Production")
            workbook = writer.book
            worksheet = writer.sheets["Fichier plat PPV Production"]
            format_extraction = workbook.add_format({"bg_color": "#E8F4FD"})
            format_physical = workbook.add_format({"bg_color": "#E8F5E8"})
            format_chemical = workbook.add_format({"bg_color": "#FFF8E1"})
            for row_idx, op in enumerate(df_final["Type Operation"], start=1):
                fmt = format_extraction if op == "Extraction" else format_physical if op == "Traitements Physiques" else format_chemical
                worksheet.set_row(row_idx, cell_format=fmt)

            # Feuilles des tables sources
            for i, extraction_source_df in enumerate(extraction_data_list):
                if not extraction_source_df.empty:
                    sheet_name = f"Extraction_{extraction_sheets[i]}"
                    extraction_source_df.to_excel(writer, index=False, sheet_name=sheet_name)
                    print(f"✅ Table source : {sheet_name} ({len(extraction_source_df)} lignes)")

            if not physical_df.empty:
                physical_df.to_excel(writer, index=False, sheet_name="VolumeOutputProduct")
                print(f"✅ Table source : VolumeOutputProduct ({len(physical_df)} lignes)")

            if not chemical_df.empty:
                chemical_df.to_excel(writer, index=False, sheet_name="Production")
                print(f"✅ Table source : Production ({len(chemical_df)} lignes)")

            # Coloration des onglets selon le type d'opération
            tab_color_map = {
                "Extraction_OIK": "#E8F4FD",
                "Extraction_OIB": "#E8F4FD",
                "Extraction_OIG": "#E8F4FD",
                "VolumeOutputProduct": "#E8F5E8",
                "Production": "#FFF8E1",
            }
            for sheet_name, color in tab_color_map.items():
                if sheet_name in writer.sheets:
                    writer.sheets[sheet_name].set_tab_color(color)

        return df_final, output.getvalue()


    # ---------------------------------------------------------------------------
    # Module 2: Ventes & Matières Premières
    # ---------------------------------------------------------------------------
    
    @staticmethod
    def tokenizer_contract(val):
        val = str(val).lower().replace(",", ".")
        tokens = re.split(r"[_\s\.,']", val)
        return {t for t in tokens if t}

    def match_par_tokens(self, volume_tokens, fulfill_df_tokens, seuil=6):
        """Match par tokens avec ordre FulfillData comme critère de départage"""
        best_matches = []
        max_communs = 0
        
        for idx, row in fulfill_df_tokens.iterrows():
            communs = volume_tokens.intersection(row["Contract_tokens"])
            if len(communs) >= seuil:
                if len(communs) > max_communs:
                    max_communs = len(communs)
                    best_matches = [(idx, row["Country"])]
                elif len(communs) == max_communs:
                    best_matches.append((idx, row["Country"]))
        
        if best_matches:
            return best_matches[0][1], best_matches[0][0]
        return "", 9999

    def generate_ppv_ventes_mp(self, uploaded_file, exercice: str, date_version):
        """Génère le DataFrame et l'Excel pour Ventes & MP."""
        xls = pd.ExcelFile(uploaded_file)
        annee = pd.to_datetime(date_version).year
        all_rows = []

        # === BLOC 1 - Vente Export ===
        sheet_sales_name = self.find_sheet(xls, "sales")
        sheet_mps_name = self.find_sheet(xls, "multi", "period", "sales")
        sheet_price_name = self.find_sheet(xls, "price")
        
        if sheet_sales_name is None or sheet_mps_name is None:
            st.error("Feuilles concernant les ventes non trouvées dans le fichier.")
            return pd.DataFrame(), b""
            
        if sheet_price_name is None:
            st.error("Feuille Price non trouvée dans le fichier.")
            return pd.DataFrame(), b""

        feuille_sales = xls.parse(sheet_sales_name, header=None)
        feuille_mps = xls.parse(sheet_mps_name, header=None)
        feuille_price = xls.parse(sheet_price_name, header=None)
        volume_df = self.extraire_table_par_nom(feuille_sales, "Volume")
        fulfill_df = self.extraire_table_par_nom(feuille_mps, "FulfillData")
        contract_price_df = self.extraire_table_par_nom(feuille_price, "ContractPrice")
        
        # DEBUG: Vérifier le contenu de la table Volume
        print(f"🔍 DEBUG Volume extraction:")
        print(f"   - Table Volume trouvée: {not volume_df.empty}")
        print(f"   - Taille: {len(volume_df)} lignes")
        print(f"   - Colonnes: {volume_df.columns.tolist() if not volume_df.empty else 'N/A'}")
        if not volume_df.empty:
            print(f"   - Aperçu des 3 premières lignes:")
            print(volume_df.head(3).to_string())
            if 'Export_or_Local' in volume_df.columns:
                print(f"   - Valeurs uniques dans Export_or_Local: {volume_df['Export_or_Local'].unique()}")
            elif 'Export or local' in volume_df.columns:
                print(f"   - Valeurs uniques dans 'Export or local': {volume_df['Export or local'].unique()}")
            else:
                export_cols = [col for col in volume_df.columns if 'export' in str(col).lower() or 'local' in str(col).lower()]
                print(f"   - Colonnes contenant 'export'/'local': {export_cols}")
        print("-" * 80)
        # DEBUG AMELIORE: Afficher toutes les valeurs de filtrage
        if not volume_df.empty:
            for col in volume_df.columns:
                if 'export' in str(col).lower() or 'local' in str(col).lower():
                    print(f"🔍 DEBUG: Colonne {col} contient: {volume_df[col].unique()}")
                    for val in volume_df[col].unique():
                        count = len(volume_df[volume_df[col] == val])
                        print(f"     '{val}': {count} lignes")
        
        
        # DEBUG: Vérifier les colonnes Volume disponibles
        print(f"🔍 DEBUG BLOC 1: Colonnes Volume disponibles: {[col for col in volume_df.columns if 'Volume' in str(col)]}")
        print(f"🔍 DEBUG BLOC 1: Taille Volume: {len(volume_df)} lignes")
        
        # Filtrer les données Export pour BLOC 1
        if "Export_or_Local" in volume_df.columns:
            volume_df_export = volume_df[volume_df["Export_or_Local"] == "EXPORT"].copy()
            print(f"🔍 DEBUG Volume Export: {len(volume_df_export)} lignes après filtre 'Export'")
        elif "Export or local" in volume_df.columns:
            volume_df_export = volume_df[volume_df["Export or local"].str.upper() == "EXPORT"].copy()
            print(f"🔍 DEBUG Volume Export: {len(volume_df_export)} lignes après filtre 'EXPORT'")
        else:
            volume_df_export = volume_df.copy()
            print("⚠️ DEBUG Volume: Aucune colonne de filtre Export/Local trouvée")
        
        print(f"🔍 DEBUG Volume Export: Entities disponibles = {volume_df_export['Entity'].unique() if 'Entity' in volume_df_export.columns else 'Colonne Entity manquante'}")
        print("-" * 80)
        # DEBUG AMELIORE: Afficher toutes les valeurs de filtrage
        if not volume_df.empty:
            for col in volume_df.columns:
                if 'export' in str(col).lower() or 'local' in str(col).lower():
                    print(f"🔍 DEBUG: Colonne {col} contient: {volume_df[col].unique()}")
                    for val in volume_df[col].unique():
                        count = len(volume_df[volume_df[col] == val])
                        print(f"     '{val}': {count} lignes")
        

        # BLOC 1 - Entités depuis Volume + Pays depuis FulfillData
        # Filtrer seulement les exports dans FulfillData pour récupérer les pays
        if "Export or local" in fulfill_df.columns:
            fulfill_df_export = fulfill_df[fulfill_df["Export or local"] == "EXPORT"].copy()
        else:
            fulfill_df_export = pd.DataFrame()
            
        # DEBUG: Vérifier le contenu de FulfillData Export
        print(f"🔍 DEBUG FulfillData Export:")
        print(f"   - Taille: {len(fulfill_df_export)} lignes")
        print(f"   - Colonnes: {fulfill_df_export.columns.tolist() if not fulfill_df_export.empty else 'N/A'}")
        if not fulfill_df_export.empty:
            print(f"   - Aperçu des 3 premières lignes:")
            print(fulfill_df_export.head(3).to_string())
            if 'Country' in fulfill_df_export.columns:
                print(f"   - Valeurs uniques dans Country: {fulfill_df_export['Country'].unique()}")
            if 'Contract' in fulfill_df_export.columns:
                print(f"   - Exemple de contracts: {fulfill_df_export['Contract'].head(5).tolist()}")
        print("-" * 80)
        
        # DEBUG: Vérifier le contenu de ContractPrice
        print(f"🔍 DEBUG ContractPrice:")
        print(f"   - Taille: {len(contract_price_df)} lignes")
        print(f"   - Colonnes: {contract_price_df.columns.tolist() if not contract_price_df.empty else 'N/A'}")
        if not contract_price_df.empty:
            print(f"   - Aperçu des 3 premières lignes:")
            print(contract_price_df.head(3).to_string())
            if 'Contract' in contract_price_df.columns:
                print(f"   - Exemple de contracts: {contract_price_df['Contract'].head(5).tolist()}")
            price_cols = [col for col in contract_price_df.columns if 'Price[' in str(col)]
            print(f"   - Colonnes de prix trouvées: {price_cols}")
        print("-" * 80)
        # DEBUG AMELIORE: Afficher toutes les valeurs de filtrage
        if not volume_df.empty:
            for col in volume_df.columns:
                if 'export' in str(col).lower() or 'local' in str(col).lower():
                    print(f"🔍 DEBUG: Colonne {col} contient: {volume_df[col].unique()}")
                    for val in volume_df[col].unique():
                        count = len(volume_df[volume_df[col] == val])
                        print(f"     '{val}': {count} lignes")
        
            
        # Ajouter colonnes de tokens pour le matching par similarité
        if not fulfill_df_export.empty and 'Contract' in fulfill_df_export.columns:
            fulfill_df_export["Contract_tokens"] = fulfill_df_export["Contract"].apply(self.tokenizer_contract)
        if not volume_df_export.empty and 'Contract' in volume_df_export.columns:
            volume_df_export["Contract_tokens"] = volume_df_export["Contract"].apply(self.tokenizer_contract)
        
        print(f"🔍 DEBUG BLOC 1: Utilisation du matching par tokens pour les pays et prix (seuil=6)")
        print("-" * 80)
        # DEBUG AMELIORE: Afficher toutes les valeurs de filtrage
        if not volume_df.empty:
            for col in volume_df.columns:
                if 'export' in str(col).lower() or 'local' in str(col).lower():
                    print(f"🔍 DEBUG: Colonne {col} contient: {volume_df[col].unique()}")
                    for val in volume_df[col].unique():
                        count = len(volume_df[volume_df[col] == val])
                        print(f"     '{val}': {count} lignes")
        
        
        for i, row in volume_df_export.iterrows():
            # Lecture directe de l'entité depuis la table Volume
            site_entite = row.get("Entity", "")
            
            # Récupérer le pays par matching de tokens et prix par mapping direct
            contract = row.get("Contract", "")
            volume_tokens = row.get("Contract_tokens", set())
            
            # Pays via matching par tokens
            pays = ""
            if not fulfill_df_export.empty and volume_tokens:
                pays = self.clean_tuple_string(
                    self.match_par_tokens(volume_tokens, fulfill_df_export, seuil=6)
                )
                
            # Prix depuis ContractPrice (nouvelle logique)
            # Note: les prix seront récupérés dynamiquement pour chaque mois dans la boucle
            
            # DEBUG: Afficher les informations de matching
            if contract and not pays:
                print(f"⚠️ DEBUG BLOC 1 Ligne {i}: Contract '{contract}' - Aucun match trouvé par tokens pour le pays")
            
            print(f"🔍 DEBUG BLOC 1 Ligne {i}: Site/Entité='{site_entite}' (Lecture directe) | Contract='{contract}' | Pays='{pays}' (Via tokens) | Prix=SERA RÉCUPÉRÉ PAR MOIS (ContractPrice)")
            
            # Traitement des volumes : prioriser les colonnes mensuelles puis trimestres
            # Mapping des mois aux numéros
            month_mapping = {
                "Juillet": 7, "juillet": 7,
                "Aout": 8, "août": 8, "aout": 8,
                "Septembre": 9, "septembre": 9,
                "Octobre": 10, "octobre": 10,
                "Novembre": 11, "novembre": 11,
                "Decembre": 12, "décembre": 12, "decembre": 12
            }
            
            volumes_traites = []  # Pour éviter les doublons
            
            # 1. D'abord traiter les colonnes mensuelles directes
            for col in row.index:
                if "Volume[" in str(col) and any(month in str(col) for month in month_mapping.keys()):
                    volume_value = row[col]
                    if pd.notna(volume_value):
                        try:
                            volume_mensuel = float(volume_value)
                            # Identifier le mois
                            mois_num = None
                            for month_name, month_num in month_mapping.items():
                                if month_name in str(col):
                                    mois_num = month_num
                                    break
                            
                            if mois_num:
                                # Récupérer le prix pour ce mois spécifique
                                prix_fob = self.get_prix_from_contract_price(contract, contract_price_df, mois_num)
                                
                                volumes_traites.append(mois_num)
                                print(f"🔍 DEBUG BLOC 1 MENSUEL: Ligne {i}, {col} = {volume_value}, Mois = {mois_num}, Prix = {prix_fob}")
                                all_rows.append({
                                    "Bloc": "Bloc 1",
                                    "Exercice": exercice,
                                    "Date de la Version": str(date_version),
                                    "Année": annee,
                                    "Mois": mois_num,
                                    "Type de transaction": "Vente Export",
                                    "Site/Entité": site_entite,
                                    "Qualité": row.get("Product", ""),
                                    "Partenaire Groupe": "HG",
                                    "Pays": pays,
                                    "Région": row.get("Region", ""),  # NOUVEAU
"Devise": "USD",
                                    "VOLUME (T)": volume_mensuel * 1000,  # Conversion kt → t
                                    "PRIX FOB": prix_fob,
                                    "Prix Fret": "",
                                    "Prix Frais d'approche MP": "",
                                    "TypeProduct": row.get("TypeProduct", ""),
                                    "_ordre_original": i
                                })
                        except (ValueError, TypeError):
                            print(f"⚠️ DEBUG BLOC 1 MENSUEL: Impossible de convertir {col} = {volume_value}")
            
            # 2. Ensuite traiter les trimestres pour les mois non couverts
            for q in ["Q1", "Q2", "Q3", "Q4"]:
                volume_colonne = f"Volume[{q}]"
                if volume_colonne in row:
                    volume_value = row[volume_colonne]
                    if pd.notna(volume_value):
                        try:
                            volume_mensuel = float(volume_value) / 3  # Diviser par 3 mois du trimestre
                            print(f"🔍 DEBUG BLOC 1 TRIMESTRE: Ligne {i}, {volume_colonne} = {volume_value}, Volume mensuel = {volume_mensuel}")
                        except (ValueError, TypeError):
                            print(f"⚠️ DEBUG BLOC 1 TRIMESTRE: Impossible de convertir {volume_colonne} = {volume_value}")
                            continue
                    
                        # Générer lignes pour les mois du trimestre non déjà traités
                        months = self.quarter_to_months[q]
                        for mois in months:
                            if mois not in volumes_traites:  # Éviter les doublons
                                # Récupérer le prix pour ce mois spécifique
                                prix_fob_mois = self.get_prix_from_contract_price(contract, contract_price_df, mois)
                                
                                all_rows.append({
                                    "Bloc": "Bloc 1",
                                    "Exercice": exercice,
                                    "Date de la Version": str(date_version),
                                    "Année": annee,
                                    "Mois": mois,
                                    "Type de transaction": "Vente Export",
                                    "Site/Entité": site_entite,
                                    "Qualité": row.get("Product", ""),
                                    "Partenaire Groupe": "HG",
                                    "Pays": pays,
                                    "Région": row.get("Region", ""),  # NOUVEAU
"Devise": "USD",
                                    "VOLUME (T)": volume_mensuel * 1000,  # Conversion kt → t
                                    "PRIX FOB": prix_fob_mois,
                                    "Prix Fret": "",
                                    "Prix Frais d'approche MP": "",
                                    "TypeProduct": row.get("TypeProduct", ""),
                                    "_ordre_original": i
                                })

        # === BLOC 2 - Vente Local ===

        # CAS 1: Ventes Locales directement depuis Volume + Prix depuis FulfillData
        if "Export_or_Local" in volume_df.columns:
            volume_df_local = volume_df[volume_df["Export_or_Local"] == "LOCAL"].copy()
            print(f"🔍 DEBUG Volume LOCAL: {len(volume_df_local)} lignes après filtre Local")
        elif "Export or local" in volume_df.columns:
            volume_df_local = volume_df[volume_df["Export or local"].str.upper() == "LOCAL"].copy()
            print(f"🔍 DEBUG Volume LOCAL: {len(volume_df_local)} lignes après filtre LOCAL")
        else:
            volume_df_local = pd.DataFrame()  # Si pas de colonne filtre, pas de données locales
            print("⚠️ DEBUG Volume LOCAL: Aucune colonne de filtre Export/Local trouvée")
        
        # Ajouter les tokens pour le volume local
        if not volume_df_local.empty and 'Contract' in volume_df_local.columns:
            volume_df_local["Contract_tokens"] = volume_df_local["Contract"].apply(self.tokenizer_contract)
        
        # Préparer FulfillData Local pour matching par tokens
        if "Export or local" in fulfill_df.columns:
            fulfill_df_local = fulfill_df[fulfill_df["Export or local"] == "LOCAL"].copy()
            # Ajouter colonnes de tokens pour le matching par similarité
            if not fulfill_df_local.empty and 'Contract' in fulfill_df_local.columns:
                fulfill_df_local["Contract_tokens"] = fulfill_df_local["Contract"].apply(self.tokenizer_contract)
        else:
            fulfill_df_local = pd.DataFrame()
        
        print(f"🔍 DEBUG BLOC 1: Utilisation du matching par tokens pour les prix LOCAL (seuil=6)")
        
        max_export_order = len(volume_df_export)
        
        for i, row in volume_df_local.iterrows():
            # Lecture directe de l'entité depuis la table Volume
            site_entite = row.get("Entity", "")
            
            # Récupérer le prix par matching de tokens
            contract = row.get("Contract", "")
            volume_tokens = row.get("Contract_tokens", set())
            
            # Prix depuis ContractPrice (nouvelle logique)
            # Note: les prix seront récupérés dynamiquement pour chaque mois dans la boucle
            
            # DEBUG: Afficher les informations de matching
            print(f"🔍 DEBUG BLOC 1 CAS 1 Ligne {i}: Site/Entité='{site_entite}' (Lecture directe) | Contract='{contract}' | Prix=SERA RÉCUPÉRÉ PAR MOIS (ContractPrice)")
            
            # Traitement des volumes : prioriser les colonnes mensuelles puis trimestres
            # Mapping des mois aux numéros
            month_mapping = {
                "Juillet": 7, "juillet": 7,
                "Aout": 8, "août": 8, "aout": 8,
                "Septembre": 9, "septembre": 9,
                "Octobre": 10, "octobre": 10,
                "Novembre": 11, "novembre": 11,
                "Decembre": 12, "décembre": 12, "decembre": 12
            }
            
            volumes_traites = []  # Pour éviter les doublons
            
            # 1. D'abord traiter les colonnes mensuelles directes
            for col in row.index:
                if "Volume[" in str(col) and any(month in str(col) for month in month_mapping.keys()):
                    volume_value = row[col]
                    if pd.notna(volume_value):
                        try:
                            volume_mensuel = float(volume_value)
                            # Identifier le mois
                            mois_num = None
                            for month_name, month_num in month_mapping.items():
                                if month_name in str(col):
                                    mois_num = month_num
                                    break
                            
                            if mois_num:
                                # Récupérer le prix pour ce mois spécifique
                                prix_fob = self.get_prix_from_contract_price(contract, contract_price_df, mois_num)
                                
                                volumes_traites.append(mois_num)
                                print(f"🔍 DEBUG BLOC 1 MENSUEL: Ligne {i}, {col} = {volume_value}, Mois = {mois_num}, Prix = {prix_fob}")
                                all_rows.append({
                                    "Bloc": "Bloc 1",
                                    "Exercice": exercice,
                                    "Date de la Version": str(date_version),
                                    "Année": annee,
                                    "Mois": mois_num,
                                    "Type de transaction": "Vente Locale",
                                    "Site/Entité": site_entite,
                                    "Qualité": row.get("Product", ""),
                                    "Partenaire Groupe": "HG",
                                    "Pays": "Maroc",
                                    "Région": row.get("Region", ""),  # NOUVEAU
"Devise": "USD",
                                    "VOLUME (T)": volume_mensuel * 1000,  # Conversion kt → t
                                    "PRIX FOB": prix_fob,
                                    "Prix Fret": "",
                                    "Prix Frais d'approche MP": "",
                                    "TypeProduct": row.get("TypeProduct", ""),
                                    "_ordre_original": max_export_order + i
                                })
                        except (ValueError, TypeError):
                            print(f"⚠️ DEBUG BLOC 1 MENSUEL: Impossible de convertir {col} = {volume_value}")
            
            # 2. Ensuite traiter les trimestres pour les mois non couverts
            for q in ["Q1", "Q2", "Q3", "Q4"]:
                volume_colonne = f"Volume[{q}]"
                if volume_colonne in row:
                    volume_value = row[volume_colonne]
                    if pd.notna(volume_value):
                        try:
                            volume_mensuel = float(volume_value) / 3  # Diviser par 3 mois du trimestre
                            print(f"🔍 DEBUG BLOC 1 TRIMESTRE: Ligne {i}, {volume_colonne} = {volume_value}, Volume mensuel = {volume_mensuel}")
                        except (ValueError, TypeError):
                            print(f"⚠️ DEBUG BLOC 1 TRIMESTRE: Impossible de convertir {volume_colonne} = {volume_value}")
                            continue
                    
                        # Générer lignes pour les mois du trimestre non déjà traités
                        months = self.quarter_to_months[q]
                        for mois in months:
                            if mois not in volumes_traites:  # Éviter les doublons
                                # Récupérer le prix pour ce mois spécifique
                                prix_fob_mois = self.get_prix_from_contract_price(contract, contract_price_df, mois)
                                
                                all_rows.append({
                                    "Bloc": "Bloc 1",
                                    "Exercice": exercice,
                                    "Date de la Version": str(date_version),
                                    "Année": annee,
                                    "Mois": mois,
                                    "Type de transaction": "Vente Locale",
                                    "Site/Entité": site_entite,
                                    "Qualité": row.get("Product", ""),
                                    "Partenaire Groupe": "HG",
                                    "Pays": "Maroc",
                                    "Région": row.get("Region", ""),  # NOUVEAU
"Devise": "USD",
                                    "VOLUME (T)": volume_mensuel * 1000,  # Conversion kt → t
                                    "PRIX FOB": prix_fob_mois,
                                    "Prix Fret": "",
                                    "Prix Frais d'approche MP": "",
                                    "TypeProduct": row.get("TypeProduct", ""),
                                    "_ordre_original": max_export_order + i
                                })

        # CAS 2: Flux Internes depuis Connexions
        feuille_connexions = xls.parse("Connexions", header=None)
        feuille_price = xls.parse("Price", header=None)
        flow_df = self.extraire_table_par_nom(feuille_connexions, "Flow")
        session_price_df = self.extraire_table_par_nom(feuille_price, "SessionPrice")
        
        destinations_autorisees = ["MP2", "MP1", "MC", "MAP SOLUBLE", "JPH", "Jorf", "JLN", 
                                  "JFC5", "JFC4", "JFC3", "JFC2", "JFC1", "IMACID", "EMAPHOS", 
                                  "OFAS", "TSP Hub", "PMP", "Unite_Co_cristallisation", "MP34"]
        # TODO TBV avant activation
        # Filtrer les connexions contenant "VenteLocale"
        # flow_df = flow_df[flow_df.apply(
        #     lambda row: 'ventelocale' in str(row.get('Connexion', '')).lower(),
        #     axis=1
        # )]

        session_price_filtered = session_price_df[session_price_df["Destination"].isin(destinations_autorisees)]
        bloc1_count = len(volume_df_export)
        bloc2_cas1_count = len(volume_df_local)
        
        for i, row in flow_df.iterrows():
            destination = row.get("Destination", "")
            # Filtrer selon destinations autorisées et autres critères
            if destination in destinations_autorisees:
                origin = row.get("Origin", "")
                product = row.get("Product", "")
                
                for q in ["Q1", "Q2", "Q3", "Q4"]:
                    colname = f"Flow[{q}]"
                    if colname in row and pd.notna(row[colname]):
                        try:
                            volume_mensuel = float(row[colname]) / 3  # Diviser par 3 mois du trimestre
                        except:
                            volume_mensuel = ""
                        
                        # Récupérer prix depuis SessionPrice avec matching
                        prix_fob = ""
                        prix_colonne = f"Price[{q}]"
                        matching_prices = session_price_filtered[
                            (session_price_filtered["Destination"] == destination) &
                            (session_price_filtered["Origin"] == origin) &
                            (session_price_filtered["Product"] == product)
                        ]
                        
                        if not matching_prices.empty and prix_colonne in matching_prices.columns:
                            prix_value = matching_prices.iloc[0][prix_colonne]
                            if pd.notna(prix_value) and float(prix_value) != 0:
                                try:
                                    prix_fob = float(prix_value)
                                except:
                                    prix_fob = ""
                        
                        # Générer lignes pour les mois du trimestre
                        months = self.quarter_to_months[q]
                        for mois in months:
                            all_rows.append({
                                "Bloc": "Bloc 2",
                                "Exercice": exercice,
                                "Date de la Version": str(date_version),
                                "Année": annee,
                                "Mois": mois,
                                "Type de transaction": "Vente Locale",
                                "Site/Entité": origin,
                                "Qualité": product,
                                "Partenaire Groupe": destination,
                                "Pays": "Maroc",
                                "Région": row.get("Region", ""),  # NOUVEAU
"Devise": "USD",
                                "VOLUME (T)": volume_mensuel * 1000,  # Conversion kt → t
                                "PRIX FOB": prix_fob,
                                "Prix Fret": "",
                                "Prix Frais d'approche MP": "",
                                "_ordre_original": bloc1_count + bloc2_cas1_count + i
                            })

        # === BLOC 3 - Consommation de MP ===
        feuille_raw = xls.parse("RawMaterials", header=None)
        used_df = self.extraire_table_par_nom(feuille_raw, "UsedVolume")
        # TODO must change UsedVolume by Flow table, TBV before activation
        # Source = Connexions #Flow avec filtre "Port dans Origin"
        # flow_df = self.extraire_table_par_nom(feuille_connexions, "Flow")
        # flow_amp = flow_df[flow_df['Origin'].str.contains('Port', case=False, na=False)]

        prices_df = self.extraire_table_par_nom(feuille_raw, "Prices")
        
        # Calculer l'ordre pour le Bloc 3 (après Bloc 1 Export + Bloc 2 Local + Flux)
        bloc2_total_count = len(volume_df_local) + len(flow_df[flow_df["Destination"].isin(destinations_autorisees)])
        max_bloc2_order = len(volume_df_export) + bloc2_total_count
        
        for i, row in used_df.iterrows():
            # Vérifier s'il y a des colonnes trimestrielles Volume[Q1/Q2/Q3/Q4]
            has_quarterly_data = any(f"Volume[{q}]" in row for q in ["Q1", "Q2", "Q3", "Q4"])
            
            if has_quarterly_data:
                # Mapping des mois aux numéros
                month_mapping = {
                    "Juillet": 7, "juillet": 7,
                    "Aout": 8, "août": 8, "aout": 8,
                    "Septembre": 9, "septembre": 9,
                    "Octobre": 10, "octobre": 10,
                    "Novembre": 11, "novembre": 11,
                    "Decembre": 12, "décembre": 12, "decembre": 12
                }
                
                volumes_traites = []  # Pour éviter les doublons
                
                # 1. D'abord traiter les colonnes mensuelles directes
                for col in row.index:
                    if "Volume[" in str(col) and any(month in str(col) for month in month_mapping.keys()):
                        volume_value = row[col]
                        if pd.notna(volume_value):
                            try:
                                volume_mensuel = float(volume_value)
                                # Identifier le mois
                                mois_num = None
                                for month_name, month_num in month_mapping.items():
                                    if month_name in str(col):
                                        mois_num = month_num
                                        break
                                
                                if mois_num:
                                    volumes_traites.append(mois_num)
                                    
                                    # Récupérer prix depuis #Prices
                                    prix_fob = ""
                                    product = row.get("Product", "")
                                    if not prices_df.empty and product:
                                        matching_prices = prices_df[prices_df.get("Product", "") == product]
                                        if not matching_prices.empty:
                                            # Chercher prix mensuel correspondant
                                            month_names = ["Juillet", "Aout", "Septembre", "Q4"]
                                            for mn in month_names:
                                                prix_col = f"Price[{mn}]"
                                                if prix_col in matching_prices.columns:
                                                    prix_value = matching_prices.iloc[0][prix_col]
                                                    if pd.notna(prix_value) and float(prix_value) != 0:
                                                        try:
                                                            prix_fob = float(prix_value)
                                                            break
                                                        except:
                                                            pass
                                    
                                    all_rows.append({
                                        "Bloc": "Bloc 3",
                                        "Exercice": exercice,
                                        "Date de la Version": str(date_version),
                                        "Année": annee,
                                        "Mois": mois_num,
                                        "Type de transaction": "Achat de MP",
                                        "Site/Entité": row.get("Facility", ""),
                                        "Qualité": row.get("Product", ""),
                                        "Partenaire Groupe": "HG",
                                        "Pays": "Maroc",
                                        "Région": row.get("Region", ""),  # NOUVEAU
"Devise": "USD",
                                        "VOLUME (T)": volume_mensuel * 1000,  # Conversion kt → t
                                        "PRIX FOB": prix_fob,
                                        "Prix Fret": "",
                                        "Prix Frais d'approche MP": "",
                                        "_ordre_original": max_bloc2_order + i
                                    })
                            except (ValueError, TypeError):
                                pass
                
                # 2. Logique trimestrielle pour les mois non couverts
                for q in ["Q1", "Q2", "Q3", "Q4"]:
                    colname = f"Volume[{q}]"
                    if colname in row and pd.notna(row[colname]):
                        try:
                            volume_mensuel = float(row[colname]) / 3  # Diviser par 3 mois du trimestre
                        except:
                            volume_mensuel = ""
                        
                        # Récupérer prix depuis #Prices (simple matching par Product si disponible)
                        prix_fob = ""
                        product = row.get("Product", "")
                        if not prices_df.empty and product:
                            prix_colonne = f"Price[{q}]"
                            matching_prices = prices_df[prices_df.get("Product", "") == product]
                            if not matching_prices.empty and prix_colonne in matching_prices.columns:
                                prix_value = matching_prices.iloc[0][prix_colonne]
                                if pd.notna(prix_value) and float(prix_value) != 0:
                                    try:
                                        prix_fob = float(prix_value)
                                    except:
                                        prix_fob = ""
                        
                        # Générer lignes pour les mois du trimestre non déjà traités
                        months = self.quarter_to_months[q]
                        for mois in months:
                            if mois not in volumes_traites:  # Éviter les doublons
                                all_rows.append({
                                    "Bloc": "Bloc 3",
                                    "Exercice": exercice,
                                    "Date de la Version": str(date_version),
                                    "Année": annee,
                                    "Mois": mois,
                                    "Type de transaction": "Achat de MP",
                                    "Site/Entité": row.get("Facility", ""),
                                    "Qualité": row.get("Product", ""),
                                    "Partenaire Groupe": "HG",
                                    "Pays": "Maroc",
                                    "Région": row.get("Region", ""),  # NOUVEAU
"Devise": "USD",
                                    "VOLUME (T)": volume_mensuel * 1000,  # Conversion kt → t
                                    "PRIX FOB": prix_fob,
                                    "Prix Fret": "",
                                    "Prix Frais d'approche MP": "",
                                    "_ordre_original": max_bloc2_order + i
                                })
            else:
                # Logique simple si pas de colonnes trimestrielles (colonne Volume simple)
                volume_simple = row.get("Volume", 0)
                if pd.notna(volume_simple):
                    try:
                        volume_mensuel = float(volume_simple) / 12  # Diviser par 12 mois
                    except:
                        volume_mensuel = ""
                    
                    # Récupérer prix depuis #Prices (simple matching par Product si disponible)
                    prix_fob = ""
                    product = row.get("Product", "")
                    if not prices_df.empty and product:
                        matching_prices = prices_df[prices_df.get("Product", "") == product]
                        if not matching_prices.empty and "Price" in matching_prices.columns:
                            prix_value = matching_prices.iloc[0]["Price"]
                            if pd.notna(prix_value) and float(prix_value) != 0:
                                try:
                                    prix_fob = float(prix_value)
                                except:
                                    prix_fob = ""
                    
                    # Générer 12 lignes (une par mois)
                    for mois in range(1, 13):
                        all_rows.append({
                            "Bloc": "Bloc 3",
                            "Exercice": exercice,
                            "Date de la Version": str(date_version),
                            "Année": annee,
                            "Mois": mois,
                            "Type de transaction": "Achat de MP",
                            "Site/Entité": row.get("Facility", ""),
                            "Qualité": row.get("Product", ""),
                            "Partenaire Groupe": "HG",
                            "Pays": "Maroc",
                            "Région": row.get("Region", ""),  # NOUVEAU
"Devise": "USD",
                            "VOLUME (T)": volume_mensuel * 1000,  # Conversion kt → t
                            "PRIX FOB": prix_fob,
                            "Prix Fret": "",
                            "Prix Frais d'approche MP": "",
                            "_ordre_original": max_bloc2_order + i
                        })

        if not all_rows:
            st.error("Aucune donnée trouvée pour Ventes & MP.")
            return pd.DataFrame(), b""

        # Tri et finalisation selon l'ordre FulfillData et prix croissant
        df_final = pd.DataFrame(all_rows)
        
        # Conserver l'ordre original de FulfillData pour les pays
        pays_order = fulfill_df["Country"].unique().tolist() if "Country" in fulfill_df.columns else []
        if "Maroc" not in pays_order:
            pays_order.insert(0, "Maroc")
        
        pays_order_map = {pays: i for i, pays in enumerate(pays_order)}
        df_final["_ordre_tri"] = df_final["Pays"].map(lambda x: pays_order_map.get(x, 9999))
        df_final["_ordre_export"] = df_final["Type de transaction"].map(
            lambda x: 0 if "export" in str(x).lower() else 1
        )
        
        # Ajout du tri par prix FOB (du plus petit au plus grand)
        def prix_pour_tri(prix_fob):
            """Convertit le prix FOB en valeur numérique pour le tri, 0 si vide"""
            if pd.isna(prix_fob) or prix_fob == "":
                return float('inf')  # Les prix vides à la fin
            try:
                return float(prix_fob)
            except:
                return float('inf')
        
        df_final["_ordre_prix"] = df_final["PRIX FOB"].apply(prix_pour_tri)
        
        # Tri final : BLOC d'abord → ordre FulfillData → Export/Local → Prix FOB croissant → Mois
        df_final = df_final.sort_values([
            "Bloc", "_ordre_tri", "_ordre_export", "_ordre_prix", "Mois"
        ]).reset_index(drop=True)
        
        df_final = df_final.drop(["_ordre_tri", "_ordre_export", "_ordre_original", "_ordre_prix"], axis=1)
        df_final.insert(0, "#ID", [f"#{i+1}" for i in range(len(df_final))])

        # Export Excel avec feuilles séparées pour chaque table source
        print("📊 Génération fichier Excel Ventes & MP...")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # Feuille 1: Fichier plat (avec formatage coloré)
            df_final.to_excel(writer, index=False, sheet_name="Fichier plat Ventes et MP")
            workbook = writer.book
            worksheet = writer.sheets["Fichier plat Ventes et MP"]
            format_bloc1 = workbook.add_format({"bg_color": "#DCE6F1"})
            format_bloc2 = workbook.add_format({"bg_color": "#E2EFDA"})
            format_bloc3 = workbook.add_format({"bg_color": "#FFF2CC"})
            for row_idx, bloc in enumerate(df_final["Bloc"], start=1):
                fmt = format_bloc1 if bloc == "Bloc 1" else format_bloc2 if bloc == "Bloc 2" else format_bloc3
                worksheet.set_row(row_idx, cell_format=fmt)
            
            # Feuilles des tables sources (avec structure originale)
            if not volume_df_export.empty:
                volume_df_export.to_excel(writer, index=False, sheet_name="Volume Export")
                print(f"✅ Table source : Volume Export ({len(volume_df_export)} lignes)")
            
            if not volume_df_local.empty:
                volume_df_local.to_excel(writer, index=False, sheet_name="Volume Local")
                print(f"✅ Table source : Volume Local ({len(volume_df_local)} lignes)")
            
            if not fulfill_df_export.empty:
                fulfill_df_export.to_excel(writer, index=False, sheet_name="FulfillData Export")
                print(f"✅ Table source : FulfillData Export ({len(fulfill_df_export)} lignes)")
            
            if not fulfill_df_local.empty:
                fulfill_df_local.to_excel(writer, index=False, sheet_name="FulfillData Local")
                print(f"✅ Table source : FulfillData Local ({len(fulfill_df_local)} lignes)")
            
            if not volume_df.empty:
                volume_df.to_excel(writer, index=False, sheet_name="Volume")
                print(f"✅ Table source : Volume ({len(volume_df)} lignes)")
            
            if not flow_df.empty:
                flow_df.to_excel(writer, index=False, sheet_name="Flow")
                print(f"✅ Table source : Flow ({len(flow_df)} lignes)")
            
            if not session_price_df.empty:
                session_price_df.to_excel(writer, index=False, sheet_name="SessionPrice")
                print(f"✅ Table source : SessionPrice ({len(session_price_df)} lignes)")
            
            if not used_df.empty:
                used_df.to_excel(writer, index=False, sheet_name="UsedVolume")
                print(f"✅ Table source : UsedVolume ({len(used_df)} lignes)")
            
            if not prices_df.empty:
                prices_df.to_excel(writer, index=False, sheet_name="Prices")
                print(f"✅ Table source : Prices ({len(prices_df)} lignes)")
            
            # 🎨 Coloration des onglets selon le bloc d'appartenance
            tab_color_map = {
                "Volume Export": "#DCE6F1",       # Bloc 1 - Bleu
                "Volume Local": "#DCE6F1",        # Bloc 1 - Bleu
                "FulfillData Export": "#DCE6F1",  # Bloc 1 - Bleu
                "FulfillData Local": "#DCE6F1",   # Bloc 1 - Bleu
                "Volume": "#DCE6F1",             # Bloc 1 - Bleu
                "Flow": "#E2EFDA",               # Bloc 2 - Vert
                "SessionPrice": "#E2EFDA",       # Bloc 2 - Vert
                "UsedVolume": "#FFF2CC",         # Bloc 3 - Jaune
                "Prices": "#FFF2CC",             # Bloc 3 - Jaune
            }
            for sheet_name, color in tab_color_map.items():
                if sheet_name in writer.sheets:
                    writer.sheets[sheet_name].set_tab_color(color)
        
        return df_final, output.getvalue()

    def generate_ppv_ventes_mp_v2(self, uploaded_file, exercice: str, date_version):
        """Génère le DataFrame et l'Excel pour Ventes & MP."""
        xls = pd.ExcelFile(uploaded_file)
        annee = pd.to_datetime(date_version).year
        all_rows = []

        # === BLOC 1 - Vente Export ===
        sheet_sales_name = self.find_sheet(xls, "sales")
        sheet_mps_name = self.find_sheet(xls, "multi", "period", "sales")
        sheet_price_name = self.find_sheet(xls, "price")

        if sheet_sales_name is None or sheet_mps_name is None:
            st.error("Feuilles concernant les ventes non trouvées dans le fichier.")
            return pd.DataFrame(), b""

        if sheet_price_name is None:
            st.error("Feuille Price non trouvée dans le fichier.")
            return pd.DataFrame(), b""

        feuille_sales = xls.parse(sheet_sales_name, header=None)
        feuille_mps = xls.parse(sheet_mps_name, header=None)
        feuille_price = xls.parse(sheet_price_name, header=None)
        volume_df = self.extraire_table_par_nom(feuille_sales, "Volume")
        fulfill_df = self.extraire_table_par_nom(feuille_mps, "FulfillData")
        contract_price_df = self.extraire_table_par_nom(feuille_price, "ContractPrice")

        # Filtrer les données Export pour BLOC 1
        if "Export_or_Local" in volume_df.columns:
            volume_df_export = volume_df[volume_df["Export_or_Local"] == "EXPORT"].copy()
        elif "Export or local" in volume_df.columns:
            volume_df_export = volume_df[volume_df["Export or local"].str.upper() == "EXPORT"].copy()
        else:
            volume_df_export = volume_df.copy()

        # Filtrer FulfillData Export
        if "Export or local" in fulfill_df.columns:
            fulfill_df_export = fulfill_df[fulfill_df["Export or local"] == "EXPORT"].copy()
        else:
            fulfill_df_export = pd.DataFrame()

        # Ajouter colonnes de tokens pour le matching
        if not fulfill_df_export.empty and 'Contract' in fulfill_df_export.columns:
            fulfill_df_export["Contract_tokens"] = fulfill_df_export["Contract"].apply(self.tokenizer_contract)
        if not volume_df_export.empty and 'Contract' in volume_df_export.columns:
            volume_df_export["Contract_tokens"] = volume_df_export["Contract"].apply(self.tokenizer_contract)

        for i, row in volume_df_export.iterrows():
            site_entite = row.get("Entity", "")
            contract = row.get("Contract", "")
            volume_tokens = row.get("Contract_tokens", set())

            # Pays via matching par tokens
            pays = ""
            if not fulfill_df_export.empty and volume_tokens:
                pays = self.clean_tuple_string(
                    self.match_par_tokens(volume_tokens, fulfill_df_export, seuil=6)
                )

            # Mapping mois
            month_mapping = {
                "Juillet": 7, "juillet": 7,
                "Aout": 8, "août": 8, "aout": 8,
                "Septembre": 9, "septembre": 9,
                "Octobre": 10, "octobre": 10,
                "Novembre": 11, "novembre": 11,
                "Decembre": 12, "décembre": 12, "decembre": 12
            }

            volumes_traites = []

            # 1. Traiter colonnes mensuelles directes
            for col in row.index:
                if "Volume[" in str(col) and any(month in str(col) for month in month_mapping.keys()):
                    volume_value = row[col]
                    if pd.notna(volume_value):
                        try:
                            volume_mensuel = float(volume_value) * 1000  # MODIFIÉ: ajout *1000 (KT→T)
                            mois_num = None
                            for month_name, month_num in month_mapping.items():
                                if month_name in str(col):
                                    mois_num = month_num
                                    break

                            if mois_num:
                                prix_fob = self.get_prix_from_contract_price(contract, contract_price_df, mois_num)
                                volumes_traites.append(mois_num)
                                all_rows.append({
                                    "Bloc": "Bloc 1",
                                    "Exercice": exercice,
                                    "Date de la Version": str(date_version),
                                    "Année": annee,
                                    "Mois": mois_num,
                                    "Type de transaction": "Vente Export",
                                    "Site/Entité": site_entite,
                                    "Qualité": row.get("Product", ""),
                                    "Partenaire Groupe": "HG",
                                    "Pays": pays,
                                    "Region": row.get("Region", ""),
                                    "Devise": "USD",
                                    "VOLUME (T)": volume_mensuel,  # Déjà *1000
                                    "PRIX FOB": prix_fob,
                                    "Prix Fret": "",
                                    "Prix Frais d'approche MP": "",
                                    "TypeProduct": row.get("TypeProduct", ""),
                                    "_ordre_original": i
                                })
                        except (ValueError, TypeError):
                            pass

            # 2. Traiter trimestres pour mois non couverts
            for q in ["Q1", "Q2", "Q3", "Q4"]:
                volume_colonne = f"Volume[{q}]"
                if volume_colonne in row:
                    volume_value = row[volume_colonne]
                    if pd.notna(volume_value):
                        try:
                            volume_mensuel = float(volume_value) * 1000 / 3  # MODIFIÉ: ajout *1000 (KT→T)
                        except (ValueError, TypeError):
                            continue

                        months = self.quarter_to_months[q]
                        for mois in months:
                            if mois not in volumes_traites:
                                prix_fob_mois = self.get_prix_from_contract_price(contract, contract_price_df, mois)
                                all_rows.append({
                                    "Bloc": "Bloc 1",
                                    "Exercice": exercice,
                                    "Date de la Version": str(date_version),
                                    "Année": annee,
                                    "Mois": mois,
                                    "Type de transaction": "Vente Export",
                                    "Site/Entité": site_entite,
                                    "Qualité": row.get("Product", ""),
                                    "Partenaire Groupe": "HG",
                                    "Pays": pays,
                                    "Region": row.get("Region", ""),
                                    "Devise": "USD",
                                    "VOLUME (T)": volume_mensuel,  # Déjà *1000
                                    "PRIX FOB": prix_fob_mois,
                                    "Prix Fret": "",
                                    "Prix Frais d'approche MP": "",
                                    "TypeProduct": row.get("TypeProduct", ""),
                                    "_ordre_original": i
                                })

        # === BLOC 2 - Vente Local ===

        # CAS 1: Ventes Locales depuis Volume
        if "Export_or_Local" in volume_df.columns:
            volume_df_local = volume_df[volume_df["Export_or_Local"] == "LOCAL"].copy()
        elif "Export or local" in volume_df.columns:
            volume_df_local = volume_df[volume_df["Export or local"].str.upper() == "LOCAL"].copy()
        else:
            volume_df_local = pd.DataFrame()

        if not volume_df_local.empty and 'Contract' in volume_df_local.columns:
            volume_df_local["Contract_tokens"] = volume_df_local["Contract"].apply(self.tokenizer_contract)

        # if "Export or local" in fulfill_df.columns:
        #     fulfill_df_local = fulfill_df[fulfill_df["Export or local"] == "LOCAL"].copy()
        #     if not fulfill_df_local.empty and 'Contract' in fulfill_df_local.columns:
        #         fulfill_df_local["Contract_tokens"] = fulfill_df_local["Contract"].apply(self.tokenizer_contract)
        # else:
        #     fulfill_df_local = pd.DataFrame()

        max_export_order = len(volume_df_export)

        for i, row in volume_df_local.iterrows():
            site_entite = row.get("Entity", "")
            contract = row.get("Contract", "")
            volume_tokens = row.get("Contract_tokens", set())

            month_mapping = {
                "Juillet": 7, "juillet": 7,
                "Aout": 8, "août": 8, "aout": 8,
                "Septembre": 9, "septembre": 9,
                "Octobre": 10, "octobre": 10,
                "Novembre": 11, "novembre": 11,
                "Decembre": 12, "décembre": 12, "decembre": 12
            }

            volumes_traites = []

            # 1. Traiter colonnes mensuelles directes
            for col in row.index:
                if "Volume[" in str(col) and any(month in str(col) for month in month_mapping.keys()):
                    volume_value = row[col]
                    if pd.notna(volume_value):
                        try:
                            volume_mensuel = float(volume_value) * 1000  # MODIFIÉ: ajout *1000 (KT→T)
                            mois_num = None
                            for month_name, month_num in month_mapping.items():
                                if month_name in str(col):
                                    mois_num = month_num
                                    break

                            if mois_num:
                                prix_fob = self.get_prix_from_contract_price(contract, contract_price_df, mois_num)
                                volumes_traites.append(mois_num)
                                all_rows.append({
                                    "Bloc": "Bloc 2",
                                    "Exercice": exercice,
                                    "Date de la Version": str(date_version),
                                    "Année": annee,
                                    "Mois": mois_num,
                                    "Type de transaction": "Vente Locale",
                                    "Site/Entité": site_entite,
                                    "Qualité": row.get("Product", ""),
                                    "Partenaire Groupe": "HG",
                                    "Pays": "Maroc",
                                    "Devise": "USD",
                                    "VOLUME (T)": volume_mensuel,  # Déjà *1000
                                    "PRIX FOB": prix_fob,
                                    "Prix Fret": "",
                                    "Prix Frais d'approche MP": "",
                                    "TypeProduct": row.get("TypeProduct", ""),
                                    "_ordre_original": max_export_order + i
                                })
                        except (ValueError, TypeError):
                            pass

            # 2. Traiter trimestres
            for q in ["Q1", "Q2", "Q3", "Q4"]:
                volume_colonne = f"Volume[{q}]"
                if volume_colonne in row:
                    volume_value = row[volume_colonne]
                    if pd.notna(volume_value):
                        try:
                            volume_mensuel = float(volume_value) * 1000 / 3  # MODIFIÉ: ajout *1000 (KT→T)
                        except (ValueError, TypeError):
                            continue

                        months = self.quarter_to_months[q]
                        for mois in months:
                            if mois not in volumes_traites:
                                prix_fob_mois = self.get_prix_from_contract_price(contract, contract_price_df, mois)
                                all_rows.append({
                                    "Bloc": "Bloc 2",
                                    "Exercice": exercice,
                                    "Date de la Version": str(date_version),
                                    "Année": annee,
                                    "Mois": mois,
                                    "Type de transaction": "Vente Locale",
                                    "Site/Entité": site_entite,
                                    "Qualité": row.get("Product", ""),
                                    "Partenaire Groupe": "HG",
                                    "Pays": "Maroc",
                                    "Devise": "USD",
                                    "VOLUME (T)": volume_mensuel,  # Déjà *1000
                                    "PRIX FOB": prix_fob_mois,
                                    "Prix Fret": "",
                                    "Prix Frais d'approche MP": "",
                                    "TypeProduct": row.get("TypeProduct", ""),
                                    "_ordre_original": max_export_order + i
                                })

        # CAS 2: Flux Internes depuis Connexions
        feuille_connexions = xls.parse("Connexions", header=None)
        feuille_price = xls.parse("Price", header=None)
        flow_df = self.extraire_table_par_nom(feuille_connexions, "Flow")
        # filter out with VenteLocale Connexion
        # flow_df = flow_df[flow_df['Connexion'].str.contains('VenteLocale', case=False, na=False)]
        # session_price_filtered = session_price_df[session_price_df['Connexion'].str.contains('VenteLocale', case=False, na=False)]

        mots_cles = ['Cession', 'Tolling', 'Achat MP', 'VenteLocale']

        # Créer un pattern regex avec OR (|)
        pattern = '|'.join(mots_cles)

        # Filtrer
        flow_df = flow_df[flow_df['Connexion'].str.contains(pattern, case=False, na=False, regex=True)]

        session_price_df = self.extraire_table_par_nom(feuille_price, "SessionPrice")
        session_price_filtered = session_price_df[session_price_df['Connexion'].str.contains(pattern, case=False, na=False, regex=True)]

        bloc1_count = len(volume_df_export)
        bloc2_cas1_count = len(volume_df_local)

        for i, row in flow_df.iterrows():
            connexion = row.get("Connexion", "")
            origin = row.get("Origin", "")
            destination = row.get("Destination", "")
            product = row.get("Product", "")

            # Mapping mois
            month_mapping = {
                "Juillet": 7, "juillet": 7,
                "Aout": 8, "août": 8, "aout": 8,
                "Septembre": 9, "septembre": 9,
                "Octobre": 10, "octobre": 10,
                "Novembre": 11, "novembre": 11,
                "Decembre": 12, "décembre": 12, "decembre": 12
            }

            volumes_traites = []

            # 1. Colonnes mensuelles
            for col in row.index:
                if "Flow[" in str(col) and any(month in str(col) for month in month_mapping.keys()):
                    volume_value = row[col]
                    if pd.notna(volume_value):
                        try:
                            volume_mensuel = float(volume_value) * 1000
                            mois_num = None
                            for month_name, month_num in month_mapping.items():
                                if month_name in str(col):
                                    mois_num = month_num
                                    break

                            if mois_num:
                                volumes_traites.append(mois_num)

                                # Prix mensuel
                                prix_fob = ""
                                prix_col_mois = f"Price[{[k for k, v in month_mapping.items() if v == mois_num][0]}]"
                                matching_prices = session_price_filtered[
                                    session_price_filtered["Connexion"] == connexion]
                                if not matching_prices.empty and prix_col_mois in matching_prices.columns:
                                    prix_value = matching_prices.iloc[0][prix_col_mois]
                                    if pd.notna(prix_value) and float(prix_value) != 0:
                                        try:
                                            prix_fob = float(prix_value)
                                        except:
                                            pass

                                all_rows.append({
                                    "Bloc": "Bloc 2",
                                    "Exercice": exercice,
                                    "Date de la Version": str(date_version),
                                    "Année": annee,
                                    "Mois": mois_num,
                                    "Type de transaction": "Vente Locale",
                                    "Site/Entité": origin,
                                    "Qualité": product,
                                    "Partenaire Groupe": destination,
                                    "Pays": "Maroc",
                                    "Devise": "USD",
                                    "VOLUME (T)": volume_mensuel,
                                    "PRIX FOB": prix_fob,
                                    "Prix Fret": "",
                                    "Prix Frais d'approche MP": "",
                                    "_ordre_original": bloc1_count + bloc2_cas1_count + i
                                })
                        except (ValueError, TypeError):
                            pass

            # 2. Trimestres (pour mois non traités)
            for q in ["Q1", "Q2", "Q3", "Q4"]:
                colname = f"Flow[{q}]"
                if colname in row and pd.notna(row[colname]):
                    try:
                        volume_mensuel = float(row[colname]) * 1000 / 3
                    except:
                        continue

                    prix_fob = ""
                    prix_colonne = f"Price[{q}]"
                    matching_prices = session_price_filtered[session_price_filtered["Connexion"] == connexion]
                    if not matching_prices.empty and prix_colonne in matching_prices.columns:
                        prix_value = matching_prices.iloc[0][prix_colonne]
                        if pd.notna(prix_value) and float(prix_value) != 0:
                            try:
                                prix_fob = float(prix_value)
                            except:
                                pass

                    months = self.quarter_to_months[q]
                    for mois in months:
                        if mois not in volumes_traites:
                            all_rows.append({
                                "Bloc": "Bloc 2",
                                "Exercice": exercice,
                                "Date de la Version": str(date_version),
                                "Année": annee,
                                "Mois": mois,
                                "Type de transaction": "Vente Locale",
                                "Site/Entité": origin,
                                "Qualité": product,
                                "Partenaire Groupe": destination,
                                "Pays": "Maroc",
                                "Devise": "USD",
                                "VOLUME (T)": volume_mensuel,
                                "PRIX FOB": prix_fob,
                                "Prix Fret": "",
                                "Prix Frais d'approche MP": "",
                                "_ordre_original": bloc1_count + bloc2_cas1_count + i
                            })

        # === BLOC 3 - Consommation de MP ===
        # MODIFIÉ: Changement de source RawMaterials#UsedVolume → Connexions#Flow
        feuille_connexions = xls.parse("Connexions", header=None)
        feuille_raw = xls.parse("RawMaterials", header=None)  # Pour Prices uniquement
        flow_df_amp = self.extraire_table_par_nom(feuille_connexions, "Flow")  # MODIFIÉ: source Flow
        prices_df = self.extraire_table_par_nom(feuille_raw, "Prices")

        # MODIFIÉ: Filtre Port dans Origin
        flow_amp = flow_df_amp[flow_df_amp['Origin'].str.contains('Port', case=False, na=False)]

        # bloc2_total_count = len(volume_df_local) + len(flow_df[flow_df["Destination"].isin(destinations_autorisees)])
        bloc2_total_count = len(volume_df_local) + len(flow_df)
        max_bloc2_order = len(volume_df_export) + bloc2_total_count

        for i, row in flow_amp.iterrows():  # MODIFIÉ: flow_amp au lieu de used_df
            # MODIFIÉ: Détection colonnes Flow[Q1/Q2/Q3/Q4] au lieu de Volume[...]
            has_quarterly_or_monthly_data = any(f"Flow[{q}]" in row for q in ["Q1", "Q2", "Q3", "Q4"]) or any(f'Flow[{month.capitalize()}]' in row for month in self.month_name_to_num.keys())

            if has_quarterly_or_monthly_data:
                month_mapping = {
                    "Juillet": 7, "juillet": 7,
                    "Aout": 8, "août": 8, "aout": 8,
                    "Septembre": 9, "septembre": 9,
                    "Octobre": 10, "octobre": 10,
                    "Novembre": 11, "novembre": 11,
                    "Decembre": 12, "décembre": 12, "decembre": 12
                }

                volumes_traites = []

                # 1. Traiter colonnes mensuelles
                for col in row.index:
                    # MODIFIÉ: Flow[ au lieu de Volume[
                    if "Flow[" in str(col) and any(month in str(col) for month in month_mapping.keys()):
                        volume_value = row[col]
                        if pd.notna(volume_value):
                            try:
                                volume_mensuel = float(volume_value) * 1000  # MODIFIÉ: ajout *1000 (KT→T)
                                mois_num = None
                                for month_name, month_num in month_mapping.items():
                                    if month_name in str(col):
                                        mois_num = month_num
                                        break

                                if mois_num:
                                    volumes_traites.append(mois_num)

                                    # Prix depuis #Prices
                                    prix_fob = ""
                                    product = row.get("Product", "")
                                    if not prices_df.empty and product:
                                        matching_prices = prices_df[prices_df.get("Product", "") == product]
                                        if not matching_prices.empty:
                                            month_names = ["Juillet", "Aout", "Septembre", "Q4"]
                                            for mn in month_names:
                                                prix_col = f"Price[{mn}]"
                                                if prix_col in matching_prices.columns:
                                                    prix_value = matching_prices.iloc[0][prix_col]
                                                    if pd.notna(prix_value) and float(prix_value) != 0:
                                                        try:
                                                            prix_fob = float(prix_value)
                                                            break
                                                        except:
                                                            pass

                                    all_rows.append({
                                        "Bloc": "Bloc 3",
                                        "Exercice": exercice,
                                        "Date de la Version": str(date_version),
                                        "Année": annee,
                                        "Mois": mois_num,
                                        "Type de transaction": "Achat de MP",
                                        "Site/Entité": row.get("Destination", ""),
                                        # MODIFIÉ: Destination au lieu de Facility
                                        "Qualité": row.get("Product", ""),
                                        "Partenaire Groupe": "HG",
                                        "Pays": "Maroc",
                                        "Devise": "USD",
                                        "VOLUME (T)": volume_mensuel,  # Déjà *1000
                                        "PRIX FOB": prix_fob,
                                        "Prix Fret": "",
                                        "Prix Frais d'approche MP": "",
                                        "_ordre_original": max_bloc2_order + i
                                    })
                            except (ValueError, TypeError):
                                pass

                # 2. Logique trimestrielle
                for q in ["Q1", "Q2", "Q3", "Q4"]:
                    colname = f"Flow[{q}]"  # MODIFIÉ: Flow[ au lieu de Volume[
                    if colname in row and pd.notna(row[colname]):
                        try:
                            volume_mensuel = float(row[colname]) * 1000 / 3  # MODIFIÉ: ajout *1000 (KT→T)
                        except:
                            continue

                        # Prix depuis #Prices
                        prix_fob = ""
                        product = row.get("Product", "")
                        if not prices_df.empty and product:
                            prix_colonne = f"Price[{q}]"
                            matching_prices = prices_df[prices_df.get("Product", "") == product]
                            if not matching_prices.empty and prix_colonne in matching_prices.columns:
                                prix_value = matching_prices.iloc[0][prix_colonne]
                                if pd.notna(prix_value) and float(prix_value) != 0:
                                    try:
                                        prix_fob = float(prix_value)
                                    except:
                                        prix_fob = ""

                        months = self.quarter_to_months[q]
                        for mois in months:
                            if mois not in volumes_traites:
                                all_rows.append({
                                    "Bloc": "Bloc 3",
                                    "Exercice": exercice,
                                    "Date de la Version": str(date_version),
                                    "Année": annee,
                                    "Mois": mois,
                                    "Type de transaction": "Achat de MP",
                                    "Site/Entité": row.get("Destination", ""),
                                    # MODIFIÉ: Destination au lieu de Facility
                                    "Qualité": row.get("Product", ""),
                                    "Partenaire Groupe": "HG",
                                    "Pays": "Maroc",
                                    "Devise": "USD",
                                    "VOLUME (T)": volume_mensuel,  # Déjà *1000
                                    "PRIX FOB": prix_fob,
                                    "Prix Fret": "",
                                    "Prix Frais d'approche MP": "",
                                    "_ordre_original": max_bloc2_order + i
                                })
            else:
                # Logique simple si pas de colonnes trimestrielles
                volume_simple = row.get("Flow", 0)  # MODIFIÉ: Flow au lieu de Volume
                if pd.notna(volume_simple):
                    try:
                        volume_mensuel = float(volume_simple) * 1000 / 12  # MODIFIÉ: ajout *1000 (KT→T)
                    except:
                        volume_mensuel = ""

                    # Prix depuis #Prices
                    prix_fob = ""
                    product = row.get("Product", "")
                    if not prices_df.empty and product:
                        matching_prices = prices_df[prices_df.get("Product", "") == product]
                        if not matching_prices.empty and "Price" in matching_prices.columns:
                            prix_value = matching_prices.iloc[0]["Price"]
                            if pd.notna(prix_value) and float(prix_value) != 0:
                                try:
                                    prix_fob = float(prix_value)
                                except:
                                    prix_fob = ""

                    for mois in range(1, 13):
                        all_rows.append({
                            "Bloc": "Bloc 3",
                            "Exercice": exercice,
                            "Date de la Version": str(date_version),
                            "Année": annee,
                            "Mois": mois,
                            "Type de transaction": "Achat de MP",
                            "Site/Entité": row.get("Destination", ""),  # MODIFIÉ: Destination au lieu de Facility
                            "Qualité": row.get("Product", ""),
                            "Partenaire Groupe": "HG",
                            "Pays": "Maroc",
                            "Devise": "USD",
                            "VOLUME (T)": volume_mensuel,  # Déjà *1000
                            "PRIX FOB": prix_fob,
                            "Prix Fret": "",
                            "Prix Frais d'approche MP": "",
                            "_ordre_original": max_bloc2_order + i
                        })

        if not all_rows:
            st.error("Aucune donnée trouvée pour Ventes & MP.")
            return pd.DataFrame(), b""

        # Tri et finalisation
        df_final = pd.DataFrame(all_rows)

        pays_order = fulfill_df["Country"].unique().tolist() if "Country" in fulfill_df.columns else []
        if "Maroc" not in pays_order:
            pays_order.insert(0, "Maroc")

        pays_order_map = {pays: i for i, pays in enumerate(pays_order)}
        df_final["_ordre_tri"] = df_final["Pays"].map(lambda x: pays_order_map.get(x, 9999))
        df_final["_ordre_export"] = df_final["Type de transaction"].map(
            lambda x: 0 if "export" in str(x).lower() else 1
        )

        def prix_pour_tri(prix_fob):
            if pd.isna(prix_fob) or prix_fob == "":
                return float('inf')
            try:
                return float(prix_fob)
            except:
                return float('inf')

        df_final["_ordre_prix"] = df_final["PRIX FOB"].apply(prix_pour_tri)

        df_final = df_final.sort_values([
            "Bloc", "_ordre_tri", "_ordre_export", "_ordre_prix", "Mois"
        ]).reset_index(drop=True)

        df_final = df_final.drop(["_ordre_tri", "_ordre_export", "_ordre_original", "_ordre_prix"], axis=1)
        df_final.insert(0, "#ID", [f"#{i + 1}" for i in range(len(df_final))])

        # Export Excel avec tables sources
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Fichier plat Ventes et MP")
            workbook = writer.book
            worksheet = writer.sheets["Fichier plat Ventes et MP"]
            format_bloc1 = workbook.add_format({"bg_color": "#DCE6F1"})
            format_bloc2 = workbook.add_format({"bg_color": "#E2EFDA"})
            format_bloc3 = workbook.add_format({"bg_color": "#FFF2CC"})
            for row_idx, bloc in enumerate(df_final["Bloc"], start=1):
                fmt = format_bloc1 if bloc == "Bloc 1" else format_bloc2 if bloc == "Bloc 2" else format_bloc3
                worksheet.set_row(row_idx, cell_format=fmt)

            # Tables sources
            if not volume_df_export.empty:
                volume_df_export.to_excel(writer, index=False, sheet_name="Volume Export")
            if not volume_df_local.empty:
                volume_df_local.to_excel(writer, index=False, sheet_name="Volume Local")
            if not fulfill_df_export.empty:
                fulfill_df_export.to_excel(writer, index=False, sheet_name="FulfillData Export")
            # if not fulfill_df_local.empty:
            #     fulfill_df_local.to_excel(writer, index=False, sheet_name="FulfillData Local")
            if not volume_df.empty:
                volume_df.to_excel(writer, index=False, sheet_name="Volume")
            if not flow_df.empty:
                flow_df.to_excel(writer, index=False, sheet_name="Flow")
            if not session_price_df.empty:
                session_price_df.to_excel(writer, index=False, sheet_name="SessionPrice")
            if not flow_amp.empty:  # MODIFIÉ: Ajout table Flow AMP filtrée
                flow_amp.to_excel(writer, index=False, sheet_name="Flow AMP (Port)")
            if not prices_df.empty:
                prices_df.to_excel(writer, index=False, sheet_name="Prices")

            # Coloration onglets
            tab_color_map = {
                "Volume Export": "#DCE6F1",
                "Volume Local": "#DCE6F1",
                "FulfillData Export": "#DCE6F1",
                "FulfillData Local": "#DCE6F1",
                "Volume": "#DCE6F1",
                "Flow": "#E2EFDA",
                "SessionPrice": "#E2EFDA",
                "Flow AMP (Port)": "#FFF2CC",  # MODIFIÉ: Nouvel onglet
                "Prices": "#FFF2CC",
            }
            for sheet_name, color in tab_color_map.items():
                if sheet_name in writer.sheets:
                    writer.sheets[sheet_name].set_tab_color(color)

        return df_final, output.getvalue()

    # ---------------------------------------------------------------------------
    # Module 3: Ratios Matières (OIJ / OIS)
    # ---------------------------------------------------------------------------
    
    def detect_blocks(self, sheet: pd.DataFrame):
        """Découpe la feuille en blocs situés sous chaque entête "Chemical Facility".

        Retourne une liste de DataFrames avec leurs colonnes correctement nommées.
        """
        blocs = []
        i = 0
        max_i = len(sheet)
        while i < max_i:
            cell = str(sheet.iat[i, 0])
            if cell.lower().startswith("chemical facility"):
                # La ligne suivante contient les en-têtes
                if i + 1 >= max_i:
                    break
                headers = sheet.iloc[i + 1].tolist()
                # Rendre les en-têtes uniques et explicites
                def make_unique(col_list):
                    counts = {}
                    result = []
                    for c in col_list:
                        c = str(c).strip()
                        if c == "" or c.lower().startswith("nan"):
                            c = "Unnamed"
                        if c in counts:
                            counts[c] += 1
                            result.append(f"{c}_{counts[c]}")
                        else:
                            counts[c] = 0
                            result.append(c)
                    return result

                headers = make_unique(headers)

                j = i + 2
                # Chercher la fin du bloc (avant prochain "Chemical Facility" ou ligne vide)
                while j < max_i:
                    next_cell = str(sheet.iat[j, 0])
                    if next_cell.lower().startswith("chemical facility"):
                        break
                    # stop when all NaN row
                    if sheet.iloc[j].isna().all():
                        break
                    j += 1
                # Extraire les lignes de données
                data_block = sheet.iloc[i + 2 : j].copy()
                data_block.columns = headers
                data_block = data_block.dropna(how="all")
                blocs.append(data_block)
                i = j  # continuer après ce bloc
            else:
                i += 1
        return blocs



    def parse_treatment_sheet(self, xls: pd.ExcelFile, sheet_name: str, bloc_name: str):
        """Lit la feuille et renvoie un DataFrame au format long avec bloc renseigné."""
        raw_sheet = xls.parse(sheet_name, header=None)
        blocks = self.detect_blocks(raw_sheet)
        if not blocks:
            st.warning(f"Aucun bloc 'Chemical Facility' détecté dans {sheet_name}")
            return pd.DataFrame()

        df = pd.concat(blocks, ignore_index=True)

        # Normaliser les noms de colonnes attendus
        rename_map = {c: c.strip() for c in df.columns}
        df.rename(columns=rename_map, inplace=True)

        # Colonnes fixes à conserver si elles existent
        fixed_cols = []
        for col in ["Budget Year", "Facility", "Treatment", "Flow", "Output"]:
            if col in df.columns:
                fixed_cols.append(col)
        
        # Colonnes ratio = seulement les colonnes avec des valeurs numériques
        # ET qui représentent de vrais types de consommation (pas des processus)
        ratio_cols = []
        for col in df.columns:
            if col not in fixed_cols:
                # Vérifier si la colonne contient des valeurs numériques
                sample_values = df[col].dropna()
                if len(sample_values) > 0:
                    # Tester si au moins une valeur est numérique
                    numeric_count = sum(1 for val in sample_values if pd.api.types.is_numeric_dtype(type(val)) or (isinstance(val, (int, float)) and not pd.isna(val)))
                    if numeric_count > 0:
                        # NOUVEAU: Filtrer pour garder seulement les vrais types de consommation
                        if self.is_valid_consumption_type(col):
                            ratio_cols.append(col)
        
        if not ratio_cols:
            st.error(f"Aucune colonne de ratio numérique identifiée dans {sheet_name}")
            return pd.DataFrame()

        df_long = df.melt(id_vars=fixed_cols, value_vars=ratio_cols,
                          var_name="Type Consommation Spécifique",
                          value_name="Ratio")
        # Nettoyage
        df_long.dropna(subset=["Ratio"], inplace=True)
        df_long = df_long[df_long["Ratio"] != 0]

        # Ajout du bloc d'origine pour coloration éventuelle
        df_long.insert(0, "Bloc", bloc_name)
        return df_long.reset_index(drop=True)

    def generate_ratios(self, ratios_file, summary_file, exercice: str, date_version):
        # ------------------------------------------------------------------
        # 1) Lecture du fichier Ratios (PPV – onglets OIJ / OIS)
        # ------------------------------------------------------------------
        xls = pd.ExcelFile(ratios_file)
        blocks_needed = {"OIJ-Treatment_Matrix": "OIJ", "OIS-Treatment_Matrix": "OIS"}
        all_rows = []
        for sheet, bloc in blocks_needed.items():
            if sheet not in xls.sheet_names:
                continue
            df = self.parse_treatment_sheet(xls, sheet, bloc)
            if len(df):
                all_rows.append(df)
        
        if not all_rows:
            st.error("Aucune donnée trouvée pour Ratios Matières.")
            return pd.DataFrame(), b""
        
        df_total = pd.concat(all_rows, ignore_index=True)
        if "Budget Year" in df_total.columns:
            annee_series = df_total["Budget Year"]
        else:
            annee_series = pd.Series([pd.to_datetime(date_version).year] * len(df_total))
        
        # ------------------------------------------------------------------
        # 2) Récupération des volumes produits par Produit
        #    Désormais depuis le fichier plat PPV Production (Traitements Chimiques)
        # ------------------------------------------------------------------
        from collections import defaultdict
        import os
        volume_par_produit = defaultdict(float)

        ppv_flat_file = "fichier_ppv_production.xlsx"
        if os.path.exists(ppv_flat_file):
            try:
                xls_ppv = pd.ExcelFile(ppv_flat_file)
                sheet_ppv = next((s for s in xls_ppv.sheet_names if "ppv production" in s.lower()), xls_ppv.sheet_names[0])
                df_ppv = pd.read_excel(ppv_flat_file, sheet_name=sheet_ppv, dtype=str, keep_default_na=False)
                # Filtrer Traitements Chimiques
                df_ppv = df_ppv[df_ppv.get("Opération", "") == "Traitements Chimiques"].copy()
                if not df_ppv.empty and "VOLUME (T)" in df_ppv.columns:
                    # Nettoyer / convertir en numérique
                    df_ppv["VOLUME (T)"] = pd.to_numeric(df_ppv["VOLUME (T)"], errors="coerce").fillna(0)
                    grouped = df_ppv.groupby("Qualité")["VOLUME (T)"].sum()
                    for prod, vol in grouped.items():
                        volume_par_produit[str(prod).strip().upper()] += vol
            except Exception as e:
                print(f"⚠️ Erreur lecture PPV Production pour volumes : {e}")
        else:
            print(f"⚠️ Fichier PPV Production non trouvé ({ppv_flat_file}), Volume Produit (T) sera à 0")

        # ---- Ajouter les volumes VENTES (Bloc 1, TypeProduct ciblés) ----------------
        ventes_flat_file = "fichier_ventes_mp.xlsx"
        allowed_typeprod = {"Fertilizers", "Fertilizers W", "MarketableAcids"}
        if os.path.exists(ventes_flat_file):
            try:
                xls_ventes = pd.ExcelFile(ventes_flat_file)
                sheet_vplat = next((s for s in xls_ventes.sheet_names if "fichier plat ventes" in s.lower()), xls_ventes.sheet_names[0])
                df_vplat = pd.read_excel(ventes_flat_file, sheet_name=sheet_vplat, dtype=str, keep_default_na=False)
                # Conserver uniquement Bloc 1 (source #Volume) et TypeProduct voulu
                df_vplat = df_vplat[(df_vplat.get("Bloc", "") == "Bloc 1") & (df_vplat.get("TypeProduct", "").isin(allowed_typeprod))].copy()
                if not df_vplat.empty and "VOLUME (T)" in df_vplat.columns:
                    df_vplat["VOLUME (T)"] = pd.to_numeric(df_vplat["VOLUME (T)"], errors="coerce").fillna(0)
                    grouped_v = df_vplat.groupby("Qualité")["VOLUME (T)"].sum()
                    for prod, vol in grouped_v.items():
                        volume_par_produit[str(prod).strip().upper()] += vol
            except Exception as e:
                print(f"⚠️ Erreur lecture volumes Ventes & MP : {e}")

        # ------------------------------------------------------------------
        # 2bis) Fallback legacy : si volume_par_produit est vide, on garde l ancienne logique SUMMARY
        # ------------------------------------------------------------------
        if not volume_par_produit and summary_file is not None:
            try:
                xls_sum = pd.ExcelFile(summary_file)
                physical_sheet = next((s for s in xls_sum.sheet_names if "physical" in s.lower() or "physique" in s.lower()), None)
                if physical_sheet:
                    physical_df = self.extraire_table_par_nom(
                        xls_sum.parse(physical_sheet, header=None), "VolumeOutputProduct"
                    )
                    for _, row in physical_df.iterrows():
                        prod = str(row.get("Output", "")).strip().upper()
                        for q in ["Q1", "Q2", "Q3", "Q4"]:
                            colname = f"VolumeOutputProduct[{q}]"
                            if colname in row and pd.notna(row[colname]):
                                try:
                                    volume_par_produit[prod] += float(row[colname])
                                except Exception:
                                    pass
            except Exception as e:
                print(f"⚠️ Erreur fallback SUMMARY pour volumes : {e}")

        # Helper normalisation produit
        def _norm_prod(val):
            return str(val).strip().upper() if pd.notna(val) else ""

        # Attribution du Volume Produit (T) basé sur le produit uniquement
        df_total["Volume Produit (T)"] = df_total["Output"].apply(lambda p: volume_par_produit.get(_norm_prod(p), 0))

        # Recalcul Ratio Pondéré après mise à jour du volume
        df_total["Ratio Pondéré"] = df_total["Ratio"] * df_total["Volume Produit (T)"]

        # Placeholder pour Ratio Moyen Pondéré (sera vide pour les lignes entités)
        df_total["Ratio Moyen Pondéré"] = ""

        # ------------------------------------------------------------------
        # 3) Calcul des ratios moyens pondérés GLOBAL par produit / type conso
        # ------------------------------------------------------------------
        aggregated_rows = []
        group_cols = ["Bloc", "Output", "Type Consommation Spécifique"]
        grouped = df_total.groupby(group_cols)

        ratio_moyen_map = {}
        for (bloc, prod, type_conso), grp in grouped:
            vol_sum = grp["Volume Produit (T)"].sum()
            if vol_sum == 0:
                continue
            ratio_moyen_map[(bloc, prod, type_conso)] = grp["Ratio Pondéré"].sum() / vol_sum

        # Remplir la colonne Ratio Moyen Pondéré pour chaque ligne entité
        df_total["Ratio Moyen Pondéré"] = df_total.apply(
            lambda r: ratio_moyen_map.get((r["Bloc"], r["Output"], r["Type Consommation Spécifique"]), ""), axis=1
        )

        # ------------------------------------------------------------------
        # 4) Assemblage final (uniquement lignes entités) et attribution #ID
        # ------------------------------------------------------------------
        df_all = df_total.copy()

        if "Budget Year" in df_total.columns:
            annee_series = df_all.get("Budget Year", pd.Series([pd.to_datetime(date_version).year] * len(df_all)))
        else:
            annee_series = pd.Series([pd.to_datetime(date_version).year] * len(df_all))

        # Créer la table finale en respectant l'ordre des colonnes demandé
        df_final = pd.DataFrame({
            "#ID": [f"#{i+1}" for i in range(len(df_all))],
            "Bloc": df_all["Bloc"],
            "Exercice": exercice,
            "Date de la Version": str(date_version),
            "Année": annee_series,
            "Site/Entité": (
                df_all["Facility"].fillna(df_all["Site/Entité"]) if "Site/Entité" in df_all.columns else df_all["Facility"]
            ) if "Facility" in df_all.columns else (
                df_all["Site/Entité"] if "Site/Entité" in df_all.columns else ""
            ),
            "Qualité": df_all["Output"].where(df_all["Output"].notna(), df_all["Qualité"] if "Qualité" in df_all.columns else ""),
            "Type Consommation Spécifique": df_all["Type Consommation Spécifique"],
            "Ratio": df_all["Ratio"],
            "Volume Produit (T)": df_all["Volume Produit (T)"],
            "Ratio Pondéré": df_all["Ratio Pondéré"],
            "Ratio Moyen Pondéré": df_all["Ratio Moyen Pondéré"],
        })
        # Remove possible duplicates created during aggregation/concat
        df_final = df_final.drop_duplicates(subset=["Site/Entité", "Qualité", "Type Consommation Spécifique"]).reset_index(drop=True)
        
        # Export Excel avec feuilles séparées pour chaque table source
        print("📊 Génération fichier Excel Ratios Matières...")
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # Feuille 1: Fichier plat (avec formatage coloré)
            df_final.to_excel(writer, index=False, sheet_name="Fichier Plat Ratios")
            workbook = writer.book
            worksheet = writer.sheets["Fichier Plat Ratios"]
            format_oij = workbook.add_format({"bg_color": "#DCE6F1"})
            format_ois = workbook.add_format({"bg_color": "#E2EFDA"})
            for row_idx, bloc in enumerate(df_final["Bloc"], start=1):
                worksheet.set_row(row_idx, cell_format=format_oij if bloc == "OIJ" else format_ois)
            
            # Feuilles des tables sources (avec structure originale)
            for sheet, bloc in blocks_needed.items():
                if sheet in xls.sheet_names:
                    # Utiliser les données brutes pour préserver la structure originale
                    raw_sheet_data = xls.parse(sheet, header=None)
                    if not raw_sheet_data.empty:
                        # Créer un nom de feuille propre
                        sheet_name = f"{bloc} Treatment Matrix"
                        raw_sheet_data.to_excel(writer, index=False, sheet_name=sheet_name, header=False)
                        print(f"✅ Table source : {sheet_name} ({len(raw_sheet_data)} lignes)")
            
            # NEW: Coloration des onglets selon le bloc (OIJ / OIS)
            ratio_tab_colors = {
                "OIJ Treatment Matrix": "#DCE6F1",  # Bleu clair
                "OIS Treatment Matrix": "#E2EFDA",  # Vert clair
            }
            for sheet_name, color in ratio_tab_colors.items():
                if sheet_name in writer.sheets:
                    writer.sheets[sheet_name].set_tab_color(color)
        
        return df_final, output.getvalue()

    def generate_ratios_v2(self, ratios_file, summary_file, exercice: str, date_version):
        """Génère les ratios matières avec moyenne simple + volumes."""

        # 1) Lecture fichier Ratios PPV
        xls = pd.ExcelFile(ratios_file)
        blocks_needed = {"OIJ-Treatment_Matrix": "OIJ", "OIS-Treatment_Matrix": "OIS"}
        all_rows = []

        for sheet, bloc in blocks_needed.items():
            if sheet not in xls.sheet_names:
                continue
            df = self.parse_treatment_sheet(xls, sheet, bloc)
            if len(df):
                all_rows.append(df)

        if not all_rows:
            st.error("Aucune donnée trouvée pour Ratios Matières.")
            return pd.DataFrame(), b""

        df_total = pd.concat(all_rows, ignore_index=True)

        def _norm_prod(val):
            return str(val).strip().upper() if pd.notna(val) else ""

        # Attribution volumes
        # df_total["Volume Produit (T)"] = df_total["Output"].apply(lambda p: volume_par_produit.get(_norm_prod(p), 0))

        # 3) Calcul ratio moyen simple (NON pondéré)
        # TODO TBC if we should remove the Bloc (OIJ/OIS) or not
        group_cols = ["Bloc", "Output", "Type Consommation Spécifique"]
        # group_cols = ["Output", "Type Consommation Spécifique"]
        ratio_moyen_map = df_total.groupby(group_cols)["Ratio"].mean().to_dict()

        df_total["Ratio Moyen"] = df_total.apply(
            lambda r: ratio_moyen_map.get((r["Bloc"], r["Output"], r["Type Consommation Spécifique"]), r["Ratio"]),
            axis=1
        )

        # 4) Assemblage final
        annee_series = df_total.get("Budget Year", pd.Series([pd.to_datetime(date_version).year] * len(df_total)))

        df_final = pd.DataFrame({
            "#ID": [f"#{i + 1}" for i in range(len(df_total))],
            "Bloc": df_total["Bloc"],
            "Exercice": exercice,
            "Date de la Version": str(date_version),
            "Année": annee_series,
            "Site/Entité": df_total.get("Facility", df_total.get("Site/Entité", "")),
            'Treatment': df_total["Treatment"],
            'Flow': df_total["Flow"],
            "Qualité": df_total["Output"],
            "Type Consommation Spécifique": df_total["Type Consommation Spécifique"],
            "Ratio": df_total["Ratio"],
            "Ratio Moyen": df_total["Ratio Moyen"]
        })

        df_final = df_final.drop_duplicates(
            subset=["Site/Entité", 'Treatment', 'Flow', "Qualité", "Type Consommation Spécifique"]
        ).reset_index(drop=True)

        # Créer la version sans "Ratio Moyen"
        df_final_simple = df_final.drop(columns=["Ratio Moyen"])

        # 5) Export Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_final_simple.to_excel(writer, index=False, sheet_name="Fichier Plat Ratios simples")
            df_final.to_excel(writer, index=False, sheet_name="Fichier Plat Ratios")
            workbook = writer.book
            worksheet = writer.sheets["Fichier Plat Ratios"]
            format_oij = workbook.add_format({"bg_color": "#DCE6F1"})
            format_ois = workbook.add_format({"bg_color": "#E2EFDA"})

            for row_idx, bloc in enumerate(df_final["Bloc"], start=1):
                worksheet.set_row(row_idx, cell_format=format_oij if bloc == "OIJ" else format_ois)

            for sheet, bloc in blocks_needed.items():
                if sheet in xls.sheet_names:
                    raw_sheet_data = xls.parse(sheet, header=None)
                    if not raw_sheet_data.empty:
                        sheet_name = f"{bloc} Treatment Matrix"
                        raw_sheet_data.to_excel(writer, index=False, sheet_name=sheet_name, header=False)

            ratio_tab_colors = {
                "OIJ Treatment Matrix": "#DCE6F1",
                "OIS Treatment Matrix": "#E2EFDA",
            }
            for sheet_name, color in ratio_tab_colors.items():
                if sheet_name in writer.sheets:
                    writer.sheets[sheet_name].set_tab_color(color)

        return df_final, output.getvalue()

    # ---------------------------------------------------------------------------
    # Interface principale avec onglets
    # ---------------------------------------------------------------------------
    
    def run(self):
        """Lance l'interface principale de l'application."""
        st.set_page_config(
            page_title="Générateur Anaplan",
            page_icon="⚡",
            layout="wide",
            initial_sidebar_state="collapsed"
        )
        
        # CSS personnalisé pour améliorer l'apparence
        st.markdown("""
        <style>
        /* Import Font Awesome */
        @import url('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css');
        
        /* Variables CSS */
        :root {
            --primary-color: #2e7d32;
            --secondary-color: #4caf50;
            --success-color: #2ca02c;
            --warning-color: #ff9800;
            --error-color: #d62728;
            --background-light: #f8f9fa;
            --text-dark: #2c3e50;
            --border-color: #dee2e6;
        }
        
        /* Header principal */
        .main-header {
            background: linear-gradient(135deg, #1b5e20, #2e7d32, #4caf50);
            color: white;
            padding: 2rem;
            border-radius: 10px;
            margin-bottom: 2rem;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            text-align: center;
        }
        
        .main-header h1 {
            margin: 0;
            font-size: 2.5rem;
            font-weight: 600;
        }
        
        .main-header p {
            margin: 0.5rem 0 0 0;
            font-size: 1.1rem;
            opacity: 0.9;
        }
        
        /* Cards de configuration */
        .config-card {
            background: white;
            padding: 1.5rem;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            border-left: 4px solid var(--primary-color);
            margin-bottom: 1rem;
        }
        
        /* Étapes de processus */
        .step-card {
            background: white;
            padding: 1.5rem;
            border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            margin-bottom: 1.5rem;
            border: 1px solid var(--border-color);
            transition: all 0.3s ease;
        }
        
        .step-card:hover {
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            transform: translateY(-2px);
        }
        
        .step-header {
            display: flex;
            align-items: center;
            margin-bottom: 1rem;
            padding-bottom: 0.5rem;
            border-bottom: 2px solid var(--background-light);
        }
        
        .step-number {
            background: var(--primary-color);
            color: white;
            width: 40px;
            height: 40px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            margin-right: 1rem;
            font-size: 1.2rem;
        }
        
        .step-title {
            font-size: 1.3rem;
            font-weight: 600;
            color: var(--text-dark);
            margin: 0;
        }
        
        /* Status indicators */
        .status-success {
            background-color: var(--success-color);
            color: white;
            padding: 0.5rem 1rem;
            border-radius: 20px;
            font-weight: 500;
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
        }
        
        .status-pending {
            background-color: var(--warning-color);
            color: white;
            padding: 0.5rem 1rem;
            border-radius: 20px;
            font-weight: 500;
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
        }
        
        /* Boutons améliorés */
        .stButton > button {
            background: linear-gradient(135deg, var(--primary-color), #1565c0);
            color: white;
            border: none;
            padding: 0.75rem 2rem;
            border-radius: 25px;
            font-weight: 600;
            font-size: 1.1rem;
            transition: all 0.3s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
        }
        
        .stButton > button:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.3);
        }
        
        /* File uploader styling */
        .stFileUploader {
            border: 2px dashed var(--border-color);
            border-radius: 10px;
            padding: 1rem;
            background: var(--background-light);
            transition: all 0.3s ease;
        }
        
        .stFileUploader:hover {
            border-color: var(--primary-color);
            background: rgba(31, 119, 180, 0.05);
        }
        
        /* Progress styling */
        .stProgress .st-bo {
            background-color: var(--background-light);
        }
        
        .stProgress .st-bp {
            background: linear-gradient(90deg, var(--primary-color), var(--secondary-color));
        }
        
        /* Info boxes */
        .info-box {
            background: linear-gradient(135deg, #e3f2fd, #bbdefb);
            border-left: 4px solid var(--primary-color);
            padding: 1rem;
            border-radius: 5px;
            margin: 1rem 0;
        }
        
        /* Results summary */
        .results-summary {
            background: white;
            padding: 1.5rem;
            border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-left: 4px solid var(--success-color);
            margin-top: 1rem;
        }
        
        /* Hide Streamlit default elements */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        
        /* Custom spacing */
        .block-container {
            padding-top: 2rem;
            padding-bottom: 2rem;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # Header principal avec design amélioré
        st.markdown("""
        <div class="main-header">
            <h1><i class="fas fa-cogs"></i> Générateur Anaplan</h1>
            <p>Génération automatisée des fichiers plats pour l'intégration Anaplan</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Configuration globale dans une card
        st.markdown('<div class="config-card">', unsafe_allow_html=True)
        st.markdown("### <i class='fas fa-cog'></i> Configuration", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            exercice = st.text_input(
                "📊 Exercice (Budget / QBR / MBR)", 
                value="QBR2",
                help="Spécifiez le type d'exercice budgétaire"
            )
        with col2:
            date_version = st.date_input(
                "📅 Date de la Version",
                help="Date de création de cette version des données"
            )
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Génération globale avec design amélioré
        st.markdown("---")
        
        # Étape 1: Upload SUMMARY
        st.markdown("""
        <div class="step-card">
            <div class="step-header">
                <div class="step-number">1</div>
                <h3 class="step-title"><i class="fas fa-upload"></i> Fichier SUMMARY</h3>
            </div>
            <div style="background: #e3f2fd; padding: 1rem; border-radius: 8px; margin-bottom: 1rem; border-left: 4px solid #1976d2;">
                <strong>📊 Ce fichier permet de générer :</strong>
                <ul style="margin: 0.5rem 0 0 1rem;">
                    <li>✅ Fichier PPV Production</li>
                    <li>✅ Fichier Ventes & MP</li>
                </ul>
            </div>
        """, unsafe_allow_html=True)
        
        summary_file = st.file_uploader(
            "Sélectionner le fichier SUMMARY (Excel)",
            type=["xlsx"],
            key="summary_global",
            help="📋 Ce fichier contient les données pour PPV Production et Ventes & MP"
        )
        
        if summary_file:
            st.markdown("""
            <div class="status-success">
                <i class="fas fa-check-circle"></i> Fichier SUMMARY chargé avec succès
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="status-pending">
                <i class="fas fa-clock"></i> En attente du fichier SUMMARY
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Étape 2: Upload PPV
        st.markdown("""
        <div class="step-card">
            <div class="step-header">
                <div class="step-number">2</div>
                <h3 class="step-title"><i class="fas fa-chart-line"></i> Fichier PPV</h3>
            </div>
            <div style="background: #e8f5e9; padding: 1rem; border-radius: 8px; margin-bottom: 1rem; border-left: 4px solid #2e7d32;">
                <strong>⚗️ Ce fichier permet de générer :</strong>
                <ul style="margin: 0.5rem 0 0 1rem;">
                    <li>✅ Fichier Ratios Matières</li>
                </ul>
            </div>
        """, unsafe_allow_html=True)
        
        ppv_file = None
        if summary_file:
            ppv_file = st.file_uploader(
                "Sélectionner le fichier PPV (Excel)",
                type=["xlsx"],
                key="ppv_global",
                help="📊 Ce fichier contient les données pour les Ratios Matières"
            )
            
            if ppv_file:
                st.markdown("""
                <div class="status-success">
                    <i class="fas fa-check-circle"></i> Fichier PPV chargé avec succès
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div class="status-pending">
                    <i class="fas fa-clock"></i> En attente du fichier PPV
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="info-box">
                <i class="fas fa-info-circle"></i> 
                Veuillez d'abord charger le fichier SUMMARY pour continuer
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Étape 3: Upload Anaplan-Ref
        st.markdown("""
        <div class="step-card">
            <div class="step-header">
                <div class="step-number">3</div>
                <h3 class="step-title"><i class="fas fa-database"></i> Fichier de Référence Anaplan</h3>
            </div>
        """, unsafe_allow_html=True)
        
        anaplan_ref_file = None
        if MAPPING_AVAILABLE:
            anaplan_ref_file = st.file_uploader(
                "Sélectionner le fichier Anaplan-Ref.xlsx",
                type=["xlsx"],
                key="anaplan_ref_global",
                help="📋 Ce fichier contient les tables de référence pour le mapping des codes Anaplan"
            )
            
            if anaplan_ref_file:
                # Ajuster les colonnes selon si le sanity check a été exécuté
                if st.session_state.sanity_check_completed:
                    col1, col2, col3 = st.columns([2, 1, 1])
                else:
                    col1, col2 = st.columns([3, 1])
                    col3 = None
                
                with col1:
                    st.markdown("""
                    <div class="status-success">
                        <i class="fas fa-check-circle"></i> Fichier Anaplan-Ref chargé avec succès
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown("""
                    <style>
                    .stButton > button.sanity-checker {
                        background: linear-gradient(135deg, #d32f2f, #f44336) !important;
                        color: white !important;
                        border: none !important;
                        padding: 0.5rem 1rem !important;
                        border-radius: 20px !important;
                        font-weight: 600 !important;
                        font-size: 0.9rem !important;
                        transition: all 0.3s ease !important;
                        box-shadow: 0 2px 4px rgba(211, 47, 47, 0.3) !important;
                    }
                    .stButton > button.sanity-checker:hover {
                        transform: translateY(-2px) !important;
                        box-shadow: 0 4px 8px rgba(211, 47, 47, 0.4) !important;
                    }
                    .stButton > button.download-sanity {
                        background: linear-gradient(135deg, #2e7d32, #4caf50) !important;
                        color: white !important;
                        border: none !important;
                        padding: 0.5rem 1rem !important;
                        border-radius: 20px !important;
                        font-weight: 600 !important;
                        font-size: 0.9rem !important;
                        transition: all 0.3s ease !important;
                        box-shadow: 0 2px 4px rgba(46, 125, 50, 0.3) !important;
                    }
                    .stButton > button.download-sanity:hover {
                        transform: translateY(-2px) !important;
                        box-shadow: 0 4px 8px rgba(46, 125, 50, 0.4) !important;
                    }
                    </style>
                    """, unsafe_allow_html=True)
                    sanity_btn = st.button("🔍 Run Sanity Checker", key="sanity_checker")
                
                # Bouton de téléchargement supprimé - utiliser le téléchargement ZIP ci-dessous
                
                # Logique du Sanity Checker
                if sanity_btn:
                    with st.spinner("🔍 Exécution du Sanity Checker..."):
                        st.info("🔍 **Sanity Checker activé !**")
                        
                        # Exécuter le sanity checker avec les fichiers uploadés
                        try:
                            # Importer et exécuter le sanity checker directement
                            from sanity_checker import AnaplanSanityChecker
                            
                            # Créer l'instance avec les fichiers uploadés
                            checker = AnaplanSanityChecker(
                                summary_file=summary_file,
                                ppv_file=ppv_file
                            )
                            
                            # Exécuter le sanity check complet
                            result = checker.run_full_sanity_check()
                            
                            # Afficher les résultats
                            if result:
                                st.success("✅ Sanity Checker exécuté avec succès !")
                                
                                # Marquer comme terminé et collecter les fichiers générés
                                st.session_state.sanity_check_completed = True
                                
                                # Collecter tous les fichiers de résultats disponibles
                                potential_files = [
                                    "anaplan_ref_sanity_report.xlsx",
                                    "fichier_ppv_production_sanity_checked.xlsx",
                                    "fichier_ventes_mp_sanity_checked.xlsx", 
                                    "fichier_ratios_matieres_sanity_checked.xlsx"
                                ]
                                
                                available_files = []
                                for file in potential_files:
                                    if os.path.exists(file):
                                        available_files.append(file)
                                
                                st.session_state.sanity_check_files = available_files
                                
                                if available_files:
                                    st.info(f"📁 {len(available_files)} fichier(s) de résultats disponible(s) pour téléchargement")
                            else:
                                st.error("❌ Erreur lors de l'exécution du Sanity Checker")
                        
                        except Exception as e:
                            st.error(f"❌ Erreur lors de l'exécution du Sanity Checker : {e}")
                        
                        st.success("🎯 Sanity Checker terminé !")
                
                # Afficher le bouton de téléchargement si des fichiers sont disponibles
                if st.session_state.sanity_check_completed and st.session_state.sanity_check_files:
                    st.markdown("### 📥 Téléchargement des résultats Sanity Check")
                    
                    # Créer un ZIP avec tous les fichiers de résultats
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for file_path in st.session_state.sanity_check_files:
                            if os.path.exists(file_path):
                                with open(file_path, 'rb') as f:
                                    zip_file.writestr(os.path.basename(file_path), f.read())
                    
                    # Proposer le téléchargement
                    st.download_button(
                        "📥 Télécharger Résultats Sanity Check (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"sanity_check_results_{exercice}_{str(date_version).replace('-', '')}.zip",
                        mime="application/zip",
                        key="download_sanity_zip"
                    )
                    
                    st.info(f"📦 Package prêt avec {len(st.session_state.sanity_check_files)} fichier(s)")
                
                # Ancien code de téléchargement supprimé
            else:
                st.markdown("""
                <div class="status-pending">
                    <i class="fas fa-clock"></i> En attente du fichier Anaplan-Ref (optionnel)
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="info-box">
                <i class="fas fa-exclamation-triangle"></i> 
                Module de mapping non disponible. Seule la génération des fichiers Excel sera possible.
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Étape 4: Génération
        if summary_file and ppv_file and exercice and date_version:
            st.markdown("""
            <div class="step-card">
                <div class="step-header">
                    <div class="step-number">4</div>
                    <h3 class="step-title"><i class="fas fa-rocket"></i> Génération</h3>
                </div>
            """, unsafe_allow_html=True)
            
            if st.session_state.generation_completed:
                st.markdown("""
                <div class="status-success">
                    <i class="fas fa-check-circle"></i> Fichiers déjà générés et disponibles au téléchargement
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div class="status-success">
                    <i class="fas fa-check-double"></i> Tous les fichiers sont prêts pour la génération
                </div>
                """, unsafe_allow_html=True)
            
            # Boutons de génération et reset
            col1, col2 = st.columns([3, 1])
            with col1:
                generate_btn = st.button("🚀 Générer Tous les Fichiers", key="btn_global")
            with col2:
                if st.button("🔄 Reset", key="btn_reset"):
                    st.session_state.generated_files = {}
                    st.session_state.generation_completed = False
                    st.session_state.mapped_files = {}
                    st.session_state.mapping_completed = False
                    st.session_state.sanity_check_completed = False
                    st.session_state.sanity_check_files = []
                    st.rerun()
            
            if generate_btn:
                with st.spinner("⚙️ Génération en cours..."):
                    all_files = {}
                    
                    # Container pour les résultats
                    progress_container = st.empty()
                    
                    with progress_container.container():
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                    
                    # PPV Production
                    status_text.text("📊 Génération PPV Production...")
                    progress_bar.progress(0.2)
                    try:
                        prod_df, prod_bytes = self.generate_ppv_production_v4(ppv_file, summary_file, exercice, date_version)
                        if len(prod_df):
                            all_files["fichier_ppv_production.xlsx"] = prod_bytes
                    except Exception as e:
                        st.error(f"❌ Erreur PPV Production : {e}")
                    
                    # Ventes & MP
                    status_text.text("💰 Génération Ventes & MP...")
                    progress_bar.progress(0.5)
                    try:
                        ventes_df, ventes_bytes = self.generate_ppv_ventes_mp_v2(summary_file, exercice, date_version)
                        if len(ventes_df):
                            all_files["fichier_ventes_mp.xlsx"] = ventes_bytes
                    except Exception as e:
                        st.error(f"❌ Erreur Ventes & MP : {e}")
                    
                    # Ratios Matières
                    status_text.text("⚗️ Génération Ratios Matières...")
                    progress_bar.progress(0.8)
                    try:
                        ratios_df, ratios_bytes = self.generate_ratios_v2(ppv_file, summary_file, exercice, date_version)
                        if len(ratios_df):
                            all_files["fichier_ratios_matieres.xlsx"] = ratios_bytes
                    except Exception as e:
                        st.error(f"❌ Erreur Ratios Matières : {e}")
                    
                    # Création du ZIP
                    status_text.text("📦 Création du fichier ZIP...")
                    progress_bar.progress(1.0)
                    
                    if all_files:
                        # Sauvegarder dans la session
                        st.session_state.generated_files = all_files
                        st.session_state.generation_completed = True
                        
                        progress_container.empty()
                        st.balloons()
                        
                        # Résumé des résultats
                        st.markdown(f"""
                        <div class="results-summary">
                            <h3><i class="fas fa-check-circle"></i> Génération terminée avec succès !</h3>
                            <p><strong>{len(all_files)} fichier(s) généré(s) :</strong></p>
                            <ul>
                        """, unsafe_allow_html=True)
                        
                        for filename in all_files.keys():
                            st.markdown(f"<li>📄 {filename}</li>", unsafe_allow_html=True)
                        
                        st.markdown("</ul></div>", unsafe_allow_html=True)
                    else:
                        progress_container.empty()
                        st.error("❌ Aucun fichier n'a pu être généré. Vérifiez vos données d'entrée.")
            
            # Affichage des téléchargements si les fichiers sont générés
            if st.session_state.generation_completed and st.session_state.generated_files:
                st.markdown("### 📥 Téléchargements disponibles")
                
                # Package complet ZIP
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for filename, file_bytes in st.session_state.generated_files.items():
                        zip_file.writestr(filename, file_bytes)
                
                st.download_button(
                    "📦 Télécharger le Package Complet (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"anaplan_files_{exercice}_{str(date_version).replace('-', '')}.zip",
                    mime="application/zip",
                    key="dl_global_persistent"
                )
                
                # Téléchargements individuels
                st.markdown("#### 📄 Téléchargements individuels")
                cols = st.columns(len(st.session_state.generated_files))
                
                for i, (filename, file_bytes) in enumerate(st.session_state.generated_files.items()):
                    with cols[i]:
                        st.download_button(
                            f"📄 {filename}",
                            data=file_bytes,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_individual_{i}"
                        )
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Étape 5: Mapping et CSV
        if MAPPING_AVAILABLE and anaplan_ref_file:
            st.markdown("---")
            st.markdown("""
            <div class="step-card">
                <div class="step-header">
                    <div class="step-number">5</div>
                    <h3 class="step-title"><i class="fas fa-exchange-alt"></i> Mapping et Export CSV</h3>
                </div>
            """, unsafe_allow_html=True)
            
            if st.session_state.mapping_completed:
                st.markdown("""
                <div class="status-success">
                    <i class="fas fa-check-circle"></i> Fichiers mappés déjà générés et disponibles au téléchargement
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div class="info-box">
                    <i class="fas fa-info-circle"></i> 
                    Cette étape génère les fichiers CSV mappés avec les codes Anaplan à partir des fichiers Excel générés à l'étape précédente.
                </div>
                """, unsafe_allow_html=True)
            
            mapping_btn = st.button("🔄 Générer les fichiers CSV mappés", key="btn_mapping")
            
            if mapping_btn:
                with st.spinner("🔄 Mapping en cours..."):
                    try:
                        # Sauvegarder temporairement le fichier de référence
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_ref:
                            tmp_ref.write(anaplan_ref_file.getvalue())
                            tmp_ref_path = tmp_ref.name
                        
                        # Créer l'instance du mapper avec le fichier de référence
                        mapper = AnaplanMapper(ref_file_path=tmp_ref_path)
                        
                        # Utiliser les fichiers de la session si disponibles, sinon générer
                        excel_files = {}
                        
                        if st.session_state.generated_files:
                            # Utiliser les fichiers déjà générés
                            st.info("📊 Utilisation des fichiers Excel déjà générés...")
                            for filename, file_bytes in st.session_state.generated_files.items():
                                if "production" in filename:
                                    with open("fichier_ppv_production.xlsx", 'wb') as f:
                                        f.write(file_bytes)
                                    excel_files['ppv_production'] = "fichier_ppv_production.xlsx"
                                elif "ventes_mp" in filename:
                                    with open("fichier_ventes_mp.xlsx", 'wb') as f:
                                        f.write(file_bytes)
                                    excel_files['ventes_mp'] = "fichier_ventes_mp.xlsx"
                                elif "ratios" in filename:
                                    with open("fichier_ratios_matieres.xlsx", 'wb') as f:
                                        f.write(file_bytes)
                                    excel_files['ratios_matieres'] = "fichier_ratios_matieres.xlsx"
                        else:
                            # Générer les fichiers Excel d'abord
                            st.info("📊 Génération des fichiers Excel...")
                            
                            # PPV Production
                            try:
                                prod_df, prod_bytes = self.generate_ppv_production_v4(ppv_file, summary_file, exercice, date_version)
                                if len(prod_df):
                                    excel_filename = "fichier_ppv_production.xlsx"
                                    with open(excel_filename, 'wb') as f:
                                        f.write(prod_bytes)
                                    excel_files['ppv_production'] = excel_filename
                            except Exception as e:
                                st.error(f"❌ Erreur PPV Production : {e}")
                            
                            # Ventes & MP
                            try:
                                ventes_df, ventes_bytes = self.generate_ppv_ventes_mp_v2(summary_file, exercice, date_version)
                                if len(ventes_df):
                                    excel_filename = "fichier_ventes_mp.xlsx"
                                    with open(excel_filename, 'wb') as f:
                                        f.write(ventes_bytes)
                                    excel_files['ventes_mp'] = excel_filename
                            except Exception as e:
                                st.error(f"❌ Erreur Ventes & MP : {e}")
                            
                            # Ratios Matières
                            try:
                                ratios_df, ratios_bytes = self.generate_ratios_v2(ppv_file, summary_file, exercice, date_version)
                                if len(ratios_df):
                                    excel_filename = "fichier_ratios_matieres.xlsx"
                                    with open(excel_filename, 'wb') as f:
                                        f.write(ratios_bytes)
                                    excel_files['ratios_matieres'] = excel_filename
                            except Exception as e:
                                st.error(f"❌ Erreur Ratios Matières : {e}")
                        
                        # Maintenant, traiter les fichiers pour le mapping
                        st.info("🔄 Application du mapping...")
                        results = mapper.process_excel_files(
                            ppv_production_file=excel_files.get('ppv_production'),
                            ventes_mp_file=excel_files.get('ventes_mp'),
                            ratios_matieres_file=excel_files.get('ratios_matieres')
                        )
                        
                        # Nettoyer le fichier temporaire
                        os.unlink(tmp_ref_path)
                        
                        # Afficher les résultats et sauvegarder dans la session
                        if results and os.path.exists("MAPPING"):
                            # Ne garder que les fichiers CSV pour téléchargement (les versions Excel sont ignorées)
                            mapping_files = [f for f in os.listdir("MAPPING") if f.endswith('.csv')]
                            
                            if mapping_files:
                                # Lire et sauvegarder les fichiers mappés dans la session
                                mapped_files_data = {}
                                for mapping_file in mapping_files:
                                    file_path = os.path.join("MAPPING", mapping_file)
                                    if os.path.exists(file_path):
                                        with open(file_path, 'rb') as f:
                                            mapped_files_data[mapping_file] = f.read()
                                
                                st.session_state.mapped_files = mapped_files_data
                                st.session_state.mapping_completed = True
                        
                    except Exception as e:
                        st.error(f"❌ Erreur lors du mapping : {e}")
            
            # Affichage des téléchargements si le mapping est terminé
            if st.session_state.mapping_completed and st.session_state.mapped_files:
                st.markdown("### 📥 Fichiers mappés disponibles")
                
                # ZIP avec tous les fichiers mappés
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for filename, file_data in st.session_state.mapped_files.items():
                        zip_file.writestr(filename, file_data)
                
                st.download_button(
                    "📥 Télécharger tous les fichiers mappés (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"anaplan_mapped_{exercice}_{str(date_version).replace('-', '')}.zip",
                    mime="application/zip",
                    key="dl_mapped_zip_persistent"
                )
                
                # Téléchargements individuels
                st.markdown("#### 📁 Téléchargements individuels")
                
                # Séparer les fichiers Excel et CSV
                excel_files = {k: v for k, v in st.session_state.mapped_files.items() if k.endswith('.xlsx')}
                csv_files = {k: v for k, v in st.session_state.mapped_files.items() if k.endswith('.csv')}
                
                if excel_files:
                    st.markdown("##### 📊 Fichiers Excel (avec couleurs)")
                    cols = st.columns(min(3, len(excel_files)))
                    for i, (excel_file, excel_data) in enumerate(excel_files.items()):
                        with cols[i % 3]:
                            st.download_button(
                                f"📄 {excel_file}",
                                data=excel_data,
                                file_name=excel_file,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_excel_mapped_{i}"
                            )
                
                if csv_files:
                    st.markdown("##### 📋 Fichiers CSV pour Anaplan")
                    cols = st.columns(min(3, len(csv_files)))
                    for i, (csv_file, csv_data) in enumerate(csv_files.items()):
                        with cols[i % 3]:
                            st.download_button(
                                f"📄 {csv_file}",
                                data=csv_data,
                                file_name=csv_file,
                                mime="text/csv",
                                key=f"dl_csv_mapped_{i}"
                            )
            
            st.markdown('</div>', unsafe_allow_html=True)
        
        elif not exercice or not date_version:
            st.warning("⚠️ Veuillez compléter la configuration (Exercice et Date) ci-dessus")
        
        # Footer informatif
        st.markdown("---")
        st.markdown("""
        <div style="text-align: center; color: #666; padding: 1rem;">
            <i class="fas fa-info-circle"></i> 
            Générateur Anaplan - Automatisation des fichiers plats pour intégration ERP
        </div>
        """, unsafe_allow_html=True)

    @staticmethod
    def clean_tuple_string(val):
        """Nettoie les valeurs lues comme "('Brazil', 820)" pour ne garder que la partie pertinente.

        Si la valeur est un tuple ou une chaîne représentant un tuple, on renvoie le premier
        élément. Sinon on renvoie la valeur convertie en chaîne telle quelle.
        """
        import ast
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ""

        # Si c'est déjà un vrai tuple / liste
        if isinstance(val, (tuple, list)):
            result = str(val[0]).strip()
            # Normaliser "NAN" comme valeur vide
            return "" if result.upper() == "NAN" else result

        # Convertir en chaîne pour traitement
        s = str(val).strip()

        # Tenter une évaluation litterale du tuple
        if s.startswith("(") and s.endswith(")"):
            try:
                lit = ast.literal_eval(s)
                if isinstance(lit, (tuple, list)) and len(lit):
                    result = str(lit[0]).strip()
                    # Normaliser "NAN" comme valeur vide
                    return "" if result.upper() == "NAN" else result
            except Exception:
                pass

            # Fallback manuel: retirer parenthèses puis split par virgule
            s_no_paren = s.strip("()")
            if "," in s_no_paren:
                result = s_no_paren.split(",")[0].strip().strip("'\"")
                # Normaliser "NAN" comme valeur vide
                return "" if result.upper() == "NAN" else result
        # Si commence par quote puis virgule
        if "," in s:
            first_part = s.split(",")[0]
            result = first_part.strip().strip("()'\"")
            # Normaliser "NAN" comme valeur vide
            return "" if result.upper() == "NAN" else result

        result = s.strip().strip("()'\"")
        # Normaliser "NAN" comme valeur vide
        return "" if result.upper() == "NAN" else result

    # ---------------------------------------------------------------------------
    # Utilitaires volumes (mois / trimestre)
    # ---------------------------------------------------------------------------

    @staticmethod
    def _strip_accents(text: str) -> str:
        """Supprime les accents pour faciliter la comparaison."""
        text = str(text) if text is not None else ""
        return "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")

    def _month_to_number(self, name: str):
        """Convertit un libellé de mois en numéro (1-12)."""
        if name is None:
            return None
        key = self._strip_accents(name).upper().strip()
        if key.isdigit():
            num = int(key)
            return num if 1 <= num <= 12 else None
        return self.month_name_to_num.get(key)

    def _process_volume_row(self, row, prefix: str, site_field: str, qual_field: str, operation_label: str, *, exercice: str, date_version, annee: int):
        """Crée les lignes mensuelles à partir d'une ligne source (Extraction/Physical/Chemical)."""
        monthly_volumes = {}
        quarter_volumes = {}

        for col, val in row.items():
            if not isinstance(col, str):
                continue
            if not col.startswith(f"{prefix}["):
                continue
            if pd.isna(val):
                continue

            # Extraire le libellé entre crochets
            inner = col[len(prefix) + 1 : -1]  # supprime prefix[
            inner_norm = self._strip_accents(inner).upper().strip()

            if inner_norm in self.quarter_to_months:
                quarter_volumes[inner_norm] = float(val)
            else:
                month_num = self._month_to_number(inner_norm)
                if month_num:
                    monthly_volumes[month_num] = float(val)

        rows = []
        # Ajouter les volumes mensuels en priorité
        for month, vol in monthly_volumes.items():
            rows.append({
                "Exercice": exercice,
                "Date de la Version": str(date_version),
                "Année": annee,
                "Mois": month,
                "Site/Entité": row.get(site_field, ""),
                "Qualité": row.get(qual_field, ""),
                "Partenaire Groupe": "",
                "Type Operation": operation_label,
                "Operation": row.get('Treatment', ''),
                "VOLUME (T)": vol * 1000, # Convertir en t
            })

        seen_months = set(monthly_volumes.keys())

        # Ajouter maintenant les volumes trimestriels (répartis /3) en évitant doublons
        for q, vol in quarter_volumes.items():
            per_month = vol / 3
            for month in self.quarter_to_months[q]:
                if month in seen_months:
                    continue
                rows.append({
                    "Exercice": exercice,
                    "Date de la Version": str(date_version),
                    "Année": annee,
                    "Mois": month,
                    "Site/Entité": row.get(site_field, ""),
                    "Qualité": row.get(qual_field, ""),
                    "Partenaire Groupe": "",
                    "Type Operation": operation_label,
                    "Operation": row.get('Treatment', ''),
                    "VOLUME (T)": per_month * 1000, # Convertir en t
                })

        return rows

# ---------------------------------------------------------------------------
# Point d'entrée principal
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # Instancier et lancer l'application
    app = AnaplanMainApp()
    app.run() 