import pandas as pd
import os
from datetime import datetime
import re

class AnaplanMapper:
    """Classe pour mapper les données Anaplan avec les tables de référence et générer des CSV."""
    
    def __init__(self, ref_file_path="FICHIERS ANAPLAN/Anaplan-Ref.xlsx"):
        """Initialise le mapper en chargeant toutes les tables de référence."""
        self.ref_file_path = ref_file_path
        self.ref_tables = {}
        self.load_reference_tables()
    
    def load_reference_tables(self):
        """Charge toutes les tables de référence depuis le fichier Anaplan-Ref.xlsx."""
        try:
            xls = pd.ExcelFile(self.ref_file_path)
            
            # Charger les tables principales
            self.ref_tables['type_transaction'] = pd.read_excel(self.ref_file_path, sheet_name='Type de transaction')
            self.ref_tables['pays'] = pd.read_excel(self.ref_file_path, sheet_name='Pays')
            self.ref_tables['devise'] = pd.read_excel(self.ref_file_path, sheet_name='Devise')
            self.ref_tables['qualite'] = pd.read_excel(self.ref_file_path, sheet_name='Qualité')
            self.ref_tables['exercice'] = pd.read_excel(self.ref_file_path, sheet_name='Exercice')
            self.ref_tables['site_entite'] = pd.read_excel(self.ref_file_path, sheet_name='SiteEntité')
            self.ref_tables['partenaire_groupe'] = pd.read_excel(self.ref_file_path, sheet_name='Partenaire Groupe')
            self.ref_tables['operation'] = pd.read_excel(self.ref_file_path, sheet_name='Opération')
            self.ref_tables['ratios'] = pd.read_excel(self.ref_file_path, sheet_name='Ckecklibs')
            self.ref_tables['site_entite_vente'] = pd.read_excel(self.ref_file_path, sheet_name='SiteEntité_vente')
            self.ref_tables['site_entite_prod'] = pd.read_excel(self.ref_file_path, sheet_name='SiteEntité_Prod')
                
        except Exception as e:
            print(f"❌ Erreur lors du chargement des tables de référence: {e}")
            raise
    
    def normalize_text(self, text):
        """Normalise le texte pour faciliter les correspondances."""
        if pd.isna(text):
            return ""
        text = str(text).strip().upper()
        # Enlever les accents et caractères spéciaux
        replacements = {
            'À': 'A', 'Á': 'A', 'Â': 'A', 'Ã': 'A', 'Ä': 'A',
            'È': 'E', 'É': 'E', 'Ê': 'E', 'Ë': 'E',
            'Ì': 'I', 'Í': 'I', 'Î': 'I', 'Ï': 'I',
            'Ò': 'O', 'Ó': 'O', 'Ô': 'O', 'Õ': 'O', 'Ö': 'O',
            'Ù': 'U', 'Ú': 'U', 'Û': 'U', 'Ü': 'U',
            'Ç': 'C', 'Ñ': 'N'
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        return text
    
    def normalize_column_names(self, df):
        """Normalise les noms de colonnes en enlevant les accents."""
        # Dictionnaire pour enlever les accents des noms de colonnes
        accent_replacements = {
            'À': 'A', 'Á': 'A', 'Â': 'A', 'Ã': 'A', 'Ä': 'A', 'à': 'a', 'á': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a',
            'È': 'E', 'É': 'E', 'Ê': 'E', 'Ë': 'E', 'è': 'e', 'é': 'e', 'ê': 'e', 'ë': 'e',
            'Ì': 'I', 'Í': 'I', 'Î': 'I', 'Ï': 'I', 'ì': 'i', 'í': 'i', 'î': 'i', 'ï': 'i',
            'Ò': 'O', 'Ó': 'O', 'Ô': 'O', 'Õ': 'O', 'Ö': 'O', 'ò': 'o', 'ó': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
            'Ù': 'U', 'Ú': 'U', 'Û': 'U', 'Ü': 'U', 'ù': 'u', 'ú': 'u', 'û': 'u', 'ü': 'u',
            'Ç': 'C', 'ç': 'c', 'Ñ': 'N', 'ñ': 'n'
        }
        
        # Créer un dictionnaire de mapping pour renommer les colonnes
        column_mapping = {}
        for col in df.columns:
            normalized_col = str(col)
            for accent, replacement in accent_replacements.items():
                normalized_col = normalized_col.replace(accent, replacement)
            column_mapping[col] = normalized_col
        
        # Renommer les colonnes
        df_normalized = df.rename(columns=column_mapping)
        print(f"✅ Noms de colonnes normalisés (accents supprimés)")
        return df_normalized
    
    def map_type_transaction(self, type_transaction):
        """Mappe le type de transaction vers le code correspondant."""
        type_norm = self.normalize_text(type_transaction)
        
        # Mapping manuel basé sur les patterns
        if 'EXPORT' in type_norm:
            return 'VEX'
        elif 'LOCAL' in type_norm or 'LOCALE' in type_norm:
            return 'VLC'
        elif 'ACHAT' in type_norm and 'MP' in type_norm:
            return 'AMP'
        elif 'CONSOMMATION' in type_norm and 'MP' in type_norm:
            return 'CMP'
        else:
            # Recherche exacte dans la table
            for _, row in self.ref_tables['type_transaction'].iterrows():
                if self.normalize_text(row['Libellé Transaction']) == type_norm:
                    return row['Code Transaction']
        
        return type_transaction  # Retourner tel quel si pas de correspondance
    
    def map_pays(self, pays):
        """Mappe le nom de pays vers le code correspondant."""
        pays_norm = self.normalize_text(pays)
        
        # Recherche exacte
        for _, row in self.ref_tables['pays'].iterrows():
            if self.normalize_text(row['Pays']) == pays_norm:
                return row['Code Pays']
        
        # Recherche partielle
        for _, row in self.ref_tables['pays'].iterrows():
            if pays_norm in self.normalize_text(row['Pays']) or self.normalize_text(row['Pays']) in pays_norm:
                return row['Code Pays']
        
        return pays  # Retourner tel quel si pas de correspondance
    
    def map_devise(self, devise):
        """Mappe la devise vers le code correspondant."""
        devise_norm = self.normalize_text(devise)
        
        for _, row in self.ref_tables['devise'].iterrows():
            if self.normalize_text(row['Devises']) == devise_norm:
                return row['Code Devises']
        
        return devise  # Retourner tel quel si pas de correspondance
    
    def map_qualite(self, qualite):
        """Mappe la qualité vers le code SAP correspondant en priorité par Libellé BS."""
        qualite_norm = self.normalize_text(qualite)
        
        # PRIORITÉ 1: Recherche exacte par libellé BS → Code SAP
        for _, row in self.ref_tables['qualite'].iterrows():
            if self.normalize_text(row['Libellé BS']) == qualite_norm:
                return row['Code SAP']
        
        # PRIORITÉ 2: Recherche exacte par libellé SAP → Code SAP  
        for _, row in self.ref_tables['qualite'].iterrows():
            if self.normalize_text(row['Libellé SAP']) == qualite_norm:
                return row['Code SAP']
        
        # PRIORITÉ 3: Recherche partielle par libellé BS d'abord
        for _, row in self.ref_tables['qualite'].iterrows():
            libelle_bs_norm = self.normalize_text(row['Libellé BS'])
            if qualite_norm in libelle_bs_norm or libelle_bs_norm in qualite_norm:
                return row['Code SAP']
        
        # PRIORITÉ 4: Recherche partielle par libellé SAP
        for _, row in self.ref_tables['qualite'].iterrows():
            libelle_sap_norm = self.normalize_text(row['Libellé SAP'])
            if qualite_norm in libelle_sap_norm or libelle_sap_norm in qualite_norm:
                return row['Code SAP']
        
        return qualite  # Retourner tel quel si pas de correspondance
    
    def map_exercice(self, exercice):
        """Mappe l'exercice vers le code correspondant."""
        exercice_norm = self.normalize_text(exercice)
        
        for _, row in self.ref_tables['exercice'].iterrows():
            if self.normalize_text(row['Libellé']) == exercice_norm:
                return row['Code Exercice']
        
        return exercice  # Retourner tel quel si pas de correspondance
    
    def map_site_entite(self, site_entite):
        """Mappe le site/entité vers le code correspondant."""
        site_norm = self.normalize_text(site_entite)
        
        # Recherche exacte
        for _, row in self.ref_tables['site_entite'].iterrows():
            if self.normalize_text(row['Libellé Site/Entité']) == site_norm:
                return row['Code P/CC']
        
        # Recherche partielle
        for _, row in self.ref_tables['site_entite'].iterrows():
            libelle_norm = self.normalize_text(row['Libellé Site/Entité'])
            if site_norm in libelle_norm or libelle_norm in site_norm:
                return row['Code P/CC']
        
        return site_entite  # Retourner tel quel si pas de correspondance

    # def map_site_entite_prod(self, site_entite):
    #     """Mappe le site/entité vers le code correspondant."""
    #     site_norm = self.normalize_text(site_entite)
    #
    #     # Recherche exacte
    #     for _, row in self.ref_tables['site_entite_prod'].iterrows():
    #         if self.normalize_text(row['Libellé Site/Entité']) == site_norm:
    #             return row['Code P/CC']
    #
    #     # Recherche partielle
    #     for _, row in self.ref_tables['site_entite'].iterrows():
    #         libelle_norm = self.normalize_text(row['Libellé Site/Entité'])
    #         if site_norm in libelle_norm or libelle_norm in site_norm:
    #             return row['Code P/CC']
    #
    #     return site_entite  # Retourner tel quel si pas de correspondance


    def map_site_entite_prod(self, site_entite, qualite):
        """Mapping spécialisé pour Site/Entité dans PPV_production : Site/Entité#Qualité → code SAP_2"""
        # Normaliser les valeurs d'entrée
        site_entite_norm = self.normalize_text(site_entite)
        qualite_norm = self.normalize_text(qualite)

        # Créer la clé de recherche : Site/Entité#Qualité
        search_key = f"{site_entite_norm}#{qualite_norm}"

        # Rechercher dans la table SiteEntité_prod
        for _, row in self.ref_tables['site_entite_prod'].iterrows():
            ref_key = self.normalize_text(row['Site/Entité#Qualité'])
            if ref_key == search_key:
                return row['code SAP_2']

        # Si pas de correspondance, retourner Site/Entité original
        return site_entite

    def map_partenaire_groupe(self, partenaire):
        """Mappe le partenaire groupe vers le code correspondant."""
        partenaire_norm = self.normalize_text(partenaire)
        
        for _, row in self.ref_tables['partenaire_groupe'].iterrows():
            if self.normalize_text(row['Partenaire Groupe']) == partenaire_norm:
                return row['Code Partenaire Groupe']
        
        return partenaire  # Retourner tel quel si pas de correspondance
    
    def map_operation(self, operation):
        """Mappe l'opération vers le code correspondant."""
        operation_norm = self.normalize_text(operation)
        
        for _, row in self.ref_tables['operation'].iterrows():
            if self.normalize_text(row['Opération BS']) == operation_norm:
                return row['Macro Opération']
        
        return operation  # Retourner tel quel si pas de correspondance
    
    def map_site_entite_ventes_specialise(self, site_entite, qualite):
        """Mapping spécialisé pour Site/Entité dans Ventes & MP : Site/Entité#Qualité → code SAP_2"""
        # Normaliser les valeurs d'entrée
        site_entite_norm = self.normalize_text(site_entite)
        qualite_norm = self.normalize_text(qualite)

        # Créer la clé de recherche : Site/Entité#Qualité
        search_key = f"{site_entite_norm}#{qualite_norm}"

        # Rechercher dans la table SiteEntité_vente
        for _, row in self.ref_tables['site_entite_vente'].iterrows():
            ref_key = self.normalize_text(row['Site/Entité#Qualité'])
            if ref_key == search_key:
                return row['code SAP_2']

        # Si pas de correspondance, retourner Site/Entité original
        return site_entite
    
    def map_ppv_production(self, df):
        """Applique le mapping sur le DataFrame PPV Production."""

        
        df_mapped = df.copy()
        
        # Remplacer directement les valeurs dans les colonnes existantes
        if 'Site/Entité' in df_mapped.columns:
            # df_mapped['Site/Entité'] = df_mapped['Site/Entité'].apply(self.map_site_entite_prod)
            df_mapped['Site/Entité'] = df_mapped.apply(
                lambda row: self.map_site_entite_prod(row['Site/Entité'], row['Qualité']),
                axis=1
            )

        if 'Qualité' in df_mapped.columns:
            df_mapped['Qualité'] = df_mapped['Qualité'].apply(self.map_qualite)
        
        if 'Operation' in df_mapped.columns:
            df_mapped['Operation'] = df_mapped['Operation'].apply(self.map_operation)
        
        # Note: PPV Production n'a pas de colonne "Bloc", elle utilise "Opération"
        
        # Normaliser les noms de colonnes (enlever les accents)
        df_mapped = self.normalize_column_names(df_mapped)

        return df_mapped
    
    def map_ventes_mp(self, df):
        """Applique le mapping sur le DataFrame Ventes & MP."""

        
        df_mapped = df.copy()
        
        # Remplacer directement les valeurs dans les colonnes existantes
        if 'Type de transaction' in df_mapped.columns:
            df_mapped['Type de transaction'] = df_mapped['Type de transaction'].apply(self.map_type_transaction)
        
        # Mapping spécialisé pour Site/Entité dans Ventes & MP : Site/Entité#Qualité → code SAP_2
        if 'Site/Entité' in df_mapped.columns and 'Qualité' in df_mapped.columns:
            df_mapped['Site/Entité'] = df_mapped.apply(
                lambda row: self.map_site_entite_ventes_specialise(row['Site/Entité'], row['Qualité']),
                axis=1
            )
        
        if 'Qualité' in df_mapped.columns:
            df_mapped['Qualité'] = df_mapped['Qualité'].apply(self.map_qualite)
        
        if 'Partenaire Groupe' in df_mapped.columns:
            df_mapped['Partenaire Groupe'] = df_mapped['Partenaire Groupe'].apply(self.map_partenaire_groupe)
        
        if 'Pays' in df_mapped.columns:
            df_mapped['Pays'] = df_mapped['Pays'].apply(self.map_pays)
        
        if 'Devise' in df_mapped.columns:
            df_mapped['Devise'] = df_mapped['Devise'].apply(self.map_devise)
        
        # Supprimer la colonne Bloc pour les fichiers mappés
        if 'Bloc' in df_mapped.columns:
            df_mapped = df_mapped.drop('Bloc', axis=1)
        
        # Supprimer les colonnes non souhaitées après mapping
        columns_to_remove_ventes = ["TypeProduct"]
        for col in columns_to_remove_ventes:
            if col in df_mapped.columns:
                df_mapped = df_mapped.drop(col, axis=1)
        
        # Multiplier les volumes par 1000 pour conversion en tonnes
        if 'VOLUME (T)' in df_mapped.columns:
            # Convertir en numérique d'abord, en remplaçant les virgules par des points
            df_mapped['VOLUME (T)'] = df_mapped['VOLUME (T)'].astype(str).str.replace(',', '.')
            df_mapped['VOLUME (T)'] = pd.to_numeric(df_mapped['VOLUME (T)'], errors='coerce').fillna(0)
            # Multiplier par 1000 pour conversion en tonnes, UPDATE IL FAUT LAISSER LES VALEURS EN kt SELON LA NOUVELLE DE DEMANDE DE ZAKARIA
            df_mapped['VOLUME (T)'] = df_mapped['VOLUME (T)'] * 1
            print(f"✅ Volumes multipliés par 1000 pour conversion en tonnes dans Ventes MP")
        
        # Normaliser les noms de colonnes (enlever les accents)
        df_mapped = self.normalize_column_names(df_mapped)

        return df_mapped
    
    def map_ratios_matieres(self, df):
        """Applique le mapping sur le DataFrame Ratios Matières."""

        
        df_mapped = df.copy()
        
        # FILTRER D'ABORD : Ne garder que les lignes avec Ratio Moyen Pondéré non vide
        if 'Ratio Moyen Pondéré' in df_mapped.columns:
            # Filtrer les lignes où Ratio Moyen Pondéré n'est pas vide/NaN/0
            df_mapped = df_mapped[
                (df_mapped['Ratio Moyen Pondéré'].notna()) & 
                (df_mapped['Ratio Moyen Pondéré'] != '') & 
                (df_mapped['Ratio Moyen Pondéré'] != 0)
            ].copy()
            print(f"✅ Filtrage ratios : {len(df)} → {len(df_mapped)} lignes (lignes avec Ratio Moyen Pondéré uniquement)")
        
        # Remplacer directement les valeurs dans les colonnes existantes
        if 'Site/Entité' in df_mapped.columns:
            df_mapped['Site/Entité'] = df_mapped['Site/Entité'].apply(self.map_site_entite)
        
        if 'Qualité' in df_mapped.columns:
            df_mapped['Qualité'] = df_mapped['Qualité'].apply(self.map_qualite)
        
        if 'Type Consommation Spécifique' in df_mapped.columns:
            df_mapped['Type Consommation Spécifique'] = df_mapped['Type Consommation Spécifique'].apply(self.map_qualite)
        
        # Supprimer la colonne Bloc pour les fichiers mappés
        if 'Bloc' in df_mapped.columns:
            df_mapped = df_mapped.drop('Bloc', axis=1)
        
        # Supprimer les colonnes non souhaitées après mapping
        columns_to_remove_ratios = ["Ratio", "Volume Produit (T)", "Ratio Pondéré"]
        for col in columns_to_remove_ratios:
            if col in df_mapped.columns:
                df_mapped = df_mapped.drop(col, axis=1)
        
        # Déduplication finale pour éviter les doublons logiques après mapping
        df_mapped = df_mapped.drop_duplicates(subset=["Site/Entité", "Qualité", "Type Consommation Spécifique"]).reset_index(drop=True)

        # Renommer la colonne Ratio Moyen Pondéré en Ratio (conformité export)
        if "Ratio Moyen Pondéré" in df_mapped.columns:
            df_mapped = df_mapped.rename(columns={"Ratio Moyen Pondéré": "Ratio"})
        
        # Normaliser les noms de colonnes (enlever les accents)
        df_mapped = self.normalize_column_names(df_mapped)

        return df_mapped

    def map_ratios_matieres_v2(self, df):
        """Applique le mapping sur le DataFrame Ratios Matières."""

        df_mapped = df.copy()

        if 'Site/Entité' in df_mapped.columns:
            df_mapped['Site/Entité'] = df_mapped['Site/Entité'].apply(self.map_site_entite)

        if 'Qualité' in df_mapped.columns:
            df_mapped['Qualité'] = df_mapped['Qualité'].apply(self.map_qualite)

        if 'Type Consommation Spécifique' in df_mapped.columns:
            df_mapped['Type Consommation Spécifique'] = df_mapped['Type Consommation Spécifique'].apply(
                self.map_qualite)

        if 'Bloc' in df_mapped.columns:
            df_mapped = df_mapped.drop('Bloc', axis=1)

        df_mapped = df_mapped.drop_duplicates(
            subset=["Site/Entité", "Qualité", "Type Consommation Spécifique"]
        ).reset_index(drop=True)

        df_mapped = self.normalize_column_names(df_mapped)

        return df_mapped

    def verify_ratio(self, row):
        """Vérifie si le ratio existe dans la table de référence."""
        qualite = self.normalize_text(row.get('Qualité', ''))
        conso_spec = self.normalize_text(row.get('Type Consommation Spécifique', ''))
        
        # Rechercher dans la table des ratios
        for _, ref_row in self.ref_tables['ratios'].iterrows():
            ref_qualite = self.normalize_text(ref_row['Lib Qualité'])
            ref_conso = self.normalize_text(ref_row['Lib Conso Spec'])
            
            if ref_qualite == qualite and ref_conso == conso_spec:
                return ref_row['Ratio']
        
        return row.get('Ratio', '')  # Retourner le ratio original si pas trouvé
    
    def save_to_csv(self, df, filename, output_dir="MAPPING"):
        """Sauvegarde le DataFrame en CSV avec les bonnes règles de formatage."""
        # Créer le dossier de sortie s'il n'existe pas
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        filepath = os.path.join(output_dir, filename)
        
        # Corriger les headers pour avoir les accents
        df_copy = df.copy()
        
        # Mapping des corrections d'accents pour les headers
        header_corrections = {
            'Site/Entite': 'Site/Entité',
            'Qualite': 'Qualité',
            'Annee': 'Année'
        }
        
        # Appliquer les corrections aux noms de colonnes
        new_columns = []
        for col in df_copy.columns:
            if col in header_corrections:
                new_columns.append(header_corrections[col])
            else:
                new_columns.append(col)
        
        df_copy.columns = new_columns
        
        # Sauvegarder en CSV avec les règles Anaplan
        df_copy.to_csv(filepath, 
                 index=False,
                 encoding='utf-8',
                 sep=';',  # Séparateur point-virgule
                 decimal=',',  # Séparateur décimal virgule
                 quoting=0)  # Pas de guillemets
        

        return filepath
    
    def save_to_excel_with_colors(self, df, original_excel_file, filename, output_dir="MAPPING"):
        """Sauvegarde le DataFrame en Excel en conservant les couleurs originales."""
        # Créer le dossier de sortie s'il n'existe pas
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        filepath = os.path.join(output_dir, filename)
        
        try:
            # Ouvrir le fichier Excel original pour récupérer les formats
            from openpyxl import load_workbook
            from openpyxl.utils.dataframe import dataframe_to_rows
            
            # Charger le workbook original
            wb_original = load_workbook(original_excel_file)
            ws_original = wb_original.active
            
            # Créer un nouveau workbook avec les données mappées
            from openpyxl import Workbook
            wb_new = Workbook()
            ws_new = wb_new.active
            
            # Copier les données mappées
            for r in dataframe_to_rows(df, index=False, header=True):
                ws_new.append(r)
            
            # Copier les couleurs de fond ligne par ligne
            max_row = min(ws_original.max_row, ws_new.max_row)
            max_col = min(ws_original.max_column, ws_new.max_column)
            
            for row in range(1, max_row + 1):
                for col in range(1, max_col + 1):
                    original_cell = ws_original.cell(row=row, column=col)
                    new_cell = ws_new.cell(row=row, column=col)
                    
                    try:
                        # Copier le format de couleur de fond
                        if original_cell.fill and hasattr(original_cell.fill, 'start_color'):
                            from openpyxl.styles import PatternFill
                            new_cell.fill = PatternFill(
                                start_color=original_cell.fill.start_color,
                                end_color=original_cell.fill.end_color,
                                fill_type=original_cell.fill.fill_type
                            )
                        
                        # Copier la couleur de police
                        if original_cell.font and hasattr(original_cell.font, 'color'):
                            from openpyxl.styles import Font
                            new_cell.font = Font(
                                color=original_cell.font.color,
                                bold=original_cell.font.bold,
                                italic=original_cell.font.italic
                            )
                    except Exception:
                        # Ignorer les erreurs de copie de style
                        pass
            
            # Sauvegarder
            wb_new.save(filepath)

            
        except Exception as e:
            # Fallback: sauvegarder normalement en cas d'erreur
            with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        return filepath
    
    def process_excel_files(self, 
                           ppv_production_file=None,
                           ventes_mp_file=None,
                           ratios_matieres_file=None):
        """Traite les fichiers Excel et génère les CSV mappés."""
        

        
        results = {}
        
        # Traiter PPV Production
        if ppv_production_file and os.path.exists(ppv_production_file):
            try:

                df_prod = pd.read_excel(ppv_production_file)
                # Remove "_source" de "PPV Production Mapped"
                df_prod.drop(columns=['_source'], inplace=True)
                # Map data
                df_prod_mapped = self.map_ppv_production(df_prod)
                # Sauvegarder en Excel avec couleurs + CSV
                excel_file = self.save_to_excel_with_colors(df_prod_mapped, ppv_production_file, "PPV_Production_mapped.xlsx")
                csv_file = self.save_to_csv(df_prod_mapped, "PPV_Production_mapped.csv")
                
                results['ppv_production'] = {
                    'original_rows': len(df_prod),
                    'mapped_rows': len(df_prod_mapped),
                    'excel_file': excel_file,
                    'csv_file': csv_file
                }
            except Exception as e:
                print(f"❌ Erreur PPV Production: {e}")
                results['ppv_production'] = {'error': str(e)}
        
        # Traiter Ventes & MP
        if ventes_mp_file and os.path.exists(ventes_mp_file):
            try:

                df_ventes = pd.read_excel(ventes_mp_file)
                df_ventes_mapped = self.map_ventes_mp(df_ventes)
                
                # Sauvegarder en Excel avec couleurs + CSV
                excel_file = self.save_to_excel_with_colors(df_ventes_mapped, ventes_mp_file, "Ventes_MP_mapped.xlsx")
                csv_file = self.save_to_csv(df_ventes_mapped, "Ventes_MP_mapped.csv")
                
                results['ventes_mp'] = {
                    'original_rows': len(df_ventes),
                    'mapped_rows': len(df_ventes_mapped),
                    'excel_file': excel_file,
                    'csv_file': csv_file
                }
            except Exception as e:
                print(f"❌ Erreur Ventes & MP: {e}")
                results['ventes_mp'] = {'error': str(e)}
        
        # Traiter Ratios Matières
        if ratios_matieres_file and os.path.exists(ratios_matieres_file):
            try:

                df_ratios = pd.read_excel(ratios_matieres_file)
                df_ratios_mapped = self.map_ratios_matieres_v2(df_ratios)

                # Ratio MP with mean value
                df_ratios_with_mean = pd.read_excel(ratios_matieres_file, sheet_name="Fichier Plat Ratios")
                df_ratios_with_mean_mapped = self.map_ratios_matieres_v2(df_ratios_with_mean)
                
                # Sauvegarder en Excel avec couleurs + CSV
                excel_file = self.save_to_excel_with_colors(df_ratios_mapped, ratios_matieres_file, "Ratios_Matieres_mapped.xlsx")
                csv_file = self.save_to_csv(df_ratios_mapped, "Ratios_Matieres_mapped.csv")

                # Sauvegarder en Excel avec couleurs + CSV with mean value
                excel_file_with_mean = self.save_to_excel_with_colors(df_ratios_with_mean_mapped, ratios_matieres_file, "Ratios_Moyen_Matieres_mapped.xlsx")
                csv_file_with_mean = self.save_to_csv(df_ratios_with_mean_mapped, "Ratios_Moyen_Matieres_mapped.csv")

                results['ratios_matieres'] = {
                    'original_rows': len(df_ratios),
                    'mapped_rows': len(df_ratios_mapped),
                    'excel_file': excel_file,
                    'csv_file': csv_file,
                    'excel_file_with_mean': excel_file_with_mean,
                    'csv_file_with_mean': csv_file_with_mean
                }
            except Exception as e:
                print(f"❌ Erreur Ratios Matières: {e}")
                results['ratios_matieres'] = {'error': str(e)}
        

        return results

def main():
    """Fonction principale pour exécuter le mapping."""
    
    # Créer l'instance du mapper
    mapper = AnaplanMapper()
    
    # Chercher les fichiers automatiquement
    files_to_process = {
        'ppv_production': None,
        'ventes_mp': None,
        'ratios_matieres': None
    }
    
    # Rechercher les fichiers dans le répertoire courant
    for file in os.listdir('.'):
        if file.endswith('.xlsx'):
            if 'ppv_production' in file.lower():
                files_to_process['ppv_production'] = file
            elif 'ventes_mp' in file.lower() or ('ventes' in file.lower() and 'mp' in file.lower()):
                files_to_process['ventes_mp'] = file
            elif 'ratios' in file.lower() and 'matieres' in file.lower():
                files_to_process['ratios_matieres'] = file
    
    print("🔍 Fichiers détectés:")
    for file_type, filepath in files_to_process.items():
        status = f"✅ {filepath}" if filepath else "❌ Non trouvé"
        print(f"   {file_type.replace('_', ' ').title()}: {status}")
    
    # Traiter les fichiers
    if any(files_to_process.values()):
        results = mapper.process_excel_files(
            ppv_production_file=files_to_process['ppv_production'],
            ventes_mp_file=files_to_process['ventes_mp'],
            ratios_matieres_file=files_to_process['ratios_matieres']
        )
        return results
    else:
        print("⚠️ Aucun fichier à traiter trouvé!")
        print("📝 Placez les fichiers Excel dans le répertoire courant avec les noms:")
        print("   - fichier_ppv_production.xlsx")
        print("   - fichier_ventes_mp.xlsx") 
        print("   - fichier_ratios_matieres.xlsx")
        return None

if __name__ == "__main__":
    main() 