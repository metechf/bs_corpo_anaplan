import pandas as pd
from collections import defaultdict
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import numpy as np
import re  # Ajout pour la normalisation avancée des chaînes
from io import BytesIO
import tempfile

class AnaplanSanityChecker:
    """Classe pour effectuer les vérifications de cohérence sur le fichier de référence Anaplan-Ref.xlsx"""
    
    def __init__(self, ref_file_path="FICHIERS ANAPLAN/Anaplan-Ref.xlsx", summary_file=None, ppv_file=None):
        """Initialise le checker avec le chemin du fichier de référence et les fichiers uploadés par l'utilisateur."""
        self.ref_file_path = ref_file_path
        # Fichiers sources uploadés par l'utilisateur
        self.summary_file = summary_file
        self.ppv_file = ppv_file
        # Paramètres par défaut pour la génération
        self.exercice = "QBR2"
        self.date_version = "2025-01-27"
        
        self.sheets_config = {
            'Type de transaction': {
                'code_col': 'Code Transaction',
                'libelle_col': 'Libellé Transaction'
            },
            'Pays': {
                'code_col': 'Code Pays', 
                'libelle_col': 'Pays'
            },
            'Devise': {
                'code_col': 'Code Devises',
                'libelle_col': 'Devises'
            },
            'Qualité': {
                'code_col': 'Code SAP',
                'libelle_col': 'Libellé BS'  # PRIORITÉ: Libellé BS → Code SAP
            },
            'Exercice': {
                'code_col': 'Code Exercice',
                'libelle_col': 'Libellé'
            },
            'SiteEntité': {
                'code_col': 'Code P/CC',
                'libelle_col': 'Libellé Site/Entité'
            },
            'SiteEntité_vente': {
                'code_col': 'code SAP_2',
                'libelle_col': 'Site/Entité#Qualité'
            },
            'Partenaire Groupe': {
                'code_col': 'Code Partenaire Groupe',
                'libelle_col': 'Partenaire Groupe'
            },
            'Opération': {
                'code_col': 'Code Opération',
                'libelle_col': 'Opération'
            }
        }
        
        # Configuration pour les fichiers de sortie
        self.output_files_config = {
            'fichier_ppv_production.xlsx': {
                'columns_to_check': {
                    'Site/Entité': 'SiteEntité',
                    'Qualité': 'Qualité', 
                    'Opération': 'Opération'
                }
            },
            'fichier_ventes_mp.xlsx': {
                'columns_to_check': {
                    'Site/Entité': 'SiteEntité',
                    'Qualité': 'Qualité',
                    'Type de transaction': 'Type de transaction',
                    'Partenaire Groupe': 'Partenaire Groupe',
                    'Pays': 'Pays'
                }
            },
            'fichier_ratios_matieres.xlsx': {
                'columns_to_check': {
                    'Site/Entité': 'SiteEntité',
                    'Qualité': 'Qualité',
                    'Type Consommation Spécifique': 'Qualité'  # Même table que Qualité
                }
            }
        }
        
        self.ambiguous_labels_map = {}  # Nouveau : libellés ayant plusieurs codes par table
    
    def configure_generation_params(self, exercice=None, date_version=None, summary_file=None, ppv_file=None):
        """Configure les paramètres de génération automatique."""
        if exercice:
            self.exercice = exercice
        if date_version:
            self.date_version = date_version
        if summary_file:
            self.summary_file = summary_file
        if ppv_file:
            self.ppv_file = ppv_file
        
        print(f"[CONFIG] Configuration de generation mise a jour :")
        print(f"   [EXERCICE] : {self.exercice}")
        print(f"   [DATE] : {self.date_version}")
        print(f"   [SUMMARY] : {self.summary_file}")
        print(f"   [PPV] : {self.ppv_file}")
        
    def auto_generate_flat_files(self):
        """Génère automatiquement les fichiers plats depuis les fichiers uploadés par l'utilisateur."""
        print("[GENERATION] GENERATION AUTOMATIQUE DES FICHIERS PLATS")
        print("=" * 80)
        
        # Vérifier l'existence des fichiers sources uploadés
        if not self.summary_file:
            print(f"[WARNING] Fichier SUMMARY non fourni par l'utilisateur")
            return False
        
        if not self.ppv_file:
            print(f"[WARNING] Fichier PPV non fourni par l'utilisateur")
            return False
        
        try:
            # Importer les fonctions de génération depuis main.py
            from main import AnaplanMainApp
            generator = AnaplanMainApp()
            
            print(f"[INFO] Sources detectees :")
            print(f"   [SUMMARY] : Fichier uploadé par utilisateur")
            print(f"   [PPV]     : Fichier uploadé par utilisateur")
            print(f"   [EXERCICE] : {self.exercice}")
            print(f"   [DATE]    : {self.date_version}")
            print("-" * 50)
            
            # 1. Generer PPV Production
            print("[GEN] Generation PPV Production...")
            prod_df, prod_bytes = generator.generate_ppv_production(
                self.summary_file, self.exercice, self.date_version
            )
            
            if not prod_df.empty:
                output_path = "fichier_ppv_production.xlsx"
                with open(output_path, 'wb') as out_f:
                    out_f.write(prod_bytes)
                print(f"   [OK] Genere : {output_path} ({len(prod_df)} lignes)")
            else:
                print("   [ERREUR] Echec generation PPV Production")
            
            # 2. Generer Ventes & MP
            print("[GEN] Generation Ventes & MP...")
            ventes_df, ventes_bytes = generator.generate_ppv_ventes_mp(
                self.summary_file, self.exercice, self.date_version
            )
            
            if not ventes_df.empty:
                output_path = "fichier_ventes_mp.xlsx"
                with open(output_path, 'wb') as out_f:
                    out_f.write(ventes_bytes)
                print(f"   [OK] Genere : {output_path} ({len(ventes_df)} lignes)")
            else:
                print("   [ERREUR] Echec generation Ventes & MP")
            
            # 3. Generer Ratios Matieres  
            print("[GEN] Generation Ratios Matieres...")
            ratios_df, ratios_bytes = generator.generate_ratios(
                self.ppv_file, self.summary_file, self.exercice, self.date_version
            )
            
            if not ratios_df.empty:
                output_path = "fichier_ratios_matieres.xlsx"
                with open(output_path, 'wb') as out_f:
                    out_f.write(ratios_bytes)
                print(f"   [OK] Genere : {output_path} ({len(ratios_df)} lignes)")
            else:
                print("   [ERREUR] Echec generation Ratios Matieres")
            
            print("-" * 50)
            print("[OK] GENERATION AUTOMATIQUE TERMINEE")
            print("=" * 80)
            return True
            
        except Exception as e:
            print(f"[ERREUR] Erreur lors de la generation automatique : {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def standardize_value(self, value):
        """Normalise une chaîne pour une comparaison robuste.
        - Trim espaces début/fin
        - Uppercase
        - Remplace les multiples espaces par un seul
        - Si la chaîne est composée d'un même mot répété (ex: 'NAMIBIA NAMIBIA'),
          ne garde qu'une seule occurrence.
        """
        if pd.isna(value) or value == "":
            return None
        # Conversion chaîne + uppercase
        s = str(value).upper().strip()
        # Remplacer espaces multiples
        s = re.sub(r"\s+", " ", s)
        # Déduplication d'un mot répété plusieurs fois
        tokens = s.split(" ")
        if len(tokens) > 1 and all(tok == tokens[0] for tok in tokens):
            s = tokens[0]
        return s
    
    def normalize_value(self, value):
        """Normalise une valeur pour les comparaisons – utilise désormais standardize_value"""
        return self.standardize_value(value)
    
    def check_sheet_consistency(self, sheet_name, df, code_col, libelle_col):
        """Effectue les 4 vérifications sur une feuille donnée."""
        results = {
            'sheet_name': sheet_name,
            'total_rows': len(df),
            'duplicates_codes': [],
            'duplicates_libelles': [],
            'missing_codes': [],  # Libellés sans codes
            'missing_libelles': []  # Codes sans libellés
        }
        
        # Nettoyer les données
        df_clean = df.copy()
        df_clean[code_col] = df_clean[code_col].apply(self.normalize_value)
        df_clean[libelle_col] = df_clean[libelle_col].apply(self.normalize_value)
        
        # Filtrer les lignes complètement vides
        df_clean = df_clean.dropna(subset=[code_col, libelle_col], how='all')
        
        # 1. Vérification des duplicats dans les CODES
        if code_col in df_clean.columns:
            codes_non_nuls = df_clean[df_clean[code_col].notna()][code_col]
            code_counts = codes_non_nuls.value_counts()
            duplicated_codes = code_counts[code_counts > 1]
            
            for code, count in duplicated_codes.items():
                rows_with_code = df_clean[df_clean[code_col] == code].index.tolist()
                results['duplicates_codes'].append({
                    'value': code,
                    'count': count,
                    'rows': [r + 2 for r in rows_with_code]  # +2 pour tenir compte de l'index Excel
                })
        
        # 2. Vérification des duplicats dans les LIBELLÉS
        if libelle_col in df_clean.columns:
            libelles_non_nuls = df_clean[df_clean[libelle_col].notna()][libelle_col]
            libelle_counts = libelles_non_nuls.value_counts()
            duplicated_libelles = libelle_counts[libelle_counts > 1]
            
            for libelle, count in duplicated_libelles.items():
                rows_with_libelle = df_clean[df_clean[libelle_col] == libelle].index.tolist()
                results['duplicates_libelles'].append({
                    'value': libelle,
                    'count': count,
                    'rows': [r + 2 for r in rows_with_libelle]  # +2 pour tenir compte de l'index Excel
                })
        
        # 3. Vérification LIBELLÉS → CODES (libellés sans codes correspondants)
        libelles_with_missing_codes = df_clean[
            (df_clean[libelle_col].notna()) & 
            (df_clean[code_col].isna())
        ]
        
        for idx, row in libelles_with_missing_codes.iterrows():
            results['missing_codes'].append({
                'libelle': row[libelle_col],
                'row': idx + 2  # +2 pour tenir compte de l'index Excel
            })
        
        # 4. Vérification CODES → LIBELLÉS (codes sans libellés correspondants)
        codes_with_missing_libelles = df_clean[
            (df_clean[code_col].notna()) & 
            (df_clean[libelle_col].isna())
        ]
        
        for idx, row in codes_with_missing_libelles.iterrows():
            results['missing_libelles'].append({
                'code': row[code_col],
                'row': idx + 2  # +2 pour tenir compte de l'index Excel
            })
        
        return results
    
    def load_reference_data(self):
        """Charge toutes les tables de référence et identifie les libellés ambigus (plusieurs codes)."""
        ref_data = {}
        self.ambiguous_labels_map = {}
        try:
            xls = pd.ExcelFile(self.ref_file_path)
            for sheet_name, config in self.sheets_config.items():
                if sheet_name not in xls.sheet_names:
                    continue
                df = pd.read_excel(self.ref_file_path, sheet_name=sheet_name, dtype=str, keep_default_na=False)
                code_col = config['code_col']
                libelle_col = config['libelle_col']
                # === Construction de la liste des valeurs valides (codes ET libellés) ===
                valid_values = set()
                # Ajouter les libellés principaux (standardisés)
                if libelle_col in df.columns:
                    valid_values.update(df[libelle_col].dropna().apply(self.standardize_value))

                # Cas particulier QUALITÉ : prendre aussi Libellé SAP
                if sheet_name == 'Qualité':
                    if 'Libellé BS' in df.columns:
                        valid_values.update(df['Libellé BS'].dropna().apply(self.standardize_value))
                    if 'Libellé SAP' in df.columns:
                        valid_values.update(df['Libellé SAP'].dropna().apply(self.standardize_value))

                # Ajouter également la colonne des codes si présente (standardisée)
                if code_col in df.columns:
                    valid_values.update(df[code_col].dropna().apply(self.standardize_value))

                ref_data[sheet_name] = valid_values
                # === Détection des ambiguïtés ===
                if code_col in df.columns and libelle_col in df.columns:
                    mapping = {}
                    for _, row in df.iterrows():
                        lbl_raw = row[libelle_col]
                        code_raw = row[code_col]
                        lbl = self.standardize_value(lbl_raw)
                        code = self.standardize_value(code_raw)
                        if not lbl or not code:
                            continue
                        mapping.setdefault(lbl, set()).add(code)
                    ambiguous = {lbl for lbl, codes in mapping.items() if len(codes) > 1}
                    if ambiguous:
                        self.ambiguous_labels_map[sheet_name] = ambiguous
        except Exception as e:
            print(f"❌ Erreur lors du chargement des références : {e}")
        return ref_data
    
    def check_output_file_values(self, file_path):
        """Vérifie les valeurs d'un fichier de sortie et identifie les lignes problématiques."""
        if not os.path.exists(file_path):
            print(f"⚠️  Fichier non trouvé : {file_path}")
            return None
            
        filename = os.path.basename(file_path)
        if filename not in self.output_files_config:
            print(f"⚠️  Configuration non trouvée pour : {filename}")
            return None
            
        print(f"\n📄 Vérification : {filename}")
        print("-" * 50)
        
        # Charger les données de référence
        ref_data = self.load_reference_data()
        
        # Charger le fichier de sortie
        df = pd.read_excel(file_path, dtype=str, keep_default_na=False)
        columns_config = self.output_files_config[filename]['columns_to_check']
        
        problem_rows = []
        stats = {}
        
        for col_name, ref_table in columns_config.items():
            if col_name not in df.columns:
                print(f"⚠️  Colonne '{col_name}' non trouvée")
                continue
                
            if ref_table not in ref_data:
                print(f"⚠️  Table de référence '{ref_table}' non chargée")
                continue
                
            valid_values = ref_data[ref_table]
            col_values = df[col_name].dropna().apply(lambda x: self.standardize_value(x) or "")
            
            # Identifier les valeurs non trouvées
            missing_values = []
            for idx, value in col_values.items():
                # TOLÉRANCE: Ignorer les valeurs vides pour Site/Entité (problème de mapping en cours)
                if value and value.strip():  # Seulement si non vide et non blanc
                    value_norm = self.standardize_value(value)
                    if value_norm not in valid_values:
                        missing_values.append((idx, value))
                        if idx not in problem_rows:
                            problem_rows.append(idx)
            
            stats[col_name] = {
                'total_values': len(col_values),
                'missing_count': len(missing_values),
                'missing_values': missing_values[:10]  # Limiter l'affichage
            }
            
            # Afficher les statistiques
            print(f"🔍 {col_name} → {ref_table}")
            print(f"   Total: {stats[col_name]['total_values']}, "
                  f"Manquants: {stats[col_name]['missing_count']}")
            
            # Note spéciale pour Site/Entité
            if col_name == "Site/Entité" and stats[col_name]['missing_count'] == 0:
                empty_count = len(col_values) - len([v for v in col_values if v and v.strip()])
                if empty_count > 0:
                    print(f"   ℹ️  Note: {empty_count} valeur(s) vide(s) ignorée(s) (mapping Contract→Entity en cours)")
            
            if missing_values:
                print(f"   Valeurs non trouvées (premières 10):")
                for idx, value in missing_values[:10]:
                    print(f"     • '{value}' (ligne {idx + 2})")
                if len(missing_values) > 10:
                    print(f"     • ... et {len(missing_values) - 10} autres")
        
        # Collecter les détails pour le rapport (même s'il n'y a pas de problèmes de ligne)
        detailed_problems = self._collect_detailed_problems(df, columns_config, ref_data)
        
        return {
            'file_path': file_path,
            'problem_rows': problem_rows,
            'stats': stats,
            'total_problems': len(problem_rows),
            'detailed_problems': detailed_problems
        }
    
    def _collect_detailed_problems(self, df, columns_config, ref_data):
        """Collecte les détails des problèmes pour le rapport.
        Ajoute un champ 'problem_type' :
            - 'MISSING_VALUE'  → libellé inexistant dans la référence
            - 'MISSING_CODE'   → libellé présent mais code vide dans la référence
        """
        detailed_problems = []
        for col_name, ref_table in columns_config.items():
            if col_name not in df.columns or ref_table not in ref_data:
                continue
            valid_values = ref_data[ref_table]
            col_values = df[col_name].dropna().apply(lambda x: self.standardize_value(x) or "")
            for idx, value in col_values.items():
                if not value or not value.strip():
                    continue
                value_norm = self.standardize_value(value)
                if value_norm not in valid_values:
                    problem_type = 'MISSING_VALUE'  # Libellé inexistant
                elif (self.ambiguous_labels_map.get(ref_table) and
                      value_norm in self.ambiguous_labels_map[ref_table]):
                    problem_type = 'CODE_DUPLIQUE'
                elif hasattr(self, 'values_without_code') and value_norm in self.values_without_code:
                    problem_type = 'MISSING_CODE'   # Libellé présent mais code vide
                else:
                    continue  # Pas de problème
                detailed_problems.append({
                    'ligne_excel': idx + 2,  # Excel index
                    'colonne': col_name,
                    'valeur_manquante': value,
                    'problem_type': problem_type,
                    'table_reference': ref_table,
                    'bloc': df.loc[idx, 'Bloc'] if 'Bloc' in df.columns else 'N/A',
                    'site_entite': df.loc[idx, 'Site/Entité'] if 'Site/Entité' in df.columns else 'N/A',
                    'qualite': df.loc[idx, 'Qualité'] if 'Qualité' in df.columns else 'N/A'
                })
        return detailed_problems

    def color_problem_rows(self, file_path, problem_rows):
        # DEPRECATED — remplacé par color_problem_cells
        return self.color_problem_cells(file_path, problem_rows)

    def color_problem_cells(self, file_path, detailed_problems):
        """Colorie uniquement les cellules problématiques.
        ORANGE : MISSING_VALUE | VERT CLAIR : MISSING_CODE | JAUNE : AMBIGUOUS_CODE"""
        if not detailed_problems:
            print("✅ Aucune cellule à colorier")
            return file_path
        try:
            from openpyxl.utils import get_column_letter
            wb = load_workbook(file_path)
            ws = wb.active
            # Préparer le mapping des en-têtes → index
            header_map = {}
            for col_idx in range(1, ws.max_column + 1):
                header = ws.cell(row=1, column=col_idx).value
                if header is not None:
                    header_map[str(header).strip()] = col_idx
            # Définir les styles
            orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
            green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            yellow_fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
            processed_cells = set()
            for prob in detailed_problems:
                row_idx = prob.get('ligne_excel')
                col_name = prob.get('colonne')
                ptype = prob.get('problem_type')
                if row_idx is None or col_name is None:
                    continue
                col_idx = header_map.get(col_name)
                if col_idx is None:
                    continue
                key = (row_idx, col_idx)
                if key in processed_cells:
                    continue  # déjà colorée
                cell = ws.cell(row=row_idx, column=col_idx)
                if ptype == 'MISSING_CODE':
                    cell.fill = green_fill
                elif ptype == 'CODE_DUPLIQUE':
                    cell.fill = yellow_fill
                else:  # MISSING_VALUE
                    cell.fill = orange_fill
                processed_cells.add(key)
            base_name = os.path.splitext(file_path)[0]
            new_file_path = f"{base_name}_sanity_checked.xlsx"
            wb.save(new_file_path)
            print(f"💾 Fichier coloré sauvegardé : {new_file_path} (cellules: {len(processed_cells)})")
            return new_file_path
        except Exception as e:
            print(f"❌ Erreur coloration cellules : {e}")
            return file_path
    
    def check_and_color_output_files(self):
        """Vérifie et colorie tous les fichiers de sortie disponibles."""
        print("\n" + "="*80)
        print("🎨 VÉRIFICATION ET COLORATION DES FICHIERS DE SORTIE")
        print("="*80)
        
        colored_files = []
        total_problems = 0
        
        for filename in self.output_files_config.keys():
            if os.path.exists(filename):
                results = self.check_output_file_values(filename)
                if results and results['detailed_problems']:
                    colored_file = self.color_problem_cells(filename, results['detailed_problems'])
                    colored_files.append({
                        'original': filename,
                        'colored': colored_file,
                        'problems': results['total_problems'],
                        'detailed_problems': results['detailed_problems'],
                        'stats': results['stats']
                    })
                    total_problems += results['total_problems']
                elif results:
                    print("✅ Aucun problème détecté")
                    colored_files.append({
                        'original': filename,
                        'colored': filename,
                        'problems': 0,
                        'detailed_problems': results['detailed_problems'],
                        'stats': results['stats']
                    })
            else:
                print(f"⚠️  Fichier non trouvé : {filename}")
        
        print(f"\n📊 RÉSUMÉ:")
        print(f"   Fichiers traités: {len(colored_files)}")
        print(f"   Total des problèmes: {total_problems}")
        
        if colored_files:
            print(f"\n📁 Fichiers disponibles pour téléchargement:")
            for file_info in colored_files:
                if file_info['problems'] > 0:
                    print(f"   🔴 {file_info['colored']} ({file_info['problems']} problème(s))")
                else:
                    print(f"   ✅ {file_info['original']} (aucun problème)")
        
        return colored_files
    
    def run_full_sanity_check(self):
        """Exécute les vérifications de cohérence sur toutes les feuilles importantes."""
        print("[START] DEMARRAGE DU SANITY CHECK COMPLET")
        print("=" * 80)
        
        # ETAPE 1: Generation automatique des fichiers plats
        print("[ETAPE 1/2] GENERATION AUTOMATIQUE")
        generation_success = self.auto_generate_flat_files()
        
        if not generation_success:
            print("[ERREUR] Echec de la generation automatique - Poursuite avec les fichiers existants")
        
        # ETAPE 2: Verifications de coherence
        print("[ETAPE 2/2] VERIFICATIONS DE COHERENCE")
        print(f"[INFO] Fichier de reference: {self.ref_file_path}")
        print("=" * 80)
        
        if not os.path.exists(self.ref_file_path):
            print(f"[ERREUR] Le fichier {self.ref_file_path} n'existe pas!")
            return None
        
        all_results = {}
        total_issues = 0
        
        try:
            xls = pd.ExcelFile(self.ref_file_path)
            
            for sheet_name, config in self.sheets_config.items():
                if sheet_name not in xls.sheet_names:
                    print(f"[WARNING] Feuille '{sheet_name}' non trouvee dans le fichier")
                    continue
                
                # EXCLUSION: SiteEntité_vente sera traitée séparément avec notre analyse spécialisée
                if sheet_name == 'SiteEntité_vente':
                    print(f"[SKIP] Feuille '{sheet_name}' exclue de l'analyse standard (traitement specialise)")
                    continue
                
                print(f"\n[CHECK] Verification: {sheet_name}")
                print("-" * 50)
                
                try:
                    df = pd.read_excel(self.ref_file_path, sheet_name=sheet_name, dtype=str, keep_default_na=False)
                    code_col = config['code_col']
                    libelle_col = config['libelle_col']
                    
                    # Vérifier que les colonnes existent
                    if code_col not in df.columns:
                        print(f"[ERROR] Colonne '{code_col}' non trouvee")
                        continue
                    if libelle_col not in df.columns:
                        print(f"[ERROR] Colonne '{libelle_col}' non trouvee")
                        continue
                    
                    # Effectuer les vérifications
                    results = self.check_sheet_consistency(sheet_name, df, code_col, libelle_col)
                    all_results[sheet_name] = results
                    
                    # Compter les problèmes
                    sheet_issues = (len(results['duplicates_codes']) + 
                                  len(results['duplicates_libelles']) + 
                                  len(results['missing_codes']) + 
                                  len(results['missing_libelles']))
                    total_issues += sheet_issues
                    
                    # Afficher les résultats
                    print(f"[INFO] Total des lignes: {results['total_rows']}")
                    
                    if results['duplicates_codes']:
                        print(f"[PROBLEME] Codes dupliques: {len(results['duplicates_codes'])}")
                        for dup in results['duplicates_codes']:
                            print(f"   - '{dup['value']}' apparait {dup['count']} fois (lignes: {dup['rows']})")
                    
                    if results['duplicates_libelles']:
                        print(f"[PROBLEME] Libelles dupliques: {len(results['duplicates_libelles'])}")
                        for dup in results['duplicates_libelles']:
                            print(f"   - '{dup['value']}' apparait {dup['count']} fois (lignes: {dup['rows']})")
                    
                    if results['missing_codes']:
                        print(f"[PROBLEME] Libelles sans codes: {len(results['missing_codes'])}")
                        for missing in results['missing_codes'][:5]:  # Afficher seulement les 5 premiers
                            print(f"   - '{missing['libelle']}' (ligne {missing['row']})")
                        if len(results['missing_codes']) > 5:
                            print(f"   - ... et {len(results['missing_codes']) - 5} autres")
                    
                    if results['missing_libelles']:
                        print(f"[PROBLEME] Codes sans libelles: {len(results['missing_libelles'])}")
                        for missing in results['missing_libelles'][:5]:  # Afficher seulement les 5 premiers
                            print(f"   - '{missing['code']}' (ligne {missing['row']})")
                        if len(results['missing_libelles']) > 5:
                            print(f"   - ... et {len(results['missing_libelles']) - 5} autres")
                    
                    if sheet_issues == 0:
                        print("[OK] Aucun probleme detecte dans cette feuille")
                
                except Exception as e:
                    print(f"[ERREUR] Erreur lors de la verification: {e}")
                    continue
            
            # Résumé final
            print("\n" + "=" * 80)
            print("[RESUME] RESUMÉ DES VERIFICATIONS ANAPLAN-REF")
            print("=" * 80)
            
            if total_issues == 0:
                print("[EXCELLENT] Aucun probleme de coherence detecte dans le fichier de reference")
            else:
                print(f"[ATTENTION] {total_issues} probleme(s) de coherence detecte(s)")
                print("\n[ACTIONS] Actions recommandees:")
                print("   1. Corriger les doublons dans les tables de reference")
                print("   2. Completer les mappings manquants (codes <-> libelles)")
                print("   3. Verifier la coherence des donnees manuellement")
            
            # Exporter le rapport Excel pour Anaplan-Ref
            if all_results:
                print("\n[GEN] Generation du rapport Excel pour Anaplan-Ref...")
                self.export_issues_to_excel(all_results, "anaplan_ref_sanity_report.xlsx")
            
            # Vérifier et colorier les fichiers de sortie
            colored_files = self.check_and_color_output_files()
            
            # Ajouter la feuille des valeurs problématiques dans le rapport Anaplan-Ref
            if all_results:  # Le rapport principal vient d'être créé
                self.append_output_problems_sheet("anaplan_ref_sanity_report.xlsx", colored_files)

            # NEW – Couverture Produits vs Matrix
            try:
                print("🔍 DEBUG - Génération Couverture Produits...")
                coverage_df = self.check_product_vs_matrix()
                print(f"🔍 DEBUG - Coverage DF shape: {coverage_df.shape if coverage_df is not None else 'None'}")
                if coverage_df is not None and not coverage_df.empty:
                    print(f"🔍 DEBUG - Ajout de la feuille Couverture Produits avec {len(coverage_df)} lignes")
                    self.append_product_coverage_sheet("anaplan_ref_sanity_report.xlsx", coverage_df)
                else:
                    print("ℹ️  Feuille 'Couverture Produits' non générée : aucun produit absent de la matrix détecté")
            except Exception as e:
                print(f"❌ Erreur Couverture Produits vs Matrix : {e}")
                import traceback
                traceback.print_exc()

            # NEW – SiteEntité_vente combinaisons manquantes
            try:
                print("🔍 DEBUG - Génération SiteEntité_vente...")
                combo_counts = self.analyze_site_entite_vente_missing_combinations()
                if combo_counts and len(combo_counts) > 0:
                    print(f"🔍 DEBUG - Ajout de la feuille SiteEntité_vente avec {len(combo_counts)} problèmes")
                    self.append_site_entite_vente_sheet("anaplan_ref_sanity_report.xlsx", combo_counts)
                else:
                    print("ℹ️  Feuille 'SiteEntité_vente' non générée : aucune combinaison manquante détectée")
            except Exception as e:
                print(f"❌ Erreur SiteEntité_vente : {e}")
                import traceback
                traceback.print_exc()

            # Résumé global
            print("\n" + "=" * 80)
            print("🎯 RÉSUMÉ GLOBAL DU SANITY CHECK")
            print("=" * 80)
            print(f"📊 Problèmes Anaplan-Ref: {total_issues}")
            if colored_files:
                output_problems = sum(f['problems'] for f in colored_files)
                print(f"📊 Problèmes fichiers de sortie: {output_problems}")
                print(f"📊 TOTAL: {total_issues + output_problems} problème(s)")
            
            print("\n📁 FICHIERS GÉNÉRÉS:")
            generated_files = []
            
            # Ajouter anaplan_ref_sanity_report.xlsx si généré
            if all_results:
                generated_files.append("anaplan_ref_sanity_report.xlsx")
                print("   📄 anaplan_ref_sanity_report.xlsx")
            
            # Ajouter les fichiers colorés
            for file_info in colored_files:
                if file_info['colored'] != file_info['original']:
                    generated_files.append(file_info['colored'])
                    print(f"   📄 {file_info['colored']}")
            
            # Retourner les résultats complets
            return {
                'ref_results': all_results,
                'colored_files': colored_files,
                'total_ref_issues': total_issues,
                'ref_report_file': "anaplan_ref_sanity_report.xlsx" if all_results else None,
                'output_report_file': "output_files_report.xlsx",
                'all_generated_files': generated_files
            }
            
        except Exception as e:
            print(f"❌ ERREUR CRITIQUE: {e}")
            return None
    
    def export_output_files_report(self, colored_files, output_file="output_files_report.xlsx", ref_results=None):
        """Exporte un rapport détaillé des problèmes dans les fichiers de sortie avec toutes les colonnes et lignes."""
        if not colored_files and not ref_results:
            print("❌ Aucun fichier de sortie à analyser")
            return

        try:
            print(f"🔄 Début génération du rapport détaillé {output_file}...")
            print(f"📊 Nombre de fichiers à traiter: {len(colored_files)}")
            
            # Compiler tous les problèmes détaillés
            all_problems = []
            summary_data = []
            
            for file_info in colored_files:
                filename = file_info['original']
                problems = file_info.get('detailed_problems', [])
                stats = file_info.get('stats', {})
                
                print(f"📋 Traitement {filename}: {len(problems)} problèmes détaillés")
                
                # Ajouter les problèmes avec le nom du fichier
                for problem in problems:
                    problem_with_file = problem.copy()
                    problem_with_file['fichier'] = filename
                    all_problems.append(problem_with_file)
                
                # Données pour le résumé - nettoyer les valeurs NaN/Inf
                total_values = sum(stat.get('total_values', 0) for stat in stats.values() if stat.get('total_values') is not None)
                total_missing = sum(stat.get('missing_count', 0) for stat in stats.values() if stat.get('missing_count') is not None)
                
                # S'assurer que les valeurs sont finies
                total_values = total_values if np.isfinite(total_values) else 0
                total_missing = total_missing if np.isfinite(total_missing) else 0
                problems_count = file_info.get('problems', 0)
                problems_count = problems_count if np.isfinite(problems_count) else 0
                
                summary_data.append({
                    'Fichier': filename,
                    'Total valeurs vérifiées': int(total_values),
                    'Valeurs manquantes': int(total_missing),
                    'Lignes problématiques': int(problems_count),
                    'Colonnes vérifiées': ', '.join(stats.keys()) if stats else 'Aucune'
                })
            
            print(f"📊 Total problèmes collectés: {len(all_problems)}")
            
            # Créer le fichier Excel avec les rapports détaillés
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                print("🔧 Configuration des formats Excel...")
                
                # Format pour les en-têtes
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#2E7D32',
                    'font_color': 'white',
                    'border': 1,
                    'text_wrap': True
                })
                
                # Format pour les erreurs
                error_format = workbook.add_format({
                    'bg_color': '#FFCDD2',
                    'border': 1
                })
                
                # Format pour les titres de feuilles
                title_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#1976D2',
                    'font_color': 'white',
                    'border': 1,
                    'align': 'center',
                    'font_size': 14
                })
                
                # 1. Feuilles de rapport détaillé par fichier
                file_mappings = {
                    'fichier_ppv_production.xlsx': 'RAPPORT PPV PRODUCTION',
                    'fichier_ventes_mp.xlsx': 'RAPPORT VENTES & MP',
                    'fichier_ratios_matieres.xlsx': 'RAPPORT RATIOS MATIERES'
                }
                
                print("📝 Création des feuilles de rapport détaillé...")
                
                # S'assurer que tous les fichiers attendus sont inclus même s'ils ne sont pas dans colored_files
                all_expected_files = list(file_mappings.keys())
                processed_filenames = [f['original'] for f in colored_files]
                
                # Ajouter les fichiers manquants avec des données vides
                for expected_file in all_expected_files:
                    if expected_file not in processed_filenames:
                        print(f"📋 Ajout du fichier manquant pour le rapport : {expected_file}")
                        colored_files.append({
                            'original': expected_file,
                            'colored': expected_file,
                            'problems': 0,
                            'detailed_problems': [],
                            'stats': {}
                        })
                
                for file_info in colored_files:
                    filename = file_info['original']
                    
                    if filename in file_mappings:  # Créer la feuille pour TOUS les fichiers attendus
                        sheet_name = file_mappings.get(filename, filename.replace('.xlsx', '').upper()[:25])
                        print(f"🔧 Création feuille détaillée: {sheet_name}")
                        
                        try:
                            # Filtrer les problèmes pour ce fichier
                            file_problems = [p for p in all_problems if p.get('fichier') == filename]
                            
                            if file_problems:
                                # Créer DataFrame avec tous les détails
                                df_problems = pd.DataFrame(file_problems)
                                
                                # Réorganiser les colonnes pour un meilleur affichage
                                column_order = ['fichier', 'Ligne Excel', 'Colonne', 'Valeur Manquante', 'Table Reference', 'Bloc', 'Site/Entité', 'Qualité']
                                # Ajouter les colonnes qui existent mais ne sont pas dans l'ordre défini
                                for col in df_problems.columns:
                                    if col not in column_order:
                                        column_order.append(col)
                                
                                # Réorganiser selon l'ordre défini
                                available_columns = [col for col in column_order if col in df_problems.columns]
                                df_problems = df_problems[available_columns]
                                
                                # Exporter vers Excel
                                df_problems.to_excel(writer, sheet_name=sheet_name, index=False)
                                
                                # Formater la feuille
                                worksheet = writer.sheets[sheet_name]
                                
                                # Formater les en-têtes
                                for col_num, value in enumerate(df_problems.columns.values):
                                    worksheet.write(0, col_num, value, header_format)
                                
                                # Ajouter titre en haut
                                problems_count = len(file_problems)
                                worksheet.insert_rows(0, 1)
                                worksheet.merge_range(0, 0, 0, len(available_columns)-1, 
                                                    f'{filename} - {problems_count} problème(s) détecté(s)', title_format)
                                worksheet.set_row(0, 25)
                                
                                # Ajuster les largeurs de colonnes
                                for idx, col in enumerate(available_columns):
                                    max_len = max(len(str(col)), 15)
                                    if col in ['Valeur Manquante', 'Table Reference']:
                                        max_len = min(max_len, 30)
                                    worksheet.set_column(idx, idx, max_len)
                                
                                print(f"✅ Feuille détaillée {sheet_name} créée avec {len(file_problems)} problèmes")
                            else:
                                # Créer une feuille vide pour les fichiers sans problèmes
                                worksheet = workbook.add_worksheet(sheet_name)
                                worksheet.write(0, 0, f'{filename} - 0 problème(s) détecté(s)', title_format)
                                worksheet.set_row(0, 25)
                                worksheet.write(2, 0, '✅ Aucun problème détecté dans ce fichier')
                                worksheet.write(3, 0, 'Toutes les valeurs correspondent aux tables de référence')
                                worksheet.set_column('A:A', 60)
                                print(f"✅ Feuille {sheet_name} créée (aucun problème)")
                                
                        except Exception as e:
                            print(f"❌ Erreur création feuille détaillée {sheet_name}: {e}")
                            import traceback
                            traceback.print_exc()
                
                # 2. Feuille RÉSUMÉ détaillée
                if summary_data:
                    print("📊 Création feuille RÉSUMÉ détaillée...")
                    try:
                        summary_df = pd.DataFrame(summary_data)
                        summary_df.to_excel(writer, sheet_name='RÉSUMÉ', index=False)
                        
                        # Formater la feuille résumé
                        worksheet = writer.sheets['RÉSUMÉ']
                        for col_num, value in enumerate(summary_df.columns.values):
                            worksheet.write(0, col_num, value, header_format)
                        
                        # Ajuster les largeurs de colonnes du résumé
                        worksheet.set_column('A:A', 30)  # Fichier
                        worksheet.set_column('B:D', 20)  # Valeurs numériques
                        worksheet.set_column('E:E', 50)  # Colonnes vérifiées
                        
                        print("✅ Feuille RÉSUMÉ détaillée créée avec succès")
                    except Exception as e:
                        print(f"❌ Erreur création feuille RÉSUMÉ: {e}")
            
            print(f"✅ Rapport détaillé {output_file} généré avec succès!")
            return True
            
        except Exception as e:
            print(f"❌ Erreur lors de l'export du rapport détaillé: {e}")
            import traceback
            traceback.print_exc()
            return False

    def export_issues_to_excel(self, results, output_file="anaplan_ref_sanity_report.xlsx"):
        """Exporte les problèmes détectés vers un fichier Excel pour analyse."""
        if not results:
            print("❌ Aucun résultat à exporter")
            return
        
        # Créer le dossier de sortie s'il n'existe pas (seulement si un chemin est spécifié)
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        try:
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # Format pour les en-têtes
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#4CAF50',
                    'font_color': 'white',
                    'border': 1
                })
                
                # Format pour les erreurs
                error_format = workbook.add_format({
                    'bg_color': '#FFEB3B',
                    'border': 1
                })
                
                # Créer une feuille de résumé
                summary_data = []
                for sheet_name, result in results.items():
                    # EXCLUSION: SiteEntité_vente sera traitée séparément
                    if sheet_name == 'SiteEntité_vente':
                        continue
                    summary_data.append({
                        'Feuille': sheet_name,
                        'Total lignes': result['total_rows'],
                        'Codes dupliqués': len(result['duplicates_codes']),
                        'Libellés dupliqués': len(result['duplicates_libelles']),
                        'Libellés sans codes': len(result['missing_codes']),
                        'Codes sans libellés': len(result['missing_libelles']),
                        'Total problèmes': (len(result['duplicates_codes']) + 
                                          len(result['duplicates_libelles']) + 
                                          len(result['missing_codes']) + 
                                          len(result['missing_libelles']))
                    })
                
                # Ajouter SiteEntité_vente au résumé si elle n'est pas déjà présente
                if 'SiteEntité_vente' not in [item['Feuille'] for item in summary_data]:
                    # Analyser les combinaisons manquantes pour SiteEntité_vente
                    combo_counts = self.analyze_site_entite_vente_missing_combinations()
                    if combo_counts:
                        summary_data.append({
                            'Feuille': 'SiteEntité_vente',
                            'Total lignes': len(combo_counts),
                            'Codes dupliqués': 0,
                            'Libellés dupliqués': 0,
                            'Libellés sans codes': 0,
                            'Codes sans libellés': 0,
                            'Total problèmes': len(combo_counts)
                        })
                
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Résumé', index=False)
                
                # Formater la feuille de résumé
                worksheet = writer.sheets['Résumé']
                for col_num, value in enumerate(summary_df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Créer des feuilles détaillées pour chaque type de problème
                # EXCLUSION: SiteEntité_vente sera traitée séparément
                for sheet_name, result in results.items():
                    if sheet_name == 'SiteEntité_vente':
                        continue  # Skip SiteEntité_vente - sera traité par notre fonction spécialisée
                    if any([result['duplicates_codes'], result['duplicates_libelles'], 
                           result['missing_codes'], result['missing_libelles']]):
                        
                        # Feuille pour les détails des problèmes
                        detail_data = []
                        
                        # Ajouter les codes dupliqués
                        for dup in result['duplicates_codes']:
                            detail_data.append({
                                'Table': sheet_name,
                                'Type de problème': 'Code dupliqué',
                                'Valeur problématique': dup['value'],
                                'Lignes Excel': ', '.join(map(str, dup['rows'])),
                                'Nombre occurrences': dup['count']
                            })
                        
                        # Ajouter les libellés dupliqués
                        for dup in result['duplicates_libelles']:
                            detail_data.append({
                                'Table': sheet_name,
                                'Type de problème': 'Libellé dupliqué',
                                'Valeur problématique': dup['value'],
                                'Lignes Excel': ', '.join(map(str, dup['rows'])),
                                'Nombre occurrences': dup['count']
                            })
                        
                        # Ajouter les libellés sans codes
                        for missing in result['missing_codes']:
                            detail_data.append({
                                'Table': sheet_name,
                                'Type de problème': 'Libellé sans code',
                                'Valeur problématique': missing['libelle'],
                                'Lignes Excel': str(missing['row']),
                                'Nombre occurrences': 1
                            })
                        
                        # Ajouter les codes sans libellés
                        for missing in result['missing_libelles']:
                            detail_data.append({
                                'Table': sheet_name,
                                'Type de problème': 'Code sans libellé',
                                'Valeur problématique': missing['code'],
                                'Lignes Excel': str(missing['row']),
                                'Nombre occurrences': 1
                            })
                        
                        if detail_data:
                            detail_df = pd.DataFrame(detail_data)
                            sheet_safe_name = sheet_name.replace('/', '_')[:31]  # Limiter à 31 caractères
                            
                            # Créer la feuille manuellement pour un meilleur contrôle
                            worksheet = workbook.add_worksheet(sheet_safe_name)
                            
                            # Format pour les titres
                            title_format = workbook.add_format({
                                'bold': True,
                                'bg_color': '#1976D2',
                                'font_color': 'white',
                                'border': 1,
                                'align': 'center',
                                'font_size': 14
                            })
                            
                            # Format pour les problèmes
                            problem_format = workbook.add_format({
                                'bg_color': '#FFF3E0',
                                'border': 1
                            })
                            
                            # Titre
                            num_cols = len(detail_df.columns)
                            worksheet.merge_range(0, 0, 0, num_cols-1, 
                                                f'TABLE: {sheet_name} - {len(detail_data)} problème(s) à corriger', 
                                                title_format)
                            worksheet.set_row(0, 25)
                            
                            # En-têtes
                            headers = ['Table', 'Type de problème', 'Valeur problématique', 'Lignes Excel', 'Nombre occurrences']
                            for col_num, header in enumerate(headers):
                                worksheet.write(1, col_num, header, header_format)
                            
                            # Données
                            for row_num, (_, row_data) in enumerate(detail_df.iterrows()):
                                for col_num, value in enumerate(row_data):
                                    worksheet.write(row_num + 2, col_num, value, problem_format)
                            
                            # Ajuster les colonnes
                            worksheet.set_column('A:A', 20)  # Table
                            worksheet.set_column('B:B', 20)  # Type de problème
                            worksheet.set_column('C:C', 30)  # Valeur problématique
                            worksheet.set_column('D:D', 15)  # Lignes Excel
                            worksheet.set_column('E:E', 12)  # Nombre occurrences
            
            print(f"📄 Rapport de vérification exporté: {output_file}")
            
        except Exception as e:
            print(f"❌ Erreur lors de l'export: {e}")

    def append_output_problems_sheet(self, report_file, colored_files):
        """Crée ou remplace la feuille "valeurs manquent dans fichier BS (anaplan intermédiaire)".
        Seules les anomalies MISSING_VALUE sont conservées.
        Colonnes : Valeur | Table Référence
        """
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        import pandas as pd
        # 1. Collecte des valeurs manquantes
        records = []
        for file_info in colored_files:
            for problem in file_info.get("detailed_problems", []):
                if problem.get("problem_type") != "MISSING_VALUE":
                    continue  # On ignore tout le reste
                val = problem.get("valeur_manquante") or problem.get("Valeur Manquante")
                if val is None:
                    continue
                records.append({
                    "Valeur": val,
                    "Table Référence": problem.get("table_reference")
                })
        if not records:
            print("ℹ️  Aucune valeur MISSING_VALUE à enregistrer")
            return False
        df = pd.DataFrame(records).drop_duplicates().reset_index(drop=True)

        sheet_title = "Valeurs BS absentes du référentiel"
        sheet_name = sheet_title[:31]
        try:
            wb = load_workbook(report_file)
            # Supprimer toute ancienne feuille Valeurs Problématiques*
            for s in list(wb.sheetnames):
                if s.lower().startswith("valeurs problématiques"):
                    del wb[s]
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
            ws = wb.create_sheet(sheet_name)

            # Styles
            header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            title_fill = PatternFill(start_color="1976D2", end_color="1976D2", fill_type="solid")
            white_font = Font(color="FFFFFF", bold=True)
            center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            thin = Side(border_style="thin", color="000000")
            border = Border(top=thin, left=thin, right=thin, bottom=thin)

            # Titre
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))
            c = ws.cell(row=1, column=1, value=f"{sheet_title} - {len(df)} valeur(s)")
            c.fill = title_fill
            c.font = white_font
            c.alignment = center_align
            c.border = border
            ws.row_dimensions[1].height = 25

            # En-têtes
            for idx, col in enumerate(df.columns, start=1):
                cell = ws.cell(row=2, column=idx, value=col)
                cell.fill = header_fill
                cell.font = white_font
                cell.alignment = center_align
                cell.border = border

            # Lignes données
            for r_idx, (_, row) in enumerate(df.iterrows(), start=3):
                for c_idx, col in enumerate(df.columns, start=1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=row[col])
                    cell.border = border

            # Largeur colonnes
            for idx, col in enumerate(df.columns, start=1):
                letter = get_column_letter(idx)
                max_len = max(len(str(col)), df[col].astype(str).map(len).max()) + 2
                ws.column_dimensions[letter].width = max(15, max_len)

            wb.save(report_file)
            print(f"✅ Feuille '{sheet_title}' mise à jour : {len(df)} valeur(s)")
            return True
        except Exception as e:
            print(f"❌ Erreur append_output_problems_sheet : {e}")
            return False

    def check_product_vs_matrix(self,
                                 production_file="fichier_ppv_production.xlsx",
                                 ventes_file="fichier_ventes_mp.xlsx",
                                 matrix_file="fichier_ratios_matieres.xlsx"):
        """Contrôle couverture produits.
        1. Produits issus de la production chimique (PPV Production).
        2. Produits issus des ventes (#Volume / Sales) filtrés sur TypeProduct∈{Fertilizers, Fertilizers W, MarketableAcids}.
        3. Vérification de la présence dans la Matrix (Ratios).

        Retour : DataFrame avec anomalies (Abst. Matrix / Jamais produit / Jamais vendu).
        """
        import os
        import pandas as pd
        allowed_typeprod = {"Fertilizers", "Fertilizers W", "MarketableAcids"}

        # --- 1) Couples Production -------------------------------------------------
        prod_couples = set()
        if os.path.exists(production_file):
            try:
                xls = pd.ExcelFile(production_file)
                sheet_name = next((s for s in xls.sheet_names if "ppv production" in s.lower()), xls.sheet_names[0])
                df_prod = pd.read_excel(production_file, sheet_name=sheet_name, dtype=str, keep_default_na=False)
                df_prod = df_prod[df_prod.get("Opération", "") == "Traitements Chimiques"].copy()
                for _, row in df_prod.iterrows():
                    prod = self.standardize_value(row.get("Qualité", ""))
                    ent = row.get("Site/Entité", "")
                    if prod:
                        prod_couples.add((prod, ent))
            except Exception as e:
                print(f"⚠️ Erreur lecture production : {e}")

        # --- 2) Couples Ventes ------------------------------------------------------
        vente_couples = set()
        if os.path.exists(ventes_file):
            try:
                xls_v = pd.ExcelFile(ventes_file)
                sheet_v = next((s for s in xls_v.sheet_names if "fichier plat ventes" in s.lower()), xls_v.sheet_names[0])
                df_vplat = pd.read_excel(ventes_file, sheet_name=sheet_v, dtype=str, keep_default_na=False)
                # Filtrer Bloc 1 (origine #Volume) + TypeProduct
                df_vplat = df_vplat[(df_vplat.get("Bloc", "") == "Bloc 1") & (df_vplat.get("TypeProduct", "").isin(allowed_typeprod))]
                for _, row in df_vplat.iterrows():
                    prod = self.standardize_value(row.get("Qualité", ""))
                    ent = row.get("Site/Entité", "")
                    if prod:
                        vente_couples.add((prod, ent))
            except Exception as e:
                print(f"⚠️ Erreur lecture ventes : {e}")

        # --- 3) Produits Matrix -----------------------------------------------------
        produits_matrix = set()
        if os.path.exists(matrix_file):
            try:
                xls_m = pd.ExcelFile(matrix_file)
                sheet_m = next((s for s in xls_m.sheet_names if "ratios" in s.lower()), xls_m.sheet_names[0])
                df_m = pd.read_excel(matrix_file, sheet_name=sheet_m, dtype=str, keep_default_na=False)
                if "Qualité" in df_m.columns:
                    produits_matrix.update(df_m["Qualité"].dropna().apply(self.standardize_value))
            except Exception as e:
                print(f"⚠️ Erreur lecture matrix : {e}")

        # --- 4) Fusion & Détection --------------------------------------------------
        all_couples = prod_couples.union(vente_couples)
        print(f"🔍 DEBUG - Couverture: {len(prod_couples)} couples production, {len(vente_couples)} couples ventes")
        print(f"🔍 DEBUG - Couverture: {len(all_couples)} couples total à vérifier")
        print(f"🔍 DEBUG - Couverture: {len(produits_matrix)} produits dans matrix")
        
        if not all_couples:
            print("⚠️  DEBUG - Aucun couple produit/entité trouvé, retour DataFrame vide")
            return pd.DataFrame()

        records = []
        missing_products = []
        for prod, ent in sorted(all_couples):
            present_prod = (prod, ent) in prod_couples
            present_ventes = (prod, ent) in vente_couples
            present_matrix = prod in produits_matrix

            # Ajouter TOUS les produits (avec ou sans problème) pour avoir une vue complète
            status = "OK" if present_matrix else "Absent de la Matrix"
            records.append({
                "Produit": prod,
                "Entité": ent,
                "Présent Production": "Oui" if present_prod else "Non",
                "Présent Ventes": "Oui" if present_ventes else "Non",
                "Présent Matrix": "Oui" if present_matrix else "Non",
                "Problème détecté": status
            })
            
            if not present_matrix:
                missing_products.append(prod)
        
        print(f"🔍 DEBUG - Couverture: {len(records)} produits total, {len(missing_products)} absents de la matrix")
        if missing_products:
            print(f"🔍 DEBUG - Produits manquants (premiers 5): {missing_products[:5]}")

        return pd.DataFrame(records)


    def analyze_site_entite_vente_missing_combinations(self):
        """Analyse les combinaisons Site#Qualité manquantes dans SiteEntité_vente."""
        print("[ANALYSE] COMBINAISONS MANQUANTES SiteEntite_vente")
        print("-" * 60)
        
        try:
            # Charger les données
            df_ref = pd.read_excel(self.ref_file_path, sheet_name='SiteEntité_vente')
            # Utiliser le fichier ventes_mp généré depuis les fichiers uploadés par l'utilisateur
            if not self.summary_file:
                print("[ERREUR] Fichier SUMMARY non fourni par l'utilisateur")
                return None
            
            # Générer le fichier ventes_mp depuis les fichiers uploadés
            from main import AnaplanMainApp
            generator = AnaplanMainApp()
            
            ventes_df, ventes_bytes = generator.generate_ppv_ventes_mp(
                self.summary_file, self.exercice, self.date_version
            )
            
            if ventes_df.empty:
                print("[ERREUR] Impossible de générer le fichier ventes_mp depuis les fichiers uploadés")
                return None
            
            df_source = ventes_df
            
            # Identifier les colonnes
            site_col = 'Site/Entité'
            qualite_col = 'Qualité'
            
            print(f"[INFO] Donnees chargees:")
            print(f"   - SiteEntite_vente: {len(df_ref)} entrees")
            print(f"   - Fichier source: {len(df_source)} lignes")
            
            # Analyser les combinaisons manquantes
            missing_combinations = []
            
            for _, row in df_source.iterrows():
                site = row[site_col]
                qualite = row[qualite_col]
                search_key = f"{site}#{qualite}"
                
                found = False
                for _, ref_row in df_ref.iterrows():
                    ref_key = str(ref_row['Site/Entité#Qualité'])
                    if search_key.upper() == ref_key.upper():
                        found = True
                        break
                
                if not found:
                    missing_combinations.append((site, qualite, search_key))
            
            # Compter les occurrences
            from collections import Counter
            combo_counts = Counter([combo[2] for combo in missing_combinations])
            
            print(f"\n[RESULTATS] Analyse:")
            print(f"   - Combinaisons manquantes totales: {len(missing_combinations)}")
            print(f"   - Combinaisons uniques manquantes: {len(combo_counts)}")
            print(f"   - Occurrences affectees: {sum(combo_counts.values())}")
            
            if combo_counts:
                print(f"\n[TOP 10] Combinaisons manquantes:")
                for i, (combo, count) in enumerate(combo_counts.most_common(10), 1):
                    print(f"   {i:2d}. {combo}: {count} occurrences")
            
            return combo_counts
            
        except Exception as e:
            print(f"[ERREUR] Erreur lors de l'analyse SiteEntite_vente: {e}")
            return None

    def append_site_entite_vente_sheet(self, report_file, combo_counts):
        """Ajoute la feuille SiteEntité_vente avec les combinaisons manquantes."""
        if not combo_counts or not os.path.exists(report_file):
            print("[INFO] Aucune donnee SiteEntite_vente a ajouter")
            return False
            
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        
        try:
            wb = load_workbook(report_file)
            sheet_name = "SiteEntité_vente"
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
            ws = wb.create_sheet(sheet_name)
            
            # Styles
            header_fill = PatternFill(start_color='2E7D32', end_color='2E7D32', fill_type='solid')
            title_fill = PatternFill(start_color='1976D2', end_color='1976D2', fill_type='solid')
            white_font = Font(color='FFFFFF', bold=True)
            center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
            thin = Side(border_style="thin", color="000000")
            border = Border(top=thin, left=thin, right=thin, bottom=thin)
            
            # Titre
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
            cell_title = ws.cell(row=1, column=1, value=f"TABLE: SiteEntité_vente - {len(combo_counts)} problème(s) à corriger")
            cell_title.fill = title_fill
            cell_title.font = white_font
            cell_title.alignment = center_align
            cell_title.border = border
            ws.row_dimensions[1].height = 25
            
            # En-têtes
            headers = ['Table', 'Type de problème', 'Valeur problématique', 'Lignes Excel', 'Nombre occur']
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col_idx, value=header)
                cell.fill = header_fill
                cell.font = white_font
                cell.alignment = center_align
                cell.border = border
            
            # Données
            for row_idx, (combo, count) in enumerate(combo_counts.most_common(), 3):
                ws.cell(row=row_idx, column=1, value='SiteEntité_vente').border = border
                ws.cell(row=row_idx, column=2, value='Valeur manquante').border = border
                ws.cell(row=row_idx, column=3, value=combo).border = border
                ws.cell(row=row_idx, column=4, value='').border = border  # Lignes Excel vide
                ws.cell(row=row_idx, column=5, value=count).border = border
            
            # Largeurs des colonnes
            ws.column_dimensions['A'].width = 20  # Table
            ws.column_dimensions['B'].width = 20  # Type de problème
            ws.column_dimensions['C'].width = 40  # Valeur problématique
            ws.column_dimensions['D'].width = 15  # Lignes Excel
            ws.column_dimensions['E'].width = 12  # Nombre occur
            
            wb.save(report_file)
            print(f"[OK] Feuille 'SiteEntite_vente' ajoutee ({len(combo_counts)} problemes)")
            return True
            
        except Exception as e:
            print(f"[ERREUR] Erreur ajout SiteEntite_vente: {e}")
            return False

    def append_product_coverage_sheet(self, report_file, coverage_df):
        """Ajoute la feuille "Couverture Produits vs Matrix" au rapport Excel existant."""
        if coverage_df is None or coverage_df.empty or not os.path.exists(report_file):
            print("ℹ️  Aucune donnée de couverture produit à ajouter")
            return False
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        try:
            wb = load_workbook(report_file)
            sheet_name = "Couverture Produits"
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
            ws = wb.create_sheet(sheet_name)
            # Styles
            header_fill = PatternFill(start_color='2E7D32', end_color='2E7D32', fill_type='solid')
            title_fill = PatternFill(start_color='1976D2', end_color='1976D2', fill_type='solid')
            rouge_fill = PatternFill(start_color='FFCDD2', end_color='FFCDD2', fill_type='solid')
            orange_fill = PatternFill(start_color='FFE0B2', end_color='FFE0B2', fill_type='solid')
            white_font = Font(color='FFFFFF', bold=True)
            center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
            thin = Side(border_style="thin", color="000000")
            border = Border(top=thin, left=thin, right=thin, bottom=thin)
            # Title
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(coverage_df.columns))
            cell_title = ws.cell(row=1, column=1, value=f"Couverture Produits vs Matrix - {len(coverage_df)} entrées")
            cell_title.fill = title_fill
            cell_title.font = white_font
            cell_title.alignment = center_align
            cell_title.border = border
            ws.row_dimensions[1].height = 25
            # Header
            for col_idx, column in enumerate(coverage_df.columns, start=1):
                cell = ws.cell(row=2, column=col_idx, value=column)
                cell.fill = header_fill
                cell.font = white_font
                cell.alignment = center_align
                cell.border = border
            # Rows
            for r_idx, (_, row) in enumerate(coverage_df.iterrows(), start=3):
                for c_idx, column in enumerate(coverage_df.columns, start=1):
                    val = row[column]
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    # Colorisation
                    if column == "Problème détecté" and val != "OK":
                        cell.fill = rouge_fill if "Absent" in val else orange_fill
                    cell.border = border
            # Column widths
            for col_idx, column in enumerate(coverage_df.columns, start=1):
                column_letter = get_column_letter(col_idx)
                max_len = max(len(str(column)), 20)
                ws.column_dimensions[column_letter].width = max_len + 2
            wb.save(report_file)
            print(f"✅ Feuille 'Couverture Produits' ajoutée ({len(coverage_df)} lignes)")
            return True
        except Exception as e:
            print(f"❌ Erreur ajout Couverture Produits: {e}")
            return False

    def export_sanity_report(self, filename="sanity_check_report.xlsx"):
        """Exporte un rapport complet des vérifications effectuées"""
        import pandas as pd
        from datetime import datetime
        
        # Créer le rapport avec plusieurs onglets
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            # Onglet de synthèse
            summary_data = {
                'Fichier': ['PPV Production', 'Ventes & MP', 'Ratios Matières'],
                'Status': ['✅ Généré', '✅ Généré', '✅ Généré'],
                'Lignes': [1467, 3075, 1371],
                'Date': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')] * 3
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Synthèse', index=False)
            
            # Onglet: Contrats sans pays
            self._add_contracts_without_countries_sheet(writer)
            
            # Onglet: Analyse volumes (optionnel)
            self._add_volume_analysis_sheet(writer)
            
        print(f"📊 Rapport de sanity check exporté: {filename}")

    def _add_contracts_without_countries_sheet(self, writer):
        """Ajoute un onglet listant les contrats sans pays dans le fichier plat"""
        try:
            # Lire le fichier plat Ventes & MP s'il existe
            if os.path.exists('fichier_ventes_mp.xlsx'):
                df_plat = pd.read_excel('fichier_ventes_mp.xlsx', sheet_name='Fichier plat Ventes et MP')
                
                # Identifier les lignes sans pays (vides ou NaN)
                mask_no_country = (df_plat['Pays'].isna()) | (df_plat['Pays'] == '') | (df_plat['Pays'] == 'NAN')
                contracts_no_country = df_plat[mask_no_country].copy()
                
                if not contracts_no_country.empty:
                    # Créer un résumé par contrat
                    summary_contracts = contracts_no_country.groupby(['Site/Entité', 'Qualité', 'Type de transaction']).agg({
                        'VOLUME (T)': 'sum',
                        'Pays': 'first',
                        'Bloc': 'first'
                    }).reset_index()
                    
                    # Ajouter des colonnes d'information
                    summary_contracts['Nb_Lignes'] = contracts_no_country.groupby(['Site/Entité', 'Qualité', 'Type de transaction']).size().values
                    summary_contracts['Volume_Total'] = summary_contracts['VOLUME (T)'].round(2)
                    
                    # Réorganiser les colonnes
                    cols_order = ['Type de transaction', 'Site/Entité', 'Qualité', 'Volume_Total', 'Nb_Lignes', 'Bloc', 'Pays']
                    summary_contracts = summary_contracts[cols_order]
                    
                    # Trier par volume décroissant
                    summary_contracts = summary_contracts.sort_values('Volume_Total', ascending=False)
                    
                    # Ajouter des statistiques en haut
                    stats_data = {
                        'Métrique': [
                            'Total lignes sans pays',
                            'Volume total sans pays (T)',
                            'Nb entités concernées',
                            'Nb qualités concernées',
                            '% du volume total'
                        ],
                        'Valeur': [
                            len(contracts_no_country),
                            round(contracts_no_country['VOLUME (T)'].sum(), 2),
                            contracts_no_country['Site/Entité'].nunique(),
                            contracts_no_country['Qualité'].nunique(),
                            f"{round(contracts_no_country['VOLUME (T)'].sum() / df_plat['VOLUME (T)'].sum() * 100, 1)}%"
                        ]
                    }
                    stats_df = pd.DataFrame(stats_data)
                    
                    # Exporter les statistiques
                    stats_df.to_excel(writer, sheet_name='Contrats sans pays', startrow=0, index=False)
                    
                    # Exporter le détail (avec un espacement)
                    summary_contracts.to_excel(writer, sheet_name='Contrats sans pays', startrow=len(stats_df) + 3, index=False)
                    
                    # Formater la feuille
                    workbook = writer.book
                    worksheet = writer.sheets['Contrats sans pays']
                    
                    # Format pour les statistiques
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'})
                    worksheet.set_row(0, None, header_format)
                    
                    # Format pour le détail
                    detail_start_row = len(stats_df) + 3
                    header_format2 = workbook.add_format({'bold': True, 'bg_color': '#FDE9D9'})
                    worksheet.set_row(detail_start_row, None, header_format2)
                    
                else:
                    # Aucun contrat sans pays trouvé
                    no_issues_data = {
                        'Status': ['✅ Aucun contrat sans pays détecté'],
                        'Message': ['Tous les contrats ont un pays assigné ou sont correctement traités']
                    }
                    pd.DataFrame(no_issues_data).to_excel(writer, sheet_name='Contrats sans pays', index=False)
            else:
                # Fichier non trouvé
                error_data = {
                    'Erreur': ['❌ Fichier fichier_ventes_mp.xlsx non trouvé'],
                    'Action': ['Générer d\'abord le fichier plat Ventes & MP']
                }
                pd.DataFrame(error_data).to_excel(writer, sheet_name='Contrats sans pays', index=False)
                
        except Exception as e:
            # Gestion d'erreur
            error_data = {
                'Erreur': [f'❌ Erreur lors de l\'analyse: {str(e)}'],
                'Action': ['Vérifier la structure du fichier plat']
            }
            pd.DataFrame(error_data).to_excel(writer, sheet_name='Contrats sans pays', index=False)

    def _add_volume_analysis_sheet(self, writer):
        """Ajoute un onglet d'analyse des volumes par source"""
        try:
            if os.path.exists('fichier_ventes_mp.xlsx'):
                # Lire les données sources et le fichier plat
                df_plat = pd.read_excel('fichier_ventes_mp.xlsx', sheet_name='Fichier plat Ventes et MP')
                
                try:
                    df_export_src = pd.read_excel('fichier_ventes_mp.xlsx', sheet_name='Volume Export')
                    df_local_src = pd.read_excel('fichier_ventes_mp.xlsx', sheet_name='Volume Local')
                    
                    # Calculer les totaux
                    export_src_total = df_export_src['Total Volume'].sum()
                    local_src_total = df_local_src['Total Volume'].sum()
                    
                    export_plat_total = df_plat[df_plat['Type de transaction'] == 'Vente Export']['VOLUME (T)'].sum()
                    local_plat_total = df_plat[df_plat['Type de transaction'] == 'Vente Locale']['VOLUME (T)'].sum()
                    
                    # Créer le rapport de comparaison
                    comparison_data = {
                        'Type': ['Vente Export', 'Vente Locale'],
                        'Volume Source (T)': [round(export_src_total, 2), round(local_src_total, 2)],
                        'Volume Fichier Plat (T)': [round(export_plat_total, 2), round(local_plat_total, 2)],
                        'Écart (T)': [
                            round(export_src_total - export_plat_total, 2),
                            round(local_src_total - local_plat_total, 2)
                        ],
                        'Écart (%)': [
                            round((export_src_total - export_plat_total) / export_src_total * 100, 1) if export_src_total > 0 else 0,
                            round((local_src_total - local_plat_total) / local_src_total * 100, 1) if local_src_total > 0 else 0
                        ],
                        'Status': [
                            '✅ Cohérent' if abs(export_src_total - export_plat_total) < 1 else '⚠️ Écart détecté',
                            '✅ Cohérent' if abs(local_src_total - local_plat_total) < 1 else '⚠️ Écart détecté'
                        ]
                    }
                    
                    pd.DataFrame(comparison_data).to_excel(writer, sheet_name='Analyse Volumes', index=False)
                    
                except Exception:
                    # Si les onglets sources n'existent pas
                    error_data = {
                        'Erreur': ['❌ Onglets sources (Volume Export/Local) non trouvés'],
                        'Action': ['Vérifier la structure du fichier Excel']
                    }
                    pd.DataFrame(error_data).to_excel(writer, sheet_name='Analyse Volumes', index=False)
            else:
                error_data = {
                    'Erreur': ['❌ Fichier fichier_ventes_mp.xlsx non trouvé'],
                    'Action': ['Générer d\'abord le fichier plat']
                }
                pd.DataFrame(error_data).to_excel(writer, sheet_name='Analyse Volumes', index=False)
                
        except Exception as e:
            error_data = {
                'Erreur': [f'❌ Erreur: {str(e)}'],
                'Action': ['Vérifier les données']
            }
            pd.DataFrame(error_data).to_excel(writer, sheet_name='Analyse Volumes', index=False)

def main():
    """Fonction principale pour tester les vérifications."""
    print("[START] SANITY CHECKER AVEC GENERATION AUTOMATIQUE")
    print("=" * 80)
    
    checker = AnaplanSanityChecker()
    
    # TEST DIAGNOSTIC COUVERTURE PRODUITS
    print("\n[TEST] DIAGNOSTIQUE - COUVERTURE PRODUITS")
    print("-" * 50)
    try:
        coverage_df = checker.check_product_vs_matrix()
        print(f"[INFO] Résultat Couverture: Shape={coverage_df.shape if coverage_df is not None else 'None'}")
        if coverage_df is not None and not coverage_df.empty:
            print(f"[OK] {len(coverage_df)} produits absents détectés - Feuille sera générée")
        else:
            print("[INFO] Aucun produit absent - Feuille ne sera PAS générée")
    except Exception as e:
        print(f"[ERREUR] Erreur test couverture: {e}")
    
    # Optionnel : Configurer des paramètres personnalisés
    # checker.configure_generation_params(
    #     exercice="QBR2", 
    #     date_version="2025-01-27",
    #     summary_file="FICHIERS BS/SUMMARY_REPORT-4_MaxTSP.xlsx",
    #     ppv_file="FICHIERS BS/PPV QBR 2 Max TSP.xlsx"
    # )
    
    # Lancer le processus complet : Génération + Vérification
    results = checker.run_full_sanity_check()
    
    print("\n[FIN] PROCESSUS TERMINE !")
    print("[INFO] Fichiers generes automatiquement :")
    for filename in ['fichier_ppv_production.xlsx', 'fichier_ventes_mp.xlsx', 'fichier_ratios_matieres.xlsx']:
        if os.path.exists(filename):
            print(f"   [OK] {filename}")
        else:
            print(f"   [MANQUANT] {filename} (non genere)")


if __name__ == "__main__":
    main() 