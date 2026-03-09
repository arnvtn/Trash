import sys
import os
import datetime
import pyvmdk
import pytsk3
import pandas as pd

# ---------------------------------------------------------
# Classe adaptateur pour relier pyvmdk à pytsk3
# ---------------------------------------------------------
class VMDK_Img_Info(pytsk3.Img_Info):
    def __init__(self, vmdk_handle):
        self._vmdk_handle = vmdk_handle
        super().__init__(url="", type=pytsk3.TSK_IMG_TYPE_EXTERNAL)

    def close(self):
        self._vmdk_handle.close()

    def read(self, offset, size):
        self._vmdk_handle.seek(offset)
        return self._vmdk_handle.read(size)

    def get_size(self):
        return self._vmdk_handle.get_media_size()

# ---------------------------------------------------------
# Fonction récursive pour parcourir le système de fichiers
# ---------------------------------------------------------
def parcourir_repertoire(repertoire, chemin_actuel, liste_fichiers, stats):
    for entree in repertoire:
        try:
            # Récupérer le nom du fichier/dossier
            nom = entree.info.name.name.decode('utf-8', errors='replace')
            
            # Ignorer les dossiers de navigation pour éviter les boucles infinies
            if nom in [".", ".."]:
                continue

            chemin_complet = f"{chemin_actuel}/{nom}".replace('//', '/')
            meta = entree.info.meta
            
            # Récupérer la taille et la date de modification
            taille = meta.size if meta else 0
            mtime = meta.mtime if meta else 0
            if mtime:
                date_modif = datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
            else:
                date_modif = "Inconnue"
            
            # Si c'est un répertoire
            if meta and meta.type == pytsk3.TSK_FS_META_TYPE_DIR:
                stats['total_repertoires'] += 1
                try:
                    sous_repertoire = entree.as_directory()
                    parcourir_repertoire(sous_repertoire, chemin_complet, liste_fichiers, stats)
                except Exception:
                    pass # Dossier inaccessible (permissions, corrompu, etc.)
            
            # Si c'est un fichier standard
            elif meta and meta.type == pytsk3.TSK_FS_META_TYPE_REG:
                stats['total_fichiers'] += 1
                
                # Extraction de l'extension
                _, ext = os.path.splitext(nom)
                ext = ext.lower() if ext else "Sans extension"
                stats['extensions'][ext] = stats['extensions'].get(ext, 0) + 1
                
                # Ajout à la liste des détails
                liste_fichiers.append({
                    "Nom du fichier": nom,
                    "Extension": ext,
                    "Chemin d'accès": chemin_complet,
                    "Date de modification": date_modif,
                    "Taille (octets)": taille
                })
        except Exception:
            continue

# ---------------------------------------------------------
# Fonction principale
# ---------------------------------------------------------
def analyser_vmdk(vmdk_path, excel_output_path):
    print(f"[*] Ouverture de l'image VMDK : {vmdk_path}")
    
    # 1. Ouverture du VMDK
    vmdk_handle = pyvmdk.handle()
    try:
        vmdk_handle.open(vmdk_path)
    except Exception as e:
        print(f"[!] Erreur lors de l'ouverture du VMDK : {e}")
        sys.exit(1)
        
    img_info = VMDK_Img_Info(vmdk_handle)

    # 2. Recherche du bon volume (Heuristique basée sur la signature de l'OS)
    try:
        vol_info = pytsk3.Volume_Info(img_info)
    except IOError:
        vol_info = None

    meilleur_fs = None
    meilleur_score = -1
    taille_fallback = 0

    # Liste des dossiers typiques à la racine d'un OS (Windows & Linux)
    marqueurs_os = ['windows', 'users', 'program files', 'etc', 'usr', 'home', 'var', 'bin']

    if vol_info:
        print("[*] Table de partitions détectée. Analyse des signatures de chaque volume...")
        for partition in vol_info:
            try:
                # Tente de monter la partition
                fs = pytsk3.FS_Info(img_info, offset=partition.start * vol_info.info.block_size)
                
                # Calcule un "score OS" en cherchant des dossiers spécifiques
                score_actuel = 0
                try:
                    racine = fs.open_dir(path="/")
                    for entree in racine:
                        nom = entree.info.name.name.decode('utf-8', errors='replace').lower()
                        if nom in marqueurs_os:
                            score_actuel += 1
                except Exception:
                    pass # Impossible de lire la racine de cette partition

                # Logique de sélection : on prend le meilleur score. 
                # En cas d'égalité (ex: 0 pour des partitions de données), on prend la plus grande taille.
                if score_actuel > meilleur_score or (score_actuel == meilleur_score and partition.len > taille_fallback):
                    meilleur_score = score_actuel
                    taille_fallback = partition.len
                    meilleur_fs = fs
                    print(f"    -> Volume candidat trouvé (Score: {score_actuel}, Taille: {partition.len})")
                    
            except Exception:
                continue # N'est pas un système de fichiers montable
    else:
        print("[*] Aucune table de partitions. Tentative de montage brut...")
        try:
            meilleur_fs = pytsk3.FS_Info(img_info)
        except Exception as e:
            print(f"[!] Impossible de trouver un système de fichiers : {e}")

    if not meilleur_fs:
        print("[!] Aucun système de fichiers exploitable n'a été trouvé dans l'image.")
        img_info.close() # Libération si échec
        sys.exit(1)

    print("[*] Volume principal sélectionné. Début de l'indexation (cela peut prendre du temps)...")

    # 3. Initialisation des variables de stockage
    liste_fichiers = []
    stats = {
        'total_fichiers': 0,
        'total_repertoires': 0,
        'extensions': {}
    }

    # 4. Parcours du système de fichiers
    repertoire_racine = meilleur_fs.open_dir(path="/")
    parcourir_repertoire(repertoire_racine, "", liste_fichiers, stats)

    print("[*] Indexation terminée. Génération du fichier Excel...")

    # 5. Préparation des données pour Excel
    resume_data = [
        {"Catégorie": "Total des fichiers", "Quantité": stats['total_fichiers']},
        {"Catégorie": "Total des répertoires", "Quantité": stats['total_repertoires']}
    ]
    
    # Tri des extensions par ordre décroissant de fréquence
    extensions_triees = sorted(stats['extensions'].items(), key=lambda x: x[1], reverse=True)
    for ext, count in extensions_triees:
        resume_data.append({"Catégorie": f"Fichiers {ext}", "Quantité": count})

    df_resume = pd.DataFrame(resume_data)
    df_fichiers = pd.DataFrame(liste_fichiers)

    # 6. Écriture dans le fichier Excel
    with pd.ExcelWriter(excel_output_path, engine='openpyxl') as writer:
        df_resume.to_excel(writer, sheet_name='Résumé', index=False)
        df_fichiers.to_excel(writer, sheet_name='Liste des fichiers', index=False)

    print(f"[+] Succès ! Le rapport a été généré : {excel_output_path}")

    # 7. Nettoyage et fermeture des handles
    print("[*] Fermeture de l'image VMDK et libération des ressources...")
    img_info.close()

# Lancement du script
if __name__ == "__main__":
    # Remplacer par le chemin de votre image VMDK
    CHEMIN_VMDK = "chemin/vers/votre/image.vmdk" 
    FICHIER_EXCEL = "Rapport_FileSystem.xlsx"
    
    analyser_vmdk(CHEMIN_VMDK, FICHIER_EXCEL)
