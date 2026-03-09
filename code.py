import sys, os, datetime
import pytsk3
import pandas as pd

def extraire_fichiers(repertoire, chemin, fichiers, stats):
    for entree in repertoire:
        nom = entree.info.name.name.decode('utf-8', 'ignore')
        if nom in [".", ".."]: continue
        
        chemin_complet = f"{chemin}/{nom}"
        meta = entree.info.meta
        if not meta: continue

        if meta.type == pytsk3.TSK_FS_META_TYPE_DIR:
            stats['dossiers'] += 1
            try: extraire_fichiers(entree.as_directory(), chemin_complet, fichiers, stats)
            except: pass
            
        elif meta.type == pytsk3.TSK_FS_META_TYPE_REG:
            stats['fichiers'] += 1
            ext = os.path.splitext(nom)[1].lower() or "Sans extension"
            stats['extensions'][ext] = stats['extensions'].get(ext, 0) + 1
            
            mtime = datetime.datetime.fromtimestamp(meta.mtime).strftime('%Y-%m-%d %H:%M:%S') if meta.mtime else "Inconnu"
            fichiers.append({"Nom": nom, "Extension": ext, "Chemin": chemin_complet, "Modification": mtime, "Taille": meta.size})

def analyser(image_path):
    print(f"[*] Ouverture de l'image : {image_path}")
    img = pytsk3.Img_Info(url=image_path)
    fs = None
    
    # 1. Méthode classique : Recherche via la table de partitions
    try:
        for part in pytsk3.Volume_Info(img):
            try:
                fs = pytsk3.FS_Info(img, offset=part.start * 512)
                print(f"[+] Système de fichiers trouvé (Offset partition : {part.start * 512})")
                break
            except: pass
    except: pass
    
    # 2. Méthode "Force Brute" : On teste les emplacements standards si la table est cassée
    if not fs:
        print("[*] Table de partitions illisible. Test des offsets de secours...")
        for off in [0, 32256, 1048576, 16777216, 536870912]:
            try:
                fs = pytsk3.FS_Info(img, offset=off)
                print(f"[+] Succès : Système de fichiers forcé à l'offset {off} !")
                break
            except: pass
            
    if not fs:
        print("[!] ÉCHEC : pytsk3 ne reconnaît aucun système de fichiers exploitable.")
        print("[!] Vérifiez que votre fichier n'est pas chiffré (BitLocker) ou gravement corrompu.")
        return
        
    print("[*] Extraction de l'arborescence en cours...")
    fichiers = []
    stats = {'fichiers': 0, 'dossiers': 0, 'extensions': {}}
    extraire_fichiers(fs.open_dir(path="/"), "", fichiers, stats)
    
    # 3. Export Excel direct
    print("[*] Création du fichier Excel...")
    resume = [{"Catégorie": "Total Dossiers", "Quantité": stats['dossiers']}, 
              {"Catégorie": "Total Fichiers", "Quantité": stats['fichiers']}]
    resume += [{"Catégorie": f"Ext {k}", "Quantité": v} for k, v in sorted(stats['extensions'].items(), key=lambda x: x[1], reverse=True)]
    
    with pd.ExcelWriter("Rapport_Forensique.xlsx") as writer:
        pd.DataFrame(resume).to_excel(writer, index=False, sheet_name="Résumé")
        pd.DataFrame(fichiers).to_excel(writer, index=False, sheet_name="Fichiers")
        
    print("[+] Terminé ! Le fichier Rapport_Forensique.xlsx est prêt.")

if __name__ == "__main__":
    fichier_cible = sys.argv[1] if len(sys.argv) > 1 else "Sample-flat.vmdk"
    analyser(fichier_cible)
