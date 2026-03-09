def analyser(image_path):
    print(f"[*] Ouverture de l'image : {image_path}")
    img = pytsk3.Img_Info(url=image_path)
    fs = None
    
    # 1. NOUVELLE MÉTHODE : Recherche intelligente de la partition OS
    try:
        volume = pytsk3.Volume_Info(img)
        partitions_lisibles = [] 
        
        print("[*] Analyse de la table des partitions...")
        for part in volume:
            offset = part.start * 512
            try:
                # Tente d'ouvrir la partition
                test_fs = pytsk3.FS_Info(img, offset=offset)
                
                # Tente de lire la racine de cette partition
                try:
                    racine = test_fs.open_dir(path="/")
                    # Récupère les noms des dossiers à la racine en minuscules
                    noms_racine = [entree.info.name.name.decode('utf-8', 'ignore').lower() for entree in racine if hasattr(entree, 'info') and hasattr(entree.info, 'name')]
                except:
                    noms_racine = []

                # Cherche les dossiers typiques d'un OS (Windows ou Linux/macOS)
                if 'users' in noms_racine or 'windows' in noms_racine or 'home' in noms_racine:
                    fs = test_fs
                    print(f"[+] Partition OS PRINCIPALE détectée (Offset : {offset}) !")
                    break # On a trouvé la bonne, on arrête de chercher
                else:
                    print(f"[-] Partition mineure ou de boot ignorée (Offset : {offset})")
                    partitions_lisibles.append(test_fs) 
                    
            except:
                pass # Impossible de lire cette partition, on passe à la suivante
                
        # Fallback : si on n'a pas vu "Users" ou "Windows", on prend la dernière partition lisible (souvent la plus grosse de données)
        if not fs and partitions_lisibles:
            fs = partitions_lisibles[-1]
            print("[!] Pas de marqueurs OS clairs. Utilisation de la dernière partition lisible par défaut.")

    except:
        pass # Erreur avec Volume_Info, on passe à la force brute
    
    # 2. Méthode "Force Brute" (Inchangement si la table est cassée)
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
        
    print("[*] Extraction de l'arborescence en cours (cela peut prendre du temps sur un OS complet)...")
    fichiers = []
    stats = {'fichiers': 0, 'dossiers': 0, 'extensions': {}}
    
    # Lancement de l'extraction récursive
    extraire_fichiers(fs.open_dir(path="/"), "", fichiers, stats)
    
    # 3. Export Excel direct
    print("[*] Création du fichier Excel...")
    resume = [{"Catégorie": "Total Dossiers", "Quantité": stats['dossiers']}, 
              {"Catégorie": "Total Fichiers", "Quantité": stats['fichiers']}]
    resume += [{"Catégorie": f"Ext {k}", "Quantité": v} for k, v in sorted(stats['extensions'].items(), key=lambda x: x[1], reverse=True)]
    
    with pd.ExcelWriter("Rapport_Forensique.xlsx") as writer:
        pd.DataFrame(resume).to_excel(writer, index=False, sheet_name="Résumé")
        pd.DataFrame(fichiers).to_excel(writer, index=False, sheet_name="Fichiers")
        
    print(f"[+] Terminé ! Le fichier Rapport_Forensique.xlsx est prêt avec {stats['fichiers']} fichiers extraits.")
