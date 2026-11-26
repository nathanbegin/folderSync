# FolderSync

Synchronisation sûre entre un dossier local (ex. `C:\...`) et OneDrive, avec aperçu des différences et archivage de toutes les anciennes versions. Python 3.9+ (tkinter inclus sur Windows), aucune dépendance.

## Utilisation

```bash
python foldersync.py                      # interface graphique
python foldersync.py --cli                # aperçu en console
python foldersync.py --cli --apply        # synchroniser en console
python foldersync.py --install-task --interval 15   # arrière-plan (Windows), à chaque ouverture de session
python foldersync.py --remove-task
```

1. Choisissez la source et la destination, puis **Analyser**.
2. Décochez ou changez les actions (clic droit), double-clic = différences texte.
3. **Synchroniser**.

## Protection contre la perte de données

- Aperçu avant toute action ; rien n'est écrasé ni supprimé directement : l'ancienne version va dans `versions/<date_heure>/<source|destination>/…`.
- Mémoire du dernier état synchronisé : détecte les vrais conflits et distingue « supprimé » de « jamais copié ».
- Conflit : les deux versions sont conservées (`fichier (conflit …).ext`).
- Copies atomiques vérifiées par SHA-256.
- Garde-fous : dossier introuvable ou illisible = arrêt ; beaucoup de suppressions = confirmation (cycle sauté en arrière-plan).
- Fichiers OneDrive « en ligne seulement » non téléchargés par défaut.
- Journal dans `%APPDATA%\FolderSync\logs`, purge des versions après N jours.

Les dossiers vides ne sont pas synchronisés.

## Exécutable Windows (.exe)

Construit automatiquement par GitHub Actions : onglet **Actions → Build Windows exe → Run workflow** (artefact `FolderSync-windows`), ou en modifiant le fichier `VERSION` (la release `vX.Y.Z` est alors créée dans **Releases**). Avec le `.exe` : `FolderSync.exe --install-task --interval 15`.
