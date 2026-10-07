# Rondes parkings

Application web hors ligne pour les rondes des agents de surveillance des parkings municipaux.

**Ce dépôt ne contient que le programme. Aucune donnée de ronde n'y transite.** Les saisies restent sur le téléphone des agents jusqu'à leur transmission, par fichier, au téléphone superviseur. L'heure et la position de chaque saisie sont chiffrées dès la saisie : seul le téléphone superviseur peut les lire.

---

## 1. Mise en ligne (une fois, environ 10 minutes)

1. Créer un compte gratuit sur github.com, de préférence avec une adresse professionnelle et un nom neutre (ex. `dpms-cachan`).
2. **New repository** → nom : `rondes` → **Public** (obligatoire : GitHub Pages n'est gratuit que pour les dépôts publics ; seul le programme est public, aucune donnée) → **Create repository**. Si le dépôt a été créé en privé : **Settings → General → Danger Zone → Change visibility → Make public**.
3. Cliquer **uploading an existing file**, glisser **tout le contenu** du dossier décompressé (y compris le dossier `lib`), puis **Commit changes**.
4. **Settings → Pages** → *Source* : **Deploy from a branch** → branche `main`, dossier `/ (root)` → **Save**.
5. Après une à deux minutes, l'application est en ligne à l'adresse `https://<compte>.github.io/rondes/`.

## 2. Téléphone superviseur

1. Ouvrir l'adresse dans Chrome, puis menu ⋮ → **Installer l'application** (ou « Ajouter à l'écran d'accueil »).
2. Ouvrir l'icône **Rondes** → **Créer le superviseur** → choisir une phrase de passe (12 caractères minimum).
3. **Télécharger la clé de secours**. La verser sur un stockage de la Ville et remettre la phrase de passe sous pli fermé à la DGS.
4. **Listes et coordonnées** : saisir les prénoms des agents (ils ne figurent volontairement pas dans le programme publié), vérifier les destinataires et la check-list, renseigner le téléphone et le courriel d'urgence.
5. **QR codes → Imprimer les QR des parkings** (une page A4 par parking). À plastifier et fixer à l'entrée de chaque parking.

## 3. Téléphone des agents (téléphone de service partagé)

1. Ouvrir l'adresse dans Chrome, puis **Installer l'application**.
2. Sur le téléphone superviseur : **QR codes → Afficher le QR de configuration**. Le scanner avec l'appareil photo du téléphone agents → **Oui, configurer**.
3. Lors de la première ronde, **autoriser la position et l'appareil photo**.
4. **Réglages** : vérifier « Stockage protégé : Oui ». Sinon, ouvrir l'application depuis l'icône installée, puis demander la protection.

## 4. Fonctionnement courant

| Qui | Quand | Geste |
|---|---|---|
| Agent | Arrivée au parking | Scanner le QR → choisir son nom → revue des signalements en cours → check-list → signalements → « Terminer la ronde » |
| Agent | Chaque semaine (ou au passage du DPMS) | **Transmettre au DPMS** → Quick Share vers le téléphone superviseur (ou courriel) |
| DPMS | À réception | **Importer une transmission** → vérifier le rapport d'intégrité → traiter et clôturer les signalements |
| DPMS | Après import | **Préparer la mise à jour** → Quick Share vers le téléphone agents → l'agent ouvre **Recevoir la mise à jour du DPMS** (accusé de réception, clôtures, purge de la mémoire du téléphone) |
| DPMS | Chaque semaine | **Sauvegarde complète** → stockage de la Ville (OneDrive professionnel ou lecteur réseau) |

Urgences : bouton d'appel et de courriel sur l'accueil du téléphone agents. La transmission hebdomadaire ne convient pas aux urgences.



## Bibliothèques incluses (licences libres)

- fflate (MIT) : compression ZIP.
- qrcode-generator (MIT) : génération des QR codes.
- jsQR (Apache 2.0) : lecture des QR codes si le navigateur n'a pas de lecteur intégré.
- SheetJS Community Edition 0.18.5 (Apache 2.0) : écriture des fichiers Excel.
