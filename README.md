# Rondes parkings — Ville de Cachan (DPMS)

Application web hors ligne pour les rondes des agents de surveillance des parkings municipaux (Hénouille, Dumotel, Arobase). Version 1.1.

**Principe.** L'outil appartient aux agents : ils font leurs rondes, signalent, suivent et clôturent eux-mêmes leurs signalements. À la fin de chaque ronde, tout part automatiquement, chiffré, vers un relais. Le téléphone superviseur relève tout seul à son ouverture : on y consulte rondes et signalements. Les interventions du superviseur (clôturer, rouvrir, modifier, écrire aux agents) sont ponctuelles, facultatives, et redescendent seules vers le téléphone des agents.

**Confidentialité.** Ce dépôt ne contient que le programme. Tout ce qui transite par le relais est chiffré sur le téléphone avant l'envoi : descriptions, plaques, noms, photos, heures, positions. Seul le téléphone superviseur peut le lire. Le retour vers les agents est chiffré pour le seul téléphone des agents.

---

## 1. Mise en ligne de l'application (une fois, environ 10 minutes)

1. Créer un compte gratuit sur github.com, de préférence avec une adresse professionnelle et un nom neutre (ex. `dpms-cachan`).
2. **New repository** → nom : `rondes` → **Public** → **Create repository**. Le dépôt doit être public : GitHub Pages n'est gratuit que pour les dépôts publics. Seul le programme est public, aucune donnée. Si le dépôt a été créé en privé : **Settings → General → Danger Zone → Change visibility → Make public**.
3. Cliquer **uploading an existing file**, glisser **tout le contenu** du dossier décompressé (y compris les dossiers `lib` et `relais`), puis **Commit changes**.
4. **Settings → Pages** → *Source* : **Deploy from a branch** → branche `main`, dossier `/ (root)` → **Save**.
5. Après une à deux minutes, l'application est en ligne à l'adresse `https://<compte>.github.io/rondes/`.

## 2. Relais de transmission automatique (une fois, environ 10 minutes, gratuit)

À faire sur un ordinateur.

1. Créer un compte Google dédié au service (ex. `dpms.cachan.rondes@gmail.com`), pas un compte personnel.
2. Avec ce compte, ouvrir **script.google.com** → **Nouveau projet**. Effacer le contenu, coller tout le fichier `relais/relais.gs`, puis enregistrer.
3. **Déployer → Nouveau déploiement** → icône engrenage → **Application web**.
   - *Exécuter en tant que* : **Moi**.
   - *Qui a accès* : **Tout le monde**. L'accès est ensuite protégé par des codes que l'application crée elle-même.
   - **Déployer**, puis **Autoriser l'accès**. Google affiche « Google n'a pas validé cette application » : c'est normal pour un script personnel. Cliquer **Paramètres avancés → Accéder à … (non sécurisé)** → **Autoriser** (accès à Drive).
4. Copier l'**URL de l'application Web** (elle se termine par `/exec`).
5. Sur le téléphone superviseur : **Transmission automatique** → coller l'adresse → **Tester et activer**.
6. Sur le téléphone des agents : scanner **une dernière fois** le QR de configuration (téléphone superviseur → QR codes). Ensuite, tout est automatique, y compris les modifications de listes.

Le relais stocke les envois chiffrés dans un dossier « Relais rondes parkings (chiffré) » du Drive du compte dédié. Ils sont mis à la corbeille dès que le téléphone superviseur les a relevés.

Si le téléphone superviseur est perdu, restaurez sa sauvegarde : elle contient les codes du relais. À défaut, dans l'éditeur du script, **Paramètres du projet → Propriétés du script** : supprimer `SUP_TOKEN` et `AGENT_TOKEN`, puis réactiver le relais.

## 3. Téléphone superviseur

1. Ouvrir l'adresse de l'application dans Chrome, puis menu ⋮ → **Installer l'application**.
2. **Créer le superviseur** → choisir une phrase de passe de 12 caractères minimum.
3. **Télécharger la clé de secours**. La verser sur un stockage de la Ville et remettre la phrase de passe sous pli fermé à la DGS.
4. **Transmission automatique** : voir la section 2.
5. **Listes et coordonnées** : saisir les prénoms des agents (ils ne figurent volontairement pas dans le programme publié), vérifier les destinataires et la check-list, renseigner le téléphone et le courriel d'urgence.
6. **QR codes → Imprimer les QR des parkings** : une page A4 par parking, à plastifier et fixer à l'entrée de chaque parking.

## 4. Téléphone des agents (téléphone de service partagé)

1. Ouvrir l'adresse dans Chrome, puis **Installer l'application**.
2. Scanner le QR de configuration affiché sur le téléphone superviseur → **Oui, configurer**.
3. Lors de la première ronde, **autoriser la position et l'appareil photo**.
4. **Réglages** : vérifier « Stockage protégé : Oui ».

## 5. Fonctionnement courant

| Qui | Quand | Ce qui se passe |
|---|---|---|
| Agent | Arrivée au parking | Scan du QR → nom → revue des signalements en cours (« Résolu » clôt le signalement) → check-list → nouveaux signalements → « Terminer la ronde » |
| Téléphone agents | Fin de ronde | Envoi automatique. Sans réseau, l'envoi repart seul au retour du réseau ou à l'ouverture suivante. Un signalement « urgent » part aussitôt. |
| Téléphone superviseur | À chaque ouverture, et toutes les 2 minutes tant qu'il est ouvert | Relève, vérification d'intégrité, mise à jour des rondes et des signalements. Les contrôles signalent les points d'attention. |
| Superviseur | Quand il le souhaite | Clôturer, rouvrir, modifier, écrire aux agents. Transmis automatiquement. |
| Superviseur | Chaque semaine | **Sauvegarde complète** → stockage de la Ville (OneDrive professionnel ou lecteur réseau) |

Urgences : bouton d'appel et de courriel sur l'accueil du téléphone agents. Le relais n'alerte pas le superviseur en temps réel.

Secours sans relais : transmission par fichier, depuis les Réglages du téléphone agents et le menu « Secours » du superviseur. Un envoi reçu par les deux voies n'est compté qu'une fois.

## 6. Ce que l'application garantit et ne garantit pas

- **Chiffrement** : ECDH P-256, HKDF-SHA-256, AES-256-GCM. Le relais et le compte Google ne lisent rien. Le téléphone agents ne peut pas relire l'heure exacte ni la position de ses saisies.
- **Intégrité** : chaque saisie contient l'empreinte SHA-256 de la précédente, et chaque photo son empreinte. Une suppression, une modification ou un trou dans la série est signalé à la relève.
- **Limites** :
  - L'heure est celle du téléphone. Verrouiller le réglage automatique de l'heure si la DSI gère la flotte.
  - Une personne compétente en informatique, ayant le téléphone en main, peut fabriquer des saisies cohérentes tant qu'elles n'ont pas été envoyées. L'envoi automatique en fin de ronde réduit fortement cette fenêtre.
  - Le QR code peut être photographié et rejoué.
  - En sous-sol, la position est souvent absente ou imprécise. Elle est enregistrée avec sa précision.
- **Fragilité** : effacer les données de Chrome efface les saisies non encore envoyées. Consigne : ne jamais vider les données du navigateur sur ce téléphone.

## 7. Mise à jour du programme

Modifier les fichiers dans le dépôt, puis changer le numéro de version dans `sw.js` (`CACHE = 'rondes-cachan-x.y.z'`). Les téléphones prennent la nouvelle version à l'ouverture suivante, avec réseau.

## 8. Préalables à la mise en service (à ne pas sauter)

- Inscription du traitement au registre des activités de traitement de la Ville (art. 30 RGPD), après avis du délégué à la protection des données. Mentionner le relais comme lieu de passage de données chiffrées, supprimées après relève.
- Information écrite des agents : finalités, données collectées dont la position au moment des saisies, durées de conservation, droits (art. 13 RGPD).
- Consultation du comité social territorial (art. 54 du décret n° 2021-571 du 10 mai 2021).
- Finalité affichée : traçabilité des rondes et des constats, suivi des anomalies. Pas le contrôle du temps de travail (CE, 15 décembre 2017, n° 403776).

## Bibliothèques incluses (licences libres)

- fflate (MIT) : compression ZIP.
- qrcode-generator (MIT) : génération des QR codes.
- jsQR (Apache 2.0) : lecture des QR codes si le navigateur n'a pas de lecteur intégré.
- SheetJS Community Edition 0.18.5 (Apache 2.0) : écriture des fichiers Excel.
