# Rondes parkings — Ville de Cachan (DPMS)

Application web hors ligne pour les rondes des agents de surveillance des parkings municipaux (Hénouille, Dumotel, Arobase). Version 1.2.

**Principe.** L'outil appartient aux agents : ils font leurs rondes, signalent, suivent et clôturent eux-mêmes leurs signalements. À la fin de chaque ronde, tout part automatiquement vers un relais, qui conserve un journal chiffré commun. Chaque appareil reconstitue l'état à partir de ce journal :

- **plusieurs téléphones agents** : chacun voit les signalements des autres ;
- **plusieurs appareils superviseurs** (téléphone, ordinateur) : chacun voit la même chose ;
- **un appareil réinstallé, perdu ou remplacé** retrouve toutes les données déjà envoyées.

Les interventions du superviseur (clôturer, rouvrir, modifier, écrire aux agents) sont facultatives et sont transmises automatiquement à tous les appareils.

**Confidentialité.** Ce dépôt ne contient que le programme. Tout ce qui est stocké sur le relais est chiffré sur l'appareil avant l'envoi. Le relais et le compte Google ne lisent rien.

| Donnée | Lisible par |
|---|---|
| Signalements, constats, photos, configuration, messages aux agents | téléphones agents et superviseurs (clé d'équipe) |
| Heure exacte et position de chaque saisie, notes internes DPMS | superviseurs seuls (clé superviseur) |
| Clés du superviseur (« coffre ») | quiconque connaît la **phrase de passe superviseur** |

---

## 1. Mise en ligne de l'application

Dépôt GitHub public `rondes`, publié par GitHub Pages (gratuit pour un dépôt public ; seul le programme est public) : `https://<compte>.github.io/rondes/`.

Pour une mise à jour : **Add file → Upload files**, glisser le **contenu** du dossier (pas le dossier lui-même), puis **Commit changes**. Sur chaque appareil, la nouvelle version s'installe à la deuxième ouverture avec réseau. Le numéro de version figure en bas de l'accueil.

## 2. Relais (une fois, gratuit)

À faire sur un ordinateur, avec un **compte Google dédié au service** (pas un compte personnel).

1. **script.google.com → Nouveau projet**. Effacer le contenu, coller tout le fichier `relais/relais.gs`, puis enregistrer.
2. **Déployer → Nouveau déploiement** → engrenage → **Application web**.
   - *Exécuter en tant que* : **Moi**.
   - *Qui a accès* : **Tout le monde**.
   - **Déployer → Autoriser l'accès**. Au message « Google n'a pas validé cette application » : **Paramètres avancés → Accéder à … (non sécurisé) → Autoriser**.
3. Copier l'**URL de l'application Web** (elle se termine par `/exec`).

**Mettre à jour le script du relais sans changer d'adresse** : coller le nouveau code, enregistrer, puis **Déployer → Gérer les déploiements → crayon → Version : Nouvelle version → Déployer**. Créer un *nouveau* déploiement change l'adresse ; archiver un déploiement n'efface rien.

**Repartir de zéro** : dans l'éditeur, choisir la fonction **`reinitialiserRelais`** dans la liste de la barre d'outils, puis cliquer sur **Exécuter**. Les codes d'accès sont effacés et tout le contenu du relais part à la corbeille du compte Google. Ces codes sont enregistrés dans le projet lui-même : redéployer ne suffit pas à les effacer. Réinitialiser ensuite chaque appareil (Réglages).

## 3. Premier superviseur

1. Ouvrir l'application dans Chrome (téléphone ou ordinateur), puis menu ⋮ → **Installer l'application**.
2. **Premier superviseur : activer un relais neuf** → coller l'adresse du relais → choisir la **phrase de passe superviseur** (12 caractères minimum) → **Créer et activer**.
3. Noter la phrase de passe et la remettre sous pli fermé à la DGS. Sans elle, impossible d'ajouter un appareil superviseur.
4. **Listes et coordonnées** : saisir les prénoms des agents (ils ne figurent volontairement pas dans le programme publié), vérifier destinataires et check-list, renseigner le téléphone et le courriel d'urgence.
5. **QR codes → Imprimer les QR des parkings** : une page A4 par parking, à plastifier et fixer à l'entrée.

## 4. Ajouter un appareil

- **Superviseur supplémentaire** (ordinateur de bureau, autre téléphone) : ouvrir l'application → **Ajouter cet appareil comme superviseur** → adresse du relais + phrase de passe. L'adresse figure dans les Réglages de tout appareil configuré.
- **Téléphone agents** (un ou plusieurs) : ouvrir l'application → **Scanner le QR de configuration**, affiché par un superviseur (menu **QR codes**). Autoriser la position et l'appareil photo à la première ronde. Ce QR donne accès aux signalements de l'équipe : ne pas le diffuser.

## 5. Fonctionnement courant

| Qui | Quand | Ce qui se passe |
|---|---|---|
| Agent | Arrivée au parking | Scan du QR → nom → revue des signalements en cours (« Résolu » clôt le signalement) → check-list → nouveaux signalements → « Terminer la ronde » |
| Téléphone agents | Fin de ronde, ouverture, retour du réseau, toutes les 2 min si ouvert | Envoi des saisies, réception des signalements des autres téléphones et des interventions du DPMS. Un signalement urgent part aussitôt. |
| Superviseur | À l'ouverture, puis toutes les 2 min | Relève et contrôles d'intégrité ; consultation des rondes et des signalements |
| Superviseur | Quand il le souhaite | Clôturer, rouvrir, modifier, écrire aux agents, notes internes |
| Superviseur | Périodiquement | **Archive complète** → stockage de la Ville (versement des archives) |

Le relais n'alerte pas en temps réel : pour une urgence, utiliser les boutons d'appel et de courriel du téléphone agents.

## 6. Garanties et limites

- **Chiffrement** :
  - clé d'équipe AES-256-GCM ;
  - clé superviseur ECDH P-256, HKDF-SHA-256, AES-256-GCM ;
  - coffre protégé par PBKDF2-SHA-256 (600 000 itérations).
- **Intégrité** : chaque saisie contient l'empreinte SHA-256 de la précédente du même téléphone, et chaque photo son empreinte. Une modification, une suppression ou un trou dans une série apparaît en « point d'attention ».
- **Limites** :
  - l'heure est celle du téléphone (verrouiller le réglage automatique de l'heure si la DSI gère la flotte) ;
  - le QR d'un parking peut être photographié et rejoué ;
  - en sous-sol, la position est souvent absente ou imprécise ; elle est enregistrée avec sa précision ;
  - la phrase de passe superviseur est la seule protection du coffre : longue et confidentielle.
- **Conflits** : si deux personnes modifient le même signalement, la dernière modification reçue par le relais l'emporte. L'historique complet reste dans le journal.
- **Saisies non envoyées** : elles n'existent que sur le téléphone. Ne jamais vider les données du navigateur sur un téléphone agents.

## 7. Préalables à la mise en service

- Inscription du traitement au registre des activités de traitement de la Ville (art. 30 RGPD), après avis du délégué à la protection des données. Mentionner le relais : compte Google dédié, données chiffrées sur l'appareil, conservation pendant la durée du dispositif.
- Information écrite des agents : finalités, données collectées dont la position au moment des saisies, durées de conservation, droits (art. 13 RGPD).
- Consultation du comité social territorial (art. 54 du décret n° 2021-571 du 10 mai 2021).
- Finalité affichée : traçabilité des rondes et des constats, suivi des anomalies. Pas le contrôle du temps de travail (CE, 15 décembre 2017, n° 403776).

## Bibliothèques incluses (licences libres)

- fflate (MIT) : compression.
- qrcode-generator (MIT) : génération des QR codes.
- jsQR (Apache 2.0) : lecture des QR codes si le navigateur n'a pas de lecteur intégré.
- SheetJS Community Edition 0.18.5 (Apache 2.0) : fichiers Excel.
