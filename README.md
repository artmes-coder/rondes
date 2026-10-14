# Rondes parkings — Ville de Cachan (DPMS)

Application web pour les rondes des agents de surveillance des parkings municipaux (Hénouille, Dumotel, Arobase). Version 2.7.2.

L'outil appartient aux agents : rondes, comptage des véhicules, signalements, barrières, traitement des demandes auprès des services et prestataires. La police municipale suit les signalements de véhicules. Les superviseurs consultent, interviennent ponctuellement et produisent des statistiques.

Les données transitent par un relais (compte Google dédié). Tout y est chiffré sur les appareils avant l'envoi : le relais ne lit rien. Ce dépôt ne contient que le programme.

---

## 1. Mise en ligne de l'application

Dépôt GitHub public `rondes`, publié par GitHub Pages : `https://<compte>.github.io/rondes/`.

Mise à jour : **Add file → Upload files**, glisser le **contenu** du dossier (pas le dossier lui-même), puis **Commit changes**. Chaque appareil charge la dernière version dès qu'il a du réseau (la page se recharge d'elle-même) ; sans réseau, il utilise sa copie locale. Le numéro de version figure en bas de l'accueil.

## 2. Relais (une fois)

Sur un ordinateur, avec un **compte Google dédié au service**.

1. **script.google.com → Nouveau projet**. Effacer le contenu, coller tout le fichier `relais/relais.gs`, enregistrer.
2. **Déployer → Nouveau déploiement** → engrenage → **Application web** : *Exécuter en tant que* **Moi**, *Qui a accès* **Tout le monde** → **Déployer → Autoriser l'accès** (au message « Google n'a pas validé cette application » : **Paramètres avancés → Accéder à … → Autoriser**).
3. Copier l'**URL de l'application Web** (elle se termine par `/exec`).

**Mettre à jour le relais** (sans rien perdre) : coller le nouveau code, enregistrer, puis **Déployer → Gérer les déploiements → crayon → Version : Nouvelle version → Déployer**. L'adresse ne change pas.

**Notifications (relais version 4 et plus)** : après la première mise à jour en version 4, choisir la fonction **`autoriserNotifications`** dans la barre d'outils, **Exécuter**, puis accepter l'autorisation demandée (connexion à un service externe : c'est elle qui permet d'envoyer les notifications). La version 5 ne demande pas de nouvelle autorisation.

**Alertes programmées (relais version 6)** : après la mise à jour en version 6, choisir la fonction **`installerAlertes`** dans la barre d'outils, **Exécuter**, puis accepter la nouvelle autorisation (exécution automatique toutes les 5 minutes). C'est ce déclencheur qui envoie « Aucune ronde aujourd'hui », « Ronde de plus d'1 h » et l'annonce de la synthèse du lundi, même si aucun téléphone n'est allumé. Ouvrir ensuite une fois l'application superviseur pour transmettre la programmation au relais.

**Repartir de zéro** (efface tout) : dans l'éditeur, choisir la fonction **`reinitialiserRelais`** dans la barre d'outils, puis **Exécuter**. Chaque appareil doit ensuite être réinitialisé (Réglages) et reconfiguré.

## 3. Premier superviseur

1. Ouvrir l'application dans Chrome (téléphone ou ordinateur), puis menu ⋮ → **Installer l'application**.
2. **Premier superviseur : activer un relais neuf** → adresse du relais → **phrase de passe superviseur** (12 caractères minimum, à conserver précieusement : elle sert à ajouter d'autres appareils superviseurs).
3. **Listes, catégories et coordonnées** :
   - prénoms des agents ;
   - barrières de chaque parking ;
   - destinataires avec leurs coordonnées, une ligne par destinataire : `Skidata | 01 23 45 67 89 | support@…` ;
   - catégories et sous-catégories de signalement ;
   - motifs d'ouverture de barrière ;
   - téléphone et courriel d'urgence.

## 4. Ajouter un appareil

- **Superviseur supplémentaire** : sur un appareil superviseur, **Configurer les appareils** → QR **Autre superviseur** ; le scanner avec le nouvel appareil (ou ouvrir le lien sur un PC) : l'adresse du relais est remplie d'office, il ne reste qu'à saisir la **phrase de passe superviseur** (elle n'est pas dans le QR). Sans QR : ouvrir l'application → **Ajouter cet appareil comme superviseur** → adresse du relais + phrase de passe.
- **Téléphone agents** : **Scanner un QR de configuration** → QR « Téléphones des agents » affiché par un superviseur (**Configurer les appareils**).
- **PC de la loge** (sans caméra) : sur le superviseur, **Configurer les appareils → Copier le lien**, puis ouvrir ce lien dans le navigateur du PC.
- **Police municipale** : QR « Police municipale ».

Un appareil ajouté ou réinstallé retrouve toutes les données conservées sur le relais. Les QR et liens de configuration donnent accès aux signalements de l'équipe : ne pas les diffuser.

**Réglages des téléphones agents et PM** : protégés par un mot de passe, choisi au premier accès aux réglages.

## 5. Utilisation

### Accueil des agents
Quatre grands boutons : **Commencer ma ronde** (ou **Reprendre la ronde**), **Signalement hors ronde**, **Barrière ouverte**, **Usager ou demande reçue** ; au-dessus, **Messages du DPMS** quand il y en a ; en bas, **Urgence : appeler le DPMS** et **Plus** (traitement des signalements, échanges, main courante, contrôles à faire, envoi au DPMS, notifications, extrait Excel, réglages).

### Ronde (agents) : un parcours guidé
**Commencer ma ronde** → parking → agent, puis les étapes s'enchaînent (« Étape 3 sur 6 », barre de progression) :
1. **Signalements en cours** dans le parking, un par écran, photo en grand : **Résolu** (clôt le signalement), **Toujours en cours** ou **Aggravé** ; commentaire ou photo en option.
2. **Comptage des véhicules** : « +1 véhicule » à chaque véhicule (« −1 », « Corriger »), puis **Valider**.
3. **Contrôle quotidien**, à chaque ronde : **liste déroulante des thèmes** (Rampes, Portails, Barrières et bornes, Éclairage, Extincteurs…), chacun avec sa pastille de couleur (orange : à vérifier ; vert : conforme ; rouge : non conforme ; gris : n'existe pas ici). L'agent ouvre les thèmes **dans l'ordre qu'il veut** et peut **revenir à tout moment** sur un thème pour signaler un problème vu en passant. Dans un thème : **Conforme** (le thème suivant à vérifier s'ouvre) ou **Non conforme : signaler** (signalement avec les constats propres au thème, puis retour à la liste) ; « N'existe pas dans ce parking » pour les équipements facultatifs (le superviseur peut **rétablir** le point). Le bouton du bas indique combien de thèmes restent et ouvre le premier ; une fois tout vérifié, **Contrôle terminé : étape suivante**.
4. **Contrôle mensuel**, **contrôle trimestriel** quand ils arrivent à échéance (30 et 90 jours après le précédent dans ce parking), même liste de thèmes : « Faire maintenant » ou « Plus tard » ; « Finir plus tard » en cours de route garde les réponses.
5. **Fin de ronde** : « Autre chose à signaler ? » (**+ Signalement**, **Barrière ouverte**), récapitulatif (avec « modifier », « revoir » le contrôle quotidien, « reprendre »), puis **Terminer la ronde** et l'écran « Ronde terminée ».

En tête de chaque écran : le nom du rondier, **Changer le rondier** (la ronde continue, ce qui a déjà été saisi passe au nom du bon agent ; le changement figure au journal) et **Annuler la ronde** (mauvais parking, fausse manipulation : la ronde n'apparaît pas dans les rondes ; les signalements déjà faits sont conservés).

Le contrôle trimestriel comprend la vérification des **dates d'entretien** affichées (extincteurs, portails et barrières, ascenseurs, désenfumage, éclairage de sécurité, bornes de recharge) : une maintenance non faite à temps se signale en « Entretien et maintenance », adressé au DPMS.

### Signalement : en étapes
(hors ronde : parking → qui signale) → **catégorie** → **« Quoi ? »** → **photo, obligatoire** (sauf quand une personne est en cause : « Pas de photo possible (personne présente, danger) ») → **précisions** (obligatoires pour « Autre » ; plaque, emplacement et véhicule pour les véhicules) → **à qui** (cases déjà cochées selon la sous-catégorie ou la catégorie, une ou plusieurs) et **urgent** → **Enregistrer le signalement**. « Étape précédente » à tout moment.

### Barrière ouverte (en étapes)
Parking → barrière (celles déjà signalées ouvertes sont grisées) → motif → depuis quand (« Maintenant » ou « Plus tôt (oubli de déclaration) » avec l'heure) → qui → récapitulatif, commentaire facultatif, **Enregistrer l'ouverture**. Pendant une ronde, le parking et le rondier sont déjà connus. La barrière reste affichée en rouge sur l'accueil de tous les appareils jusqu'à **Barrière refermée** (maintenant ou plus tôt → qui → **Enregistrer la fermeture**).

### Traitement des signalements (téléphone ou PC de la loge)
Accueil → **Plus** → **Traitement des signalements** : liste regroupée par étape (*À transmettre aux services*, *Relancer si pas de réponse*, *Suivre l'intervention*, *Vérifier le passage*, *À clore si le problème est réglé*). La fiche d'un signalement montre la photo, puis **Que voulez-vous faire ?** :
- **Transmettre** (ou **Relancer**) : « Appeler » et « Courriel » (message prérempli) pour chaque destinataire, « Un seul courriel à tous », « Courriel avec les photos ». Au retour dans l'application, une question demande de confirmer (« Avez-vous pu joindre… ? », « Avez-vous envoyé le courriel… ? ») : la réponse enregistre l'action, appel sans réponse compris ; un appel passé ou un courriel envoyé ramène l'agent à l'accueil. « Transmis autrement » pour une remise en main propre ;
- **Une réponse a été reçue**, **Une intervention est prévue** (date), **L'intervention a été réalisée** (l'application demande ensuite si le problème est réglé, et clôt le signalement si oui) : de qui, précisions, qui enregistre, **Enregistrer** ;
- **Problème résolu : clore le signalement** ; **Ajouter une note** ;
- **Historique** des actions et des constats, replié en bas de la fiche.

### Échanges DPMS ↔ agents
Chaque signalement a son fil d'échanges.
- **Superviseur** (fiche du signalement, « Échanges avec les agents ») : **Envoyer** (question, information) ou **Demande d'action** (consigne à exécuter) ; « Clore la demande sans réponse » si besoin. Filtre **Échanges** dans la liste des signalements ; l'accueil signale les réponses non lues.
- **Agents** : l'accueil affiche **Messages du DPMS** (icône courrier, nombre de messages non lus, demandes en attente) ; la demande apparaît aussi sur le signalement pendant la revue de ronde. L'agent choisit son nom puis **Répondre** ou **Action faite** (qui clôt la demande).
- **Notifications** : sur chaque appareil, une fois, **Activer les notifications sur cet appareil** (carte **Cet appareil** : page **Plus** des agents, haut de l'accueil superviseur), puis accepter.
  - appareils **superviseurs** : barrière ouverte et refermée ; nouveau signalement (« URGENT » en tête s'il l'est) ; signalement aggravé ou clos ; traitement (appel, courriel, réponse reçue, intervention prévue ou réalisée, note) ; suite donnée par la police municipale ; ronde terminée (véhicules, signalements) ou annulée ; réponse d'un agent ;
  - **téléphone des agents** : message ou demande d'action du DPMS ;
  - **alertes programmées** (relais version 6), sur les appareils superviseurs : « Aucune ronde aujourd'hui à … » (heure limite réglable, par défaut 20 h, fin de service des agents), « Ronde de plus d'1 h » (ronde commencée, pas terminée ni annulée une heure après), « Synthèse de la semaine » (touchez : l'application ouvre la synthèse).
  - **Choisir ses notifications** : **Réglages → Notifications reçues sur cet appareil**, type par type (urgents, autres signalements, aggravés, clos, traitement, PM, réponses des agents, barrière ouverte, barrière refermée, ronde terminée, ronde annulée, main courante, alertes programmées). Chaque appareil a ses propres choix.

  Le texte de la notification est chiffré par l'appareil qui l'envoie, pour l'appareil qui la reçoit : le relais et le service de notification ne voient qu'un bloc illisible. Les signalements faits pendant une ronde partent à la fin de la ronde (une seule notification), sauf s'ils sont urgents.

### Main courante : usagers et demandes reçues
Toute demande reçue (usager sur place, appel, courriel, service ou élu) est enregistrée et horodatée, par les agents (**Usager ou demande reçue** sur l'accueil) comme par le DPMS (**Main courante → Nouvelle entrée**).
- **Saisie en étapes** : parking (ou aucun en particulier), canal, objet (liste réglable : renseignement, réclamation, paiement ou ticket, badge d'abonné, véhicule bloqué, objet perdu, objet trouvé, incivilité, autre), **nom et courriel obligatoires** (sans courriel : téléphone obligatoire), demande, heure de réception (à l'instant ou plus tôt), qui enregistre, suite donnée.
- **Suites** : réponse donnée et demande réglée (réponse à indiquer) ; barrière ouverte pour l'usager (l'application enchaîne sur la déclaration de la barrière) ; problème à signaler (enchaîne sur le signalement, demande reprise dans les précisions, signalement lié à la fiche) ; orienté vers un autre service ; à traiter par le DPMS.
- **Fiche** : **Répondre par courriel** (courriel pré-rempli avec la référence ; au retour, l'application demande si la réponse est partie, réglée ou d'attente), **Appeler** (même confirmation, ou « pas de réponse »), **Enregistrer une réponse déjà donnée** (sur place, téléphone, courriel, courrier), note, clore sans réponse, rouvrir ; le DPMS peut annuler une saisie erronée.
- **Horodatage** : heure de réception, heure de chaque réponse et délai de la première réponse ; liste « À traiter » (les plus anciennes en premier), carte sur l'accueil superviseur, onglet **Main courante** de l'Excel complet, rubrique de la synthèse hebdomadaire (entrées par canal et par objet, délai moyen de réponse, demandes en attente), notification « Main courante » (réglable).

### Planning des agents
**Superviseur** : carte **Planning des agents** de l'accueil (mois en cours et mois suivant ; à partir du 15, « planning du mois suivant à préparer »).
- **Une fois** : **Planning-type et règles** — le cycle de deux semaines (une ligne par agent : `Nom : 7 codes semaine 1 | 7 codes semaine 2`, R = repos), le lundi d'une semaine 1, les règles contrôlées (`couverture : 08:00-20:00`, `heures par quinzaine : 70`, `jamais sur la même vacation : Nom1, Nom2`, `jamais le dimanche : Nom`, `lignes internes : Vacataire`) et les horaires des vacations.
- **Chaque mois** (semaines complètes, du lundi qui précède le 1er au dimanche qui suit le dernier jour) :
  1. **Absences connues** : agent, dates, congés / maladie / formation / absence ; ou toucher une case du tableau pour la changer.
  2. Le tableau est pré-rempli par le planning-type ; les **contrôles** s'affichent en direct (présence continue, vacations incompatibles, jours interdits, heures par quinzaine — une absence un jour normalement travaillé compte comme travaillée ; jours de repos du planning-type travaillés, pour information).
  3. Mois compliqué : **Préparer le prompt pour Claude** (avec une consigne facultative) copie un texte complet — règles, planning-type, absences, projet, problèmes, format de réponse. Le coller dans la conversation Claude consacrée au planning ; copier sa réponse ; **Appliquer la réponse de Claude** : le tableau se met à jour, les cases proposées par Claude sont encadrées en vert ; les absences saisies ne peuvent pas être modifiées par la proposition.
  4. **Publier** : les agents reçoivent une notification (« Planning de novembre publié », puis, à chaque modification, la liste des changements). Les cases modifiées depuis la dernière publication sont encadrées en orange. Export **Excel**.
- **Rappel** : notification le 15 de chaque mois à 9 h, « Planning de … à préparer » (relais version 6), sauf si le mois suivant est déjà publié.
- **Confidentialité** : les motifs d'absence et les lignes internes ne sont visibles que des superviseurs ; les agents voient « Absent ».

**Agents** : carte **Planning** sur l'accueil (qui est du matin, du midi, du soir aujourd'hui ; « mis à jour » après une publication) ; onglet Planning par semaine, ou mois par agent ; notes du DPMS.

### Police municipale
Signalements de véhicules (ou tout ce qui est adressé à la PM) avec plaque, emplacement, photos et constats. Observations, puis **Traité** (clôt le signalement) ou **Non traité** (le laisse ouvert, observation visible des agents). Export Excel.

### Superviseur
- **Signalements** : sur PC, liste et fiche côte à côte ; filtres par statut, parking, catégorie, recherche (plaque, référence, texte) ; photos en galerie, agrandies d'un clic (flèches du clavier, Échap) ; traitement, constats, suivi PM ; interventions facultatives (clore, rouvrir, modifier, écrire aux agents, notes internes).
- **Clôturer une ronde oubliée** : bouton **Clôturer la ronde** sur toute ronde sans fin (carte « Rondes du jour » au-delà d'1 h, liste **Toutes les rondes**). Le téléphone qui porte la ronde l'arrête à sa prochaine connexion, envoie le contrôle quotidien déjà fait (marqué « partiel », il ne compte pas comme contrôle du jour) et affiche « Ronde clôturée par le DPMS ». Si l'agent avait terminé sa ronde hors réseau, sa propre fin de ronde prime. La ronde est marquée « clôturée par le DPMS » ; sa durée n'entre pas dans les moyennes.
- **Rondes du jour** (haut de l'accueil) : une ronde attendue par jour et par parking ; heure de la ronde pour chaque parking, « Aucune ronde aujourd'hui à … » passée l'heure limite (20 h par défaut), rondes ouvertes depuis plus d'1 h. Heure limite et jours concernés : **Listes, catégories et coordonnées**. Une ronde annulée ne compte pas.
- **Synthèse hebdomadaire (PDF)** (Exports, ou touchez la notification du lundi) : semaine dernière ou en cours ; rondes par parking (nombre, durée et véhicules moyens, rondes non terminées, jours sans ronde), contrôles (quotidiens, mensuel, trimestriel, registres, points non conformes), signalements (nouveaux par catégorie, urgents ouverts, en cours par état, ouverts depuis plus de 30 jours, demandes d'action sans réponse), main courante (entrées, délai moyen de réponse, demandes en attente), barrières (ouvertures, durée cumulée, encore ouvertes). Jour et heure de l'annonce réglables. Le PDF est fabriqué sur l'appareil du superviseur : le relais ne peut pas lire les données, il ne l'envoie donc pas par courriel.
- **Statistiques et graphiques** :
  - période (7 jours à 12 mois, tout, ou dates) et parking ;
  - indicateurs clés : signalements, en cours, clos, délai moyen de clôture, rondes, contrôles réalisés et non-conformités, véhicules par ronde, ouvertures et durée cumulée des barrières ;
  - graphiques prêts, et **Créer un graphique** : un indicateur (signalements, en cours / clos, délai de clôture, actions de traitement, contrôles réalisés, non-conformités, rondes, véhicules comptés, ouvertures ou durée d'ouverture des barrières) par catégorie, sous-catégorie, point de contrôle, type de contrôle, agent, parking, destinataire, barrière, motif, jour, semaine ou mois ;
  - chaque graphique : vue tableau, image, Excel.
- **Contrôles** : par parking, dernier contrôle quotidien, mensuel et trimestriel, **registre de sécurité** (contrôle mensuel par le DPMS : rappel sur l'accueil après 30 jours, bouton « Registre contrôlé » avec observations), échéances et retards, points « n'existe pas ici » (avec « Rétablir »), historique avec les non-conformités.
- **Exports** : synthèse hebdomadaire (PDF), Excel complet (Signalements, Rondes, Traitement, Contrôles, Barrières, Main courante, Journal) et archive avec photos.
- **Suppression** : un signalement (bouton dans sa fiche), tous les signalements affichés par les filtres, une ronde ou toutes les rondes (**Toutes les rondes**). Les éléments supprimés disparaissent de tous les appareils, des statistiques et des exports ; ils restent restaurables depuis **Corbeille** (Paramétrage).

### Points de contrôle et catégories
Modifiables par le superviseur, sans toucher au programme (**Listes, catégories et coordonnées**) :
```
[Mensuel]
- Bac d’absorbant : rempli, pelle présente | Sécurité incendie     ← point : ce qu'on regarde | catégorie en cas de non-conformité
- ?Bornes de recharge : en service, câbles intacts | Bornes de …   ← « ? » : l'agent peut indiquer qu'il n'existe pas ici
- ?Portails : fonctionnent | Accès et circulation | Portail en panne ; Bouton-poussoir désactivé ; Portail endommagé   ← 3e champ : les « Quoi ? » proposés quand le point est non conforme (« ! » = urgent)
```
```
Sécurité incendie | Ateliers          ← catégorie de signalement | destinataires proposés
- Extincteur absent ou déplacé        ← sous-catégorie
- Odeur de fumée ou de brûlé !        ← « ! » : urgent par défaut
- Tags ou graffitis | Ateliers, DPMS, Police municipale   ← destinataires propres à la sous-catégorie
```

## 6. Bon à savoir

- **Mises à jour** : chaque appareil vérifie la présence d'une nouvelle version à chaque retour sur l'application et toutes les 15 minutes, et l'installe dès que l'écran ne contient aucune saisie en cours (jamais au milieu d'un signalement). La version figure en bas de l'accueil. Un appareil qui reçoit une information d'une version plus récente affiche « Mettre à jour maintenant ».

- Ne jamais vider les données du navigateur sur un téléphone agents : les saisies pas encore envoyées n'existent que sur lui.
- **Synchronisation** : chaque appareil n'envoie que ses propres saisies, une seule fois ; il ne renvoie jamais ce qu'il a reçu. Plusieurs téléphones agents peuvent être configurés avec le même QR. Les décisions prises pendant une relève attendent leur tour dans la file de l'appareil ; une décision reçue deux fois (renvoi après une réponse perdue du relais) n'est appliquée qu'une fois. Une suppression faite par le DPMS s'applique sur tous les appareils à leur relève ; les saisies d'agents qui visent un signalement supprimé sont ignorées (ni affichées, ni notifiées, ni exportées).
- **Qui l'emporte** : un constat ou une clôture d'agent ne change le statut d'un signalement que si le téléphone connaissait la dernière décision (clôture ou réouverture) au moment de la saisie. Un constat fait hors réseau sur un signalement que le DPMS a clos entre-temps est conservé dans l'historique (« sans effet sur le statut ») et signalé dans les points d'attention, mais ne le rouvre pas. Le téléphone tente une relève au début de chaque ronde.
- Pour les autres modifications d'un même signalement, la dernière reçue l'emporte ; l'historique reste dans le journal.
- Les notifications et alertes ne remplacent pas l'appel : pour une urgence, utiliser les boutons d'appel de l'accueil.

## Bibliothèques incluses (licences libres)

fflate (MIT), qrcode-generator (MIT), jsQR (Apache 2.0), SheetJS Community Edition 0.18.5 (Apache 2.0), Chart.js 4.5 (MIT), jsPDF 4.2.1 (MIT).
