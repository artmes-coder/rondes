# Rondes parkings — Ville de Cachan (DPMS)

Application web pour les rondes des agents de surveillance des parkings municipaux (Hénouille, Dumotel, Arobase). Version 1.5.

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

- **Superviseur supplémentaire** : ouvrir l'application → **Ajouter cet appareil comme superviseur** → adresse du relais + phrase de passe.
- **Téléphone agents** : **Scanner un QR de configuration** → QR « Téléphones des agents » affiché par un superviseur (**Configurer les appareils**).
- **PC de la loge** (sans caméra) : sur le superviseur, **Configurer les appareils → Copier le lien**, puis ouvrir ce lien dans le navigateur du PC.
- **Police municipale** : QR « Police municipale ».

Un appareil ajouté ou réinstallé retrouve toutes les données conservées sur le relais. Les QR et liens de configuration donnent accès aux signalements de l'équipe : ne pas les diffuser.

**Réglages des téléphones agents et PM** : protégés par un mot de passe, choisi au premier accès aux réglages.

## 5. Utilisation

### Ronde (agents)
**Commencer ma ronde** → parking → agent → revue des signalements en cours (« Résolu » clôt le signalement) → écran de ronde :
- **comptage des véhicules** : « +1 véhicule » à chaque véhicule, « −1 » et « Corriger » en cas d'erreur, puis **Valider le comptage** ; le compteur se replie (« modifier » pour le rouvrir) ;
- **contrôle quotidien** (à la première ronde du jour dans le parking) : pour chaque point, « Conforme » ou « Non conforme » ; « Non conforme » ouvre un signalement prérempli, autant de fois que nécessaire ;
- **contrôles mensuel et trimestriel** : proposés pendant la ronde quand ils arrivent à échéance (30 et 90 jours après le précédent dans ce parking) ; « Plus tard » garde les réponses déjà données ;
- **N'existe pas ici** (points facultatifs : caisses, ascenseurs, désenfumage, interphones, bornes de recharge, colonnes sèches, locaux techniques) : le point n'est plus proposé dans ce parking ; le superviseur peut le **rétablir** ;
- **+ Signalement** pour tout autre constat ;
- **Terminer la ronde** : une fois le comptage validé et le contrôle quotidien complet.

L'accueil indique les **contrôles à faire** par parking.

### Signalement
Catégorie → « Quoi ? » (sous-catégorie ou « Autre (préciser) ») → précisions, photos, destinataires (proposés selon la catégorie), urgence. Véhicules : plaque, emplacement (numéro de place, niveau), marque/modèle/couleur. **Signalement hors ronde** depuis l'accueil.

### Barrière ouverte
Parking, barrière, motif, heure (« Maintenant » ou « Plus tôt » en cas d'oubli). La barrière reste affichée en rouge sur l'accueil de tous les appareils jusqu'à **Barrière refermée** (maintenant ou plus tôt).

### Traitement des signalements (téléphone ou PC de la loge)
Accueil → **Traitement des signalements** : signalements classés en *À transmettre*, *En attente de réponse*, *Réponse reçue*, *Intervention prévue*, *Intervention réalisée*. Dans chaque fiche :
- **Transmettre** : boutons « Appeler » et « Courriel » (message prérempli) pour chaque destinataire, et « Partager avec les photos » (WhatsApp, messagerie…) ;
- **Ajouter une action** : appel, courriel, relance, réponse reçue, intervention prévue (date), intervention réalisée, note, avec le destinataire, le détail et l'agent ;
- **Problème résolu : clore le signalement**.

### Police municipale
Signalements de véhicules (ou tout ce qui est adressé à la PM) avec plaque, emplacement, photos et constats. Observations, puis **Traité** (clôt le signalement) ou **Non traité** (le laisse ouvert, observation visible des agents). Export Excel.

### Superviseur
- **Signalements** : sur PC, liste et fiche côte à côte ; filtres par statut, parking, catégorie, recherche (plaque, référence, texte) ; photos en galerie, agrandies d'un clic (flèches du clavier, Échap) ; traitement, constats, suivi PM ; interventions facultatives (clore, rouvrir, modifier, écrire aux agents, notes internes).
- **Statistiques et graphiques** :
  - période (7 jours à 12 mois, tout, ou dates) et parking ;
  - indicateurs clés : signalements, en cours, clos, délai moyen de clôture, rondes, contrôles réalisés et non-conformités, véhicules par ronde, ouvertures et durée cumulée des barrières ;
  - graphiques prêts, et **Créer un graphique** : un indicateur (signalements, en cours / clos, délai de clôture, actions de traitement, contrôles réalisés, non-conformités, rondes, véhicules comptés, ouvertures ou durée d'ouverture des barrières) par catégorie, sous-catégorie, point de contrôle, type de contrôle, agent, parking, destinataire, barrière, motif, jour, semaine ou mois ;
  - chaque graphique : vue tableau, image, Excel.
- **Contrôles** : par parking, dernier contrôle quotidien, mensuel et trimestriel, échéances et retards, points « n'existe pas ici » (avec « Rétablir »), historique avec les non-conformités.
- **Exports** : Excel complet (Signalements, Rondes, Traitement, Contrôles, Barrières, Journal) et archive avec photos.
- **Suppression** : un signalement (bouton dans sa fiche), tous les signalements affichés par les filtres, une ronde ou toutes les rondes (**Toutes les rondes**). Les éléments supprimés disparaissent de tous les appareils, des statistiques et des exports ; ils restent restaurables depuis **Corbeille** (Paramétrage).

### Points de contrôle et catégories
Modifiables par le superviseur, sans toucher au programme (**Listes, catégories et coordonnées**) :
```
[Mensuel]
- Bac d’absorbant : rempli, pelle présente | Sécurité incendie     ← point : ce qu'on regarde | catégorie en cas de non-conformité
- ?Bornes de recharge : en service, câbles intacts | Bornes de …   ← « ? » : l'agent peut indiquer qu'il n'existe pas ici
```
```
Sécurité incendie | Ateliers          ← catégorie de signalement | destinataires proposés
- Extincteur absent ou déplacé        ← sous-catégorie
- Odeur de fumée ou de brûlé !        ← « ! » : urgent par défaut
```

## 6. Bon à savoir

- Ne jamais vider les données du navigateur sur un téléphone agents : les saisies pas encore envoyées n'existent que sur lui.
- Si deux personnes modifient le même signalement, la dernière modification reçue l'emporte ; l'historique reste dans le journal.
- Le relais n'alerte pas en temps réel : pour une urgence, utiliser les boutons d'appel de l'accueil.

## Bibliothèques incluses (licences libres)

fflate (MIT), qrcode-generator (MIT), jsQR (Apache 2.0), SheetJS Community Edition 0.18.5 (Apache 2.0), Chart.js 4.5 (MIT).
