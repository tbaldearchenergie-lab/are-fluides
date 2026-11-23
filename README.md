# ARE Fluides

Application de suivi des fluides frigorigènes d'Archenergie Renouvelable : mises en service, fiches CERFA 15497*04, rapports d'installation, stock de bouteilles et déclaration annuelle.

## Organisation du dépôt

| Dossier | Contenu |
|---|---|
| `docs/` | Appli installable publiée sur **https://fluides.toumany.com** (GitHub Pages). Elle affiche l'application Apps Script en plein écran avec l'icône ARE. |
| `apps-script/` | Code de l'application, à copier dans le projet Apps Script rattaché au Google Sheets « ARE Fluides - Base » : `Code.gs` et `Index.html` (fichier HTML nommé **Index**). |

Aucune donnée client n'est stockée ici : les données restent dans le Google Sheets et le Google Drive d'ARE, et l'accès à l'appli est protégé par les comptes et mots de passe.

## Autoriser le Drive (une fois, depuis la livraison 2)

Dans l'éditeur Apps Script, choisir la fonction `autoriserDrive` puis ▶ Exécuter et accepter l'accès au Drive. Le journal affiche « Drive opérationnel ».

## Mettre à jour l'application

1. Copier le contenu de `apps-script/Code.gs` et `apps-script/Index.html` dans l'éditeur Apps Script et enregistrer.
2. **Déployer › Gérer les déploiements › crayon › Nouvelle version › Déployer**. L'adresse `/exec` ne change pas.

L'adresse de l'application est réglée dans `docs/config.js`.

## Livraisons

- **L1** : connexion, comptes admin / opérateurs (désactivation, suppression), chantiers (création, modification, recherche).
- **Base machines** : onglet « Machines » du tableur (tous les appareils : UE, UI, modules hydrauliques, ballons ; seules les UE et monoblocs sont proposées dans l’appli). Le choix du modèle d'UE remplit marque, fluide, longueur préchargée et charge en g/m.
- **L2** : mise en service en 5 étapes (équipement, opération, fluide, étanchéité, photos + signature), lecture du n° de série sur la photo de la plaque, CERFA 15497*04 rempli et signé (modèle vierge : `docs/cerfa_15497-04.pdf`), rapport d'installation PDF, rangement dans le Drive (ARE Fluides › Chantiers), registre des mouvements et pesée des bouteilles. Droits : l'administrateur modifie ou supprime toute mise en service ; un opérateur seulement celles de ses propres chantiers. Un chantier ne se supprime qu'après suppression de sa mise en service.
- **L3** : bouteilles (ajout d'une bouteille neuve, de récupération ou de transfert, fiche avec historique, statut au dépôt / rendue au distributeur avec date et BSFF, correction des pesées pour initialiser le stock existant), profil (modifier mon profil, mot de passe, dépôt de l'attestation dans le Drive › ARE Fluides › Opérateurs, jamais affichée dans l'appli), et côté administrateur : détecteurs de fuite (contrôle annuel), distributeurs (case 13 du CERFA) et déclaration annuelle calculée depuis le registre et les pesées, avec export PDF (Drive › ARE Fluides › Déclarations). Droits L2 revus : un opérateur peut ajouter des manipulations et compléter les n° de série absents sur le chantier d'un autre, sans autre modification.
- **L4** (à venir) :
  - **CERFA papier** : photographier un CERFA rempli à la main ; l'appli lit client, adresse, date et manipulation, propose les chantiers qui correspondent (ou en crée un), enregistre la manipulation avec le n° du CERFA. Pour un ajout ou une récupération, elle demande la bouteille utilisée (bouteilles du bon fluide) et prend la date écrite sur le CERFA.
  - **Règle de date des bouteilles** : une manipulation ne peut pas être datée d'avant la dernière opération déjà enregistrée sur la même bouteille ; sinon refus et demande d'une autre date (saisie dans l'appli comme CERFA papier).
  - **Pesée de contrôle** (par exemple chaque trimestre ou semestre) : comparaison avec le poids théorique, signalement de l'écart, régularisation en choisissant un chantier existant ou nouveau et en y ajoutant la manipulation qui explique l'écart.
  - **Responsable de chantier** : l'admin crée un chantier et l'attribue à un opérateur (ou à lui-même) ; le responsable le voit dans ses chantiers et ses chantiers à mettre en service, avec les droits du créateur ; seul l'admin peut changer de responsable.
  - **Alertes d'échéance** : sur l'accueil (contrôle de détecteur à venir, bouteille de récupération presque pleine) et par e-mail automatique pour les bouteilles et le contrôle des détecteurs.
  - **Export du registre** : manipulations d'une période (date, chantier, machine, fluide, quantité, bouteille, opérateur, n° CERFA) en Excel et en PDF.
  - **Machines à vérifier dans l'appli** : liste des ajouts et propositions des opérateurs avec Accepter (base mise à jour) ou Refuser, sans passer par le Google Sheets.
  - **Brouillon de mise en service** : la saisie en cours (photos comprises) est gardée sur le téléphone si le réseau coupe ou si l'appli est fermée, et reprise là où on s'était arrêté.
  - **Envoi des documents au client** (option activable / désactivable par l'admin dans son profil) : bouton « Envoyer au client » sur la fiche chantier, qui envoie le rapport d'installation et le CERFA par e-mail depuis tbald.archenergie@gmail.com ; e-mail client facultatif sur la fiche chantier.
