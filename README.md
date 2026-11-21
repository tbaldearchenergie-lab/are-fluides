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
- **L3** (à venir) : bouteilles, détecteurs, distributeurs, déclaration annuelle.
