# ARE Fluides

Application de suivi des fluides frigorigènes d'Archenergie Renouvelable : mises en service, fiches CERFA 15497*04, rapports d'installation, stock de bouteilles et déclaration annuelle.

## Organisation du dépôt

| Dossier | Contenu |
|---|---|
| `docs/` | Appli installable publiée sur **https://fluides.toumany.com** (GitHub Pages). Elle affiche l'application Apps Script en plein écran avec l'icône ARE. |
| `apps-script/` | Code de l'application, à copier dans le projet Apps Script rattaché au Google Sheets « ARE Fluides - Base » : `Code.gs` et `Index.html` (fichier HTML nommé **Index**). |

Aucune donnée client n'est stockée ici : les données restent dans le Google Sheets et le Google Drive d'ARE, et l'accès à l'appli est protégé par les comptes et mots de passe.

## Mettre à jour l'application

1. Copier le contenu de `apps-script/Code.gs` et `apps-script/Index.html` dans l'éditeur Apps Script et enregistrer.
2. **Déployer › Gérer les déploiements › crayon › Nouvelle version › Déployer**. L'adresse `/exec` ne change pas.

L'adresse de l'application est réglée dans `docs/config.js`.

## Livraisons

- **L1** : connexion, comptes admin / opérateurs (désactivation, suppression), chantiers (création, modification, recherche).
- **Base machines** : onglet « Machines » du tableur (tous les appareils : UE, UI, modules hydrauliques, ballons ; seules les UE et monoblocs sont proposées dans l’appli). Le choix du modèle d'UE remplit marque, fluide, longueur préchargée et charge en g/m.
- **L2** (à venir) : mise en service, manipulations, photos, signature, CERFA PDF, rapport d'installation.
- **L3** (à venir) : bouteilles, détecteurs, distributeurs, déclaration annuelle.
