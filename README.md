# ClassifyMe

Add-in Office MVP permettant aux utilisateurs de marquer manuellement leurs documents et emails selon un niveau de confidentialité.

## Objectif

Ce projet vise à fournir une alternative simple et peu coûteuse aux étiquettes de sensibilité Microsoft Purview pour une entreprise utilisant Microsoft 365 Business Standard.

L'add-in ne chiffre pas les documents, ne bloque pas les partages et ne remplace pas Microsoft Purview.

Il permet uniquement de :

- choisir un niveau de classification ;
- ajouter un marquage visuel ;
- stocker une métadonnée simple ;
- responsabiliser l'utilisateur.

## Niveaux de classification

Les niveaux cibles sont définis dans `classification-rules.md`.

Codes techniques :

- `PUBLIC`
- `RESTREINT`
- `CONFIDENTIEL`
- `SECRET`

## Cadre technologique

Le projet doit être développé comme un Office Add-in basé sur :

- Office.js ;
- TypeScript ;
- Node.js ;
- npm ;
- Yeoman Generator for Office Add-ins ;
- Visual Studio Code.

Le générateur Yeoman officiel Microsoft permet de créer des projets Office Add-ins basés sur Node.js. Microsoft recommande Visual Studio si l'on veut du code serveur .NET ou un hébergement IIS, ce qui n'est pas le cas ici.

Documentation Microsoft utile :

- Yeoman Generator for Office Add-ins  
  https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview

- Développement des Office Add-ins  
  https://learn.microsoft.com/en-us/office/dev/add-ins/develop/develop-overview

- Déploiement centralisé des Office Add-ins  
  https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins

## Pré-requis Windows

Installer les outils suivants :

### 1. Visual Studio Code

Télécharger et installer :

https://code.visualstudio.com/

Extensions recommandées :

- ESLint
- Prettier
- Microsoft 365 Agents Toolkit, optionnel

### 2. Node.js LTS

Télécharger et installer la version LTS :

https://nodejs.org/

Vérifier l'installation :

```powershell
node --version
npm --version
````

### 3. Git

Télécharger et installer :

[https://git-scm.com/](https://git-scm.com/)

Vérifier l'installation :

```powershell
git --version
```

### 4. Yeoman et générateur Office

Installer Yeoman et le générateur Office :

```powershell
npm install -g yo generator-office
```

Vérifier :

```powershell
yo --version
```

## Création initiale du projet

Depuis le dossier parent du projet :

```powershell
yo office
```

Choix recommandés pour le premier MVP :

```text
Project type: Office Add-in Task Pane project
Script type: TypeScript
Office client application: Word
Project name: ClassifyMe
```

Pourquoi commencer par Word :

* API plus simple qu'Outlook ;
* validation rapide du geste utilisateur ;
* bon support pour les documents bureautiques ;
* extension plus facile ensuite vers Excel et PowerPoint.

## Installation des dépendances

Dans le dossier du projet :

```powershell
npm install
```

## Lancement en développement

Pour lancer l'add-in en local :

```powershell
npm start
```

Le modèle Yeoman configure généralement un serveur local HTTPS et ouvre l'application Office cible.

Selon la configuration du poste, il peut être nécessaire d'accepter un certificat de développement local.

## Arrêt du serveur local

```powershell
npm stop
```

ou fermer le terminal actif avec `Ctrl + C`.

## Structure documentaire du dépôt

Le dépôt doit contenir au minimum :

```text
classification-rules.md
AGENTS.md
README.md
```

Après génération Yeoman, le projet contiendra également les fichiers techniques de l'add-in.

## Règles de développement

Avant toute modification :

1. lire `classification-rules.md` ;
2. respecter le périmètre MVP ;
3. éviter toute fonctionnalité avancée non demandée ;
4. mettre à jour ce README.

## État actuel

À compléter après génération du projet.

Exemple :

```text
- Projet Yeoman généré : non
- Add-in Word fonctionnel : non
- Bandeau de classification : non
- Propriété documentaire personnalisée : non
- Outlook : non démarré
```

## Vérification manuelle attendue

Lorsque le premier prototype Word sera développé :

```text
1. Lancer npm start.
2. Ouvrir Word.
3. Afficher le panneau de l'add-in.
4. Choisir un niveau de classification.
5. Cliquer sur Appliquer.
6. Vérifier que le bandeau apparaît dans le document.
7. Vérifier que le niveau affiché dans l'add-in correspond au choix utilisateur.
8. Enregistrer puis rouvrir le document.
9. Vérifier que la classification reste disponible.
```

## Hors périmètre

Le MVP ne doit pas inclure :

* chiffrement ;
* DLP ;
* restriction d'impression ;
* restriction de transfert ;
* IA ;
* reporting centralisé ;
* base de données ;
* API Graph ;
* authentification Entra ID ;
* intégration Microsoft Purview.

## Déploiement cible

Le déploiement final pourra être réalisé via le centre d'administration Microsoft 365, dans :

```text
Settings > Integrated apps
```

Le déploiement centralisé permet d'attribuer un add-in à des utilisateurs ou groupes Microsoft 365.

Cette étape n'est pas nécessaire pour le développement local.