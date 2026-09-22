# ClassifyMe

ClassifyMe est un MVP d'Office Add-in pour Word, PowerPoint, Excel et Outlook en mode
composition. Il permet à l'utilisateur de choisir manuellement un niveau `PUBLIC`,
`RESTREINT`, `CONFIDENTIEL` ou `SECRET`, d'appliquer un marquage visible et de stocker des
métadonnées limitées.

Le MVP ne chiffre pas les fichiers ou les e-mails, ne bloque pas l'enregistrement ou
l'envoi, n'analyse pas le contenu et ne remplace pas Microsoft Purview.

## Objectif actuel

Le projet couvre actuellement :

- Word : bandeau de classification dans l'en-tête du document ;
- PowerPoint : pied de page de classification sur les diapositives existantes ;
- Excel : forme de classification sur les feuilles existantes et pied de page pour
  impression/PDF ;
- Outlook : bandeau HTML idempotent dans un e-mail ou une réunion organisée, en mode
  composition ;
- Outlook : rappel automatique non bloquant au début d'une composition ;
- un panneau latéral commun permettant de choisir et de modifier le niveau ;
- une option Outlook, désactivée par défaut, permettant d'ajouter un préfixe de
  classification à l'objet.

## Cadre technique

- Office Add-in généré avec Yeoman ;
- Office.js et TypeScript ;
- interface HTML/CSS simple ;
- exécution locale avec Node.js et npm ;
- aucun backend, base de données, Microsoft Graph ou mécanisme d'authentification.

## Organisation du code

```text
src/
  core/
    classificationConstants.ts
  hosts/
    office/
      wordClassification.ts
      excelClassification.ts
      powerpointClassification.ts
      officeRouter.ts
    outlook/
      outlookClassification.ts
      outlookClassificationReminder.ts
      outlookClassificationReminder.html
  ui/
    taskpane/
      taskpane.ts
      taskpane.html
      taskpane.css
```

`manifest.xml` cible Word, Excel et PowerPoint. `manifest.outlook.xml` cible Outlook en
développement. Les variantes `manifest.office.production.xml` et
`manifest.outlook.production.xml` utilisent les ressources HTTPS de production.

## Installation et commandes utiles

Installer les dépendances :

```powershell
npm install
```

Compiler et contrôler le projet :

```powershell
npm run build:dev
npm run build
npm run lint
npm run validate
npm run validate:outlook
npm run validate:office:production
npm run validate:outlook:production
```

Lancer ou arrêter les compléments locaux :

```powershell
npm start
npm run start:word
npm run start:powerpoint
npm run start:excel
npm run start:outlook
npm stop
npm run stop:outlook
```

Le poste peut demander l'approbation du certificat HTTPS de développement local.

## Rappel de classification Outlook

Lorsqu'une composition Outlook commence, ClassifyMe affiche automatiquement une
notification de rappel. Le mécanisme utilise :

- `OnNewMessageCompose` pour les nouveaux messages, réponses, réponses à tous et
  transferts ;
- `OnNewAppointmentOrganizer` pour les nouvelles réunions et nouveaux rendez-vous
  organisés ;
- un `InsightMessage` avec l'action `Ouvrir ClassifyMe` ;
- `Office.actions.associate()` et un appel garanti à `event.completed()`.

Le rappel distingue le message de la réunion et ne choisit aucun niveau. Il ne modifie ni
le contenu, ni l'objet, ni les métadonnées et ne bloque aucune action. Un clic sur
`Ouvrir ClassifyMe` ouvre le panneau existant ; le bouton manuel du ruban reste disponible.

La clé `classifyme-reminder` est stable et respecte la limite Outlook de 32 caractères :
un nouvel ajout remplace le rappel portant la même clé au lieu d'afficher plusieurs
notifications identiques.

L'ouverture entièrement automatique du panneau n'est pas utilisée :
`Office.addin.showAsTaskpane()` exige un runtime partagé, non pris en charge par Outlook.
Le rapport [RAPPORT-SPIKE-AUTOOPEN-OUTLOOK.md](RAPPORT-SPIKE-AUTOOPEN-OUTLOOK.md) conserve
l'historique de cette décision et des diagnostics ayant conduit à l'implémentation.

### Runtime événementiel

Webpack produit deux compilateurs cohérents dans le même dossier de sortie :

- le compilateur de l'application génère le task pane et les commandes avec le
  comportement standard de `webpack-dev-server` en développement ;
- le compilateur `outlook-classification-reminder` génère un petit bundle autonome ; en
  mode DEV, il est précompilé sur disque puis servi comme ressource statique avec
  `devServer: false`, ce qui l'exclut de l'injection du client WebSocket et du HMR.

Outlook Classic charge directement le fichier JavaScript. Outlook Web et le nouvel
Outlook chargent la page HTML, qui initialise Office.js avec `Office.initialize` avant le
bundle. Le script est chargé de manière bloquante afin d'associer les handlers assez tôt.
Le hash du bundle et la version des URL du manifeste limitent la réutilisation de
ressources obsolètes par le cache Outlook.

### Compatibilité validée pendant le spike

| Scénario | Outlook Classic | Web / New Outlook |
| --- | ---: | ---: |
| Nouveau message | OK | OK |
| Réponse | OK | OK |
| Répondre à tous | OK | OK |
| Transfert | OK | OK |
| Nouvelle réunion | OK | OK |
| Ouverture manuelle | OK | OK |

Cette matrice décrit les validations fonctionnelles réalisées avant l'industrialisation.
Le protocole ci-dessous doit être rejoué après déploiement ou sideload de la nouvelle
version ; un build réussi ne remplace pas ce contrôle humain.

### Validation manuelle après industrialisation

Dans Outlook Classic, puis dans Outlook Web ou le nouvel Outlook, tester successivement :

1. un nouveau message ;
2. une réponse ;
3. une réponse à tous ;
4. un transfert ;
5. une nouvelle réunion organisée ;
6. le bouton manuel ClassifyMe.

Pour chaque scénario, vérifier :

1. l'apparition d'un seul rappel ;
2. le texte français adapté au message ou à la réunion ;
3. l'ouverture du panneau après le clic sur `Ouvrir ClassifyMe` ;
4. le fonctionnement normal de la classification ;
5. l'absence de modification automatique du contenu et de l'objet ;
6. l'absence de blocage d'Outlook et de l'envoi.

## Validation manuelle de la classification

### Word

1. Lancer `npm run start:word` et ouvrir le panneau ClassifyMe.
2. Cliquer sur chaque niveau et vérifier le niveau affiché dans le panneau.
3. Vérifier que le bandeau de l'en-tête est créé puis remplacé sans doublon.

### PowerPoint

1. Lancer `npm run start:powerpoint` avec une présentation de plusieurs diapositives.
2. Appliquer deux niveaux successifs.
3. Vérifier que chaque diapositive existante contient un seul pied de page à jour.
4. Ajouter une diapositive et vérifier qu'elle n'est pas marquée automatiquement.

### Excel

1. Lancer `npm run start:excel` avec un classeur de plusieurs feuilles.
2. Appliquer deux niveaux successifs.
3. Vérifier qu'une seule forme `ClassifyMeBanner` à jour est visible par feuille.
4. Vérifier le pied de page dans l'aperçu avant impression ou dans un export PDF.
5. Ajouter une feuille et vérifier qu'elle n'est pas marquée automatiquement.

### Outlook

1. Lancer `npm run start:outlook` et composer un e-mail ou une réunion organisée.
2. Ouvrir ClassifyMe depuis le rappel ou depuis le bouton manuel.
3. Appliquer successivement plusieurs niveaux et vérifier qu'un seul bandeau reste présent.
4. Vérifier que la signature, les citations et le contenu existant sont conservés.
5. Activer l'option de préfixe, appliquer deux niveaux et vérifier que le préfixe est
   remplacé sans doublon.
6. Désactiver l'option et vérifier le comportement documenté pour l'e-mail et la réunion.

## Publication GitHub Pages et Microsoft 365

Les manifestes de production publiés utilisent l'URL ClassifyMe existante :

```text
https://MindGroup29.github.io/ClassifyMe/
```

Pour publier sur une autre URL, définir explicitement :

```powershell
$env:CLASSIFYME_PRODUCTION_BASE_URL="https://<org>.github.io/<repo>/"
npm run build:github-pages
```

La commande génère le dossier `docs`, qui doit notamment contenir :

```text
docs/
  assets/
  commands.html
  taskpane.html
  outlook-classification-reminder.html
  outlook-classification-reminder.js
  manifest.office.production.xml
  manifest.outlook.production.xml
```

Avant un déploiement Microsoft 365 :

1. vérifier que toutes les ressources sont accessibles en HTTPS ;
2. vérifier qu'aucun manifeste de production ne contient `localhost` ;
3. valider les deux manifestes de production ;
4. sideloader les manifestes et rejouer les validations manuelles ;
5. déployer d'abord les compléments à un groupe pilote depuis Microsoft 365 Admin Center.

GitHub Pages ne fournit que des ressources statiques. Un changement d'URL impose une
reconstruction des manifestes publiés. Le cache Office peut conserver une ancienne
version pendant quelques minutes.

## Ce qui fonctionne

- Les quatre niveaux sont disponibles dans un panneau commun et s'appliquent au clic.
- Word met à jour un bandeau et tente de stocker les propriétés personnalisées prévues.
- PowerPoint remplace les pieds de page ClassifyMe sur les diapositives existantes.
- Excel remplace le bandeau et le pied de page sur les feuilles existantes et tente de
  stocker les propriétés personnalisées prévues.
- Outlook remplace les bandeaux identifiés, y compris les identifiants préfixés par `x_`
  dans Outlook Web, et conserve le reste du corps HTML.
- Outlook tente de stocker les propriétés personnalisées ; l'échec de cette opération ne
  retire pas le bandeau visible.
- Les manifestes Outlook DEV et PROD proposent les commandes manuelles pour les messages
  et réunions ainsi que les deux événements de rappel.
- Les builds DEV et PROD incluent le runtime événementiel autonome.

## Limites connues du MVP

- Les nouvelles diapositives et feuilles ajoutées après classification ne sont pas
  marquées automatiquement.
- PowerPoint ne modifie pas les masques et ne stocke pas encore de métadonnées
  personnalisées.
- Le panneau ne relit pas automatiquement une classification existante à son ouverture.
- Outlook prend en charge uniquement la composition d'e-mails et de réunions organisées ;
  le mode lecture et les invitations reçues ne sont pas implémentés.
- Les propriétés personnalisées Outlook dépendent du client, du réseau et du type de
  compte, et ne sont pas transmises aux destinataires.
- Un élément Outlook en texte brut peut refuser l'insertion du bandeau HTML.
- Certaines surfaces Outlook non standard peuvent ne pas déclencher l'activation
  événementielle.
- Le rappel Outlook nécessite un clic pour ouvrir le panneau et dépend d'une connexion.
- Le MVP ne chiffre pas, ne bloque pas, n'analyse pas et ne produit aucun reporting.

## Prochaines étapes recommandées

1. Rejouer le protocole Outlook complet après sideload des manifestes industrialisés.
2. Valider les ressources de production sur l'URL HTTPS réelle avec un groupe pilote.
3. Vérifier manuellement les quatre niveaux dans Word, PowerPoint et Excel.
4. Décider ultérieurement si PowerPoint doit utiliser les masques de diapositives.
5. Décider si la lecture d'une classification existante doit être ajoutée au panneau.

## Hors périmètre

- ouverture automatique du task pane sans clic ;
- blocage avant envoi ou contrôle DLP ;
- classification automatique ou suggestion par IA ;
- chiffrement ou restriction d'impression/transfert ;
- Microsoft Graph, backend, base de données ou reporting ;
- authentification Entra ID ou intégration Microsoft Purview.
