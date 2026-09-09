# ClassifyMe

ClassifyMe est un MVP d'Office Add-in pour Microsoft Word, Microsoft PowerPoint, Microsoft Excel et Outlook en mode composition. Il permet a un utilisateur de choisir manuellement un niveau de classification, d'ajouter un marquage visible et de garder l'interface volontairement simple.

Le MVP ne chiffre pas les fichiers ou emails, ne bloque pas l'enregistrement ou l'envoi, n'analyse pas le contenu et ne remplace pas Microsoft Purview.

## Objectif actuel

Le projet couvre actuellement :

- Word : application d'un bandeau de classification dans l'en-tete du document ;
- PowerPoint : application d'un footer de classification sur les slides existantes ;
- Excel : application d'une shape de classification sur les feuilles existantes, avec un footer pour impression/PDF ;
- Outlook : application idempotente d'un bandeau HTML en haut d'un email ou d'une réunion dont l'utilisateur est l'organisateur, en mode composition ;
- un meme panneau lateral `ClassifyMe` pour choisir `PUBLIC`, `RESTREINT`, `CONFIDENTIEL` ou `SECRET` ;
- une application directe de la classification lors du clic sur une carte ;
- un affichage du niveau selectionne dans le panneau avec les couleurs du niveau choisi ;
- une option Outlook, desactivee par defaut, pour ajouter un prefixe de classification a l'objet du courriel ou de la réunion.

## Cadre technique

Le projet est base sur le template Yeoman officiel pour les Office Add-ins :

- Office Add-in Task Pane ;
- Office.js ;
- TypeScript ;
- cibles Word, PowerPoint, Excel et Outlook compose mode ;
- execution locale via Node.js et npm ;
- aucun backend ;
- aucune base de donnees ;
- aucune API Graph ;
- aucune authentification.

## Organisation du code

Le code est separe pour garder un produit unique `ClassifyMe` tout en distinguant les usages Office et Outlook :

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
  ui/
    taskpane/
      taskpane.ts
      taskpane.html
      taskpane.css
```

- `core/` contient les niveaux, textes, couleurs, noms de shapes et noms de metadonnees communs.
- `hosts/office/` contient les implementations Word, Excel et PowerPoint, plus le routeur qui choisit la bonne implementation selon l'hote actif.
- `hosts/outlook/` contient l'implementation Outlook en mode composition.
- `ui/taskpane/` contient l'interface du panneau lateral et ne porte pas la logique specifique Word, Excel, PowerPoint ou Outlook.

## Manifestes

Le projet utilise deux manifestes pour eviter de fragiliser les cibles deja fonctionnelles :

- `manifest.xml` pour Word, Excel et PowerPoint ;
- `manifest.outlook.xml` pour Outlook en mode composition (emails et réunions organisées).

## Installation

Installer les dependances depuis le dossier du projet :

```powershell
npm install
```

## Commandes utiles

Compiler en mode developpement :

```powershell
npm run build:dev
```

Valider le manifeste Office :

```powershell
npm run validate
```

Valider le manifeste Outlook :

```powershell
npm run validate:outlook
```

Valider les manifestes production :

```powershell
npm run validate:office:production
npm run validate:outlook:production
```

Lancer l'add-in localement :

```powershell
npm start
```

Lancer explicitement dans Word :

```powershell
npm run start:word
```

Lancer explicitement dans PowerPoint :

```powershell
npm run start:powerpoint
```

Lancer explicitement dans Excel :

```powershell
npm run start:excel
```

Lancer explicitement dans Outlook :

```powershell
npm run start:outlook
```

Arreter le serveur local et le debug Office :

```powershell
npm stop
```

Arreter le debug Outlook :

```powershell
npm run stop:outlook
```

Selon le poste, Office ou le navigateur peut demander d'approuver un certificat de developpement local HTTPS.

Pour tester en mode web (webmail et Nouvel Outlook) si le plugin n'est pas sideloadé par la commande `npm` (par défaut, le plugin EST sidelaodé):

```text
1. npm start
2. vérifier https://localhost:3000
3. ouvrir https://aka.ms/olksideload
4. My add-ins
5. Custom Add-ins
6. Add from File
7. choisir manifest.outlook.dev.xml
8. ouvrir un nouveau mail dans OWA
9. Apps > ClassifyMe DEV
```

## Publication GitHub Pages et Microsoft 365

Cette section prepare une publication statique sur GitHub Pages. Elle ne publie rien automatiquement et ne remplace pas une validation fonctionnelle dans Office.

### Prerequis

- Un depot GitHub qui contient ce projet.
- GitHub Pages active sur le depot.
- Une URL GitHub Pages au format suivant :

```text
https://<org>.github.io/<repo>/
```

- Un compte administrateur Microsoft 365 autorise a charger des add-ins integres dans Microsoft 365 Admin Center.
- Un groupe pilote Microsoft 365 dedie au deploiement initial.

### URL de production a modifier

Les manifestes production utilisent volontairement une URL generique :

```text
https://your-org.github.io/your-repo/
```

Avant publication, remplacer cette valeur par l'URL GitHub Pages reelle, ou definir la variable d'environnement `CLASSIFYME_PRODUCTION_BASE_URL` avant le build GitHub Pages :

```powershell
$env:CLASSIFYME_PRODUCTION_BASE_URL="https://<org>.github.io/<repo>/"
npm run build:github-pages
```

La commande accepte aussi une valeur Webpack explicite :

```powershell
npm run build:github-pages -- --env productionUrl=https://<org>.github.io/<repo>/
```

### Commandes de build

La commande de build production standard du projet est :

```powershell
npm run build
```

Elle genere le dossier statique `dist`.

Pour GitHub Pages, utiliser :

```powershell
npm run build:github-pages
```

Cette commande genere le dossier statique `docs`, compatible avec l'option GitHub Pages "Deploy from a branch" puis dossier `/docs`.

### Structure GitHub Pages attendue

Apres `npm run build:github-pages`, le dossier `docs` doit contenir notamment :

```text
docs/
  assets/
    icon-16.png
    icon-32.png
    icon-80.png
  commands.html
  taskpane.html
  manifest.office.production.xml
  manifest.outlook.production.xml
```

Les fichiers JavaScript et CSS generes par Webpack sont egalement presents dans `docs`.

### Manifestes production

Les manifestes de production a utiliser pour Microsoft 365 sont :

- `manifest.office.production.xml` pour Word, Excel et PowerPoint ;
- `manifest.outlook.production.xml` pour Outlook en mode composition.

Ces fichiers doivent pointer vers l'URL GitHub Pages de production. Ils ne doivent contenir aucune URL `localhost`.

Valider les manifestes avant publication :

```powershell
npm run validate:office:production
npm run validate:outlook:production
```

### Configuration GitHub Pages

1. Executer le build GitHub Pages avec l'URL de production correcte.
2. Committer le dossier `docs` genere.
3. Dans GitHub, ouvrir les parametres du depot.
4. Aller dans Pages.
5. Choisir la source "Deploy from a branch".
6. Choisir la branche de publication.
7. Choisir le dossier `/docs`.
8. Attendre la publication GitHub Pages.
9. Verifier en HTTPS :
   - `https://<org>.github.io/<repo>/taskpane.html`
   - `https://<org>.github.io/<repo>/commands.html`
   - `https://<org>.github.io/<repo>/assets/icon-16.png`
   - `https://<org>.github.io/<repo>/assets/icon-32.png`
   - `https://<org>.github.io/<repo>/assets/icon-80.png`

### Sideload avec manifest production

Pour une verification avant deploiement centralise :

1. Verifier que GitHub Pages sert bien `taskpane.html` et les icones en HTTPS.
2. Ouvrir `manifest.office.production.xml` et confirmer que les URL pointent vers GitHub Pages.
3. Sideload `manifest.office.production.xml` pour verifier Word, Excel et PowerPoint.
4. Ouvrir `manifest.outlook.production.xml` et confirmer que les URL pointent vers GitHub Pages.
5. Sideload `manifest.outlook.production.xml` pour verifier Outlook en mode composition.

### Deploiement Microsoft 365 Admin Center

1. Ouvrir Microsoft 365 Admin Center.
2. Aller dans Settings > Integrated apps.
3. Choisir Upload custom apps.
4. Charger `manifest.office.production.xml`.
5. Limiter le deploiement au groupe pilote.
6. Repeter l'operation avec `manifest.outlook.production.xml`.
7. Verifier l'apparition de ClassifyMe dans Word, Excel, PowerPoint et Outlook pour un utilisateur pilote.
8. Elargir le deploiement uniquement apres validation metier et support.

### Limites connues de la publication statique

- GitHub Pages sert uniquement des fichiers statiques : aucune logique serveur n'est disponible.
- Les manifestes doivent etre reconstruits ou modifies si l'URL GitHub Pages change.
- Le cache navigateur ou Office peut conserver une ancienne version du task pane pendant quelques minutes.
- Le MVP reste sans chiffrement, DLP, Graph, authentification ou reporting centralise.
- Le deploiement Microsoft 365 doit d'abord rester limite a un groupe pilote.

### Checklist avant publication

```text
- [ ] aucune URL localhost dans les manifests production
- [ ] taskpane accessible en HTTPS
- [ ] icônes accessibles en HTTPS
- [ ] Word testé
- [ ] Excel testé
- [ ] PowerPoint testé
- [ ] Outlook testé
- [ ] pilote M365 créé
- [ ] add-in déployé uniquement au groupe pilote
```

## Verification manuelle dans Word

```text
1. Lancer npm run start:word.
2. Ouvrir Word si le script ne l'ouvre pas automatiquement.
3. Ouvrir le panneau ClassifyMe depuis le ruban.
4. Cliquer sur Confidentiel.
5. Verifier que le panneau affiche Confidentiel (CONFIDENTIEL).
6. Verifier que l'encart du niveau selectionne reprend les couleurs de la classification.
7. Verifier qu'un bandeau de classification apparait dans l'en-tete Word du document.
8. Cliquer sur Secret.
9. Verifier que le bandeau existant est mis a jour sans creer de doublon visible.
10. Cliquer sur Public.
11. Verifier que le bandeau ClassifyMe est mis a jour avec le niveau Public.
```

## Verification manuelle dans PowerPoint

```text
1. Lancer npm run start:powerpoint.
2. Ouvrir PowerPoint si le script ne l'ouvre pas automatiquement.
3. Ouvrir le panneau ClassifyMe depuis le ruban.
4. Creer ou ouvrir une presentation avec plusieurs slides.
5. Cliquer sur Confidentiel.
6. Verifier qu'un footer ClassifyMe apparait en bas de chaque slide existante.
7. Cliquer sur Secret.
8. Verifier que le footer est remplace sur chaque slide sans creer de doublon.
9. Ajouter une nouvelle slide.
10. Verifier que la nouvelle slide n'est pas marquee automatiquement.
```

## Verification manuelle dans Excel

```text
1. Lancer npm run start:excel.
2. Ouvrir Excel si le script ne l'ouvre pas automatiquement.
3. Ouvrir le panneau ClassifyMe depuis le ruban.
4. Creer ou ouvrir un classeur avec plusieurs feuilles.
5. Cliquer sur Confidentiel.
6. Verifier qu'une shape ClassifyMeBanner apparait en haut de chaque feuille existante.
7. Ouvrir l'aperçu avant impression ou exporter en PDF.
8. Verifier que le footer contient le texte de classification.
9. Cliquer sur Secret.
10. Verifier que le bandeau est remplace sur chaque feuille sans creer de doublon.
11. Ajouter une nouvelle feuille.
12. Verifier que la nouvelle feuille n'est pas marquee automatiquement.
```

## Verification manuelle dans Outlook

```text
1. Lancer npm run start:outlook.
2. Ouvrir Outlook si le script ne l'ouvre pas automatiquement.
3. Creer un nouvel email ou une nouvelle réunion dont vous êtes l'organisateur.
4. Ouvrir le panneau ClassifyMe depuis le ruban de l'élément en composition.
5. Verifier que l'option Ajouter un prefixe a l'objet est decochee.
6. Cliquer successivement sur Public, Restreint, Confidentiel puis Secret.
7. Apres chaque clic, verifier qu'il reste exactement un bandeau ClassifyMe, avec le dernier niveau choisi.
8. Verifier que le contenu, la signature, les citations eventuelles ou le contenu existant de la réunion restent en place.
9. Verifier que l'objet n'a pas ete modifie.
10. Cocher Ajouter un prefixe a l'objet puis cliquer sur Confidentiel.
11. Verifier que l'objet contient [CONFIDENTIEL] sans doublon.
12. Cliquer sur Secret et verifier que le prefixe devient [SECRET] sans doublon.
13. Decocher Ajouter un prefixe a l'objet puis cliquer sur Secret.
14. Verifier que le prefixe [SECRET] est retire de l'objet.
```

### Matrice de validation Outlook

Executer cette matrice dans Outlook Classic Windows, le nouvel Outlook pour Windows et Outlook sur le web. Le resultat attendu de chaque changement est : **exactement un bandeau ClassifyMe**.

| Cas                         | Outlook Classic | Nouvel Outlook | Outlook sur le web |
| --------------------------- | --------------- | -------------- | ------------------ |
| Nouvelle classification     | a verifier      | a verifier     | a verifier         |
| PUBLIC vers RESTREINT       | a verifier      | a verifier     | a verifier         |
| RESTREINT vers CONFIDENTIEL | a verifier      | a verifier     | a verifier         |
| CONFIDENTIEL vers SECRET    | a verifier      | a verifier     | a verifier         |
| Changement repete 4 fois    | a verifier      | a verifier     | a verifier         |
| Message avec signature      | a verifier      | a verifier     | a verifier         |
| Reponse a un email existant | a verifier      | a verifier     | a verifier         |

### Matrice de validation des réunions Outlook

Exécuter cette matrice dans Outlook Classic Windows, le nouvel Outlook pour Windows et Outlook sur le web, sur une réunion créée ou modifiée par son organisateur. Le résultat attendu après chaque action est : **exactement un bandeau ClassifyMe**, le contenu existant de la réunion conservé et, si l'option est active, un seul préfixe d'objet.

| Cas                            | Outlook Classic | Nouvel Outlook | Outlook sur le web |
| ------------------------------ | --------------- | -------------- | ------------------ |
| Nouvelle réunion               | a vérifier      | a vérifier     | a vérifier         |
| PUBLIC                         | a vérifier      | a vérifier     | a vérifier         |
| PUBLIC → CONFIDENTIEL          | a vérifier      | a vérifier     | a vérifier         |
| CONFIDENTIEL → SECRET          | a vérifier      | a vérifier     | a vérifier         |
| Préfixe objet activé           | a vérifier      | a vérifier     | a vérifier         |
| Préfixe objet désactivé        | a vérifier      | a vérifier     | a vérifier         |
| Modification répétée du niveau | a vérifier      | a vérifier     | a vérifier         |
| Réunion avec contenu existant  | a vérifier      | a vérifier     | a vérifier         |

### Idempotence du bandeau Outlook

Le premier pilote identifiait le bandeau uniquement avec les commentaires HTML `ClassifyMe:BannerStart` et `ClassifyMe:BannerEnd`. Outlook sur le web et le nouvel Outlook pour Windows peuvent reecrire le HTML du corps et ne conservent pas necessairement ces commentaires. Le code ne retrouvait alors plus le bandeau et en ajoutait un autre.

Le bandeau courant porte desormais l'identifiant HTML `classifyme-classification-banner`. Outlook sur le web peut le reecrire en `x_classifyme-classification-banner` (et ajouter plusieurs prefixes `x_` dans du HTML cite) afin d'isoler le DOM du message ou de la réunion. A chaque application, l'add-in normalise ces prefixes, supprime tous les bandeaux ClassifyMe trouves, nettoie les anciens bandeaux pilote encore reconnaissables, ajoute le nouveau bandeau une seule fois en tete puis remplace le corps HTML. Cette methode n'utilise pas `prependAsync()`.

L'option Outlook de prefixe d'objet est volontairement desactivee par defaut. Quand elle est cochee, ClassifyMe ajoute ou remplace uniquement les prefixes connus `[PUBLIC]`, `[RESTREINT]`, `[RESTRAINT]`, `[CONFIDENTIEL]` et `[SECRET]` pour l'email ou la réunion. Pour une réunion, une option décochée ne modifie jamais l'objet. Pour préserver le comportement email existant, une option décochée y retire un préfixe ClassifyMe connu sans modifier le reste de l'objet.

## Ce qui fonctionne

- Le panneau lateral `ClassifyMe` est disponible dans Word, PowerPoint et Excel via `manifest.xml`.
- Le panneau lateral `ClassifyMe` est disponible dans Outlook compose mode via `manifest.outlook.xml`, pour les emails et les réunions dont l'utilisateur est l'organisateur.
- Le code commun est isole dans `src/core`.
- Les implementations Word, PowerPoint et Excel sont isolees dans `src/hosts/office`.
- L'implementation Outlook est isolee dans `src/hosts/outlook`.
- Les quatre niveaux de classification sont affiches avec un libelle, un code et une courte description.
- Le niveau choisi est memorise dans l'etat local du panneau.
- L'encart du niveau selectionne reprend le fond et la couleur de texte du niveau choisi.
- Un clic sur une carte applique directement le niveau selectionne.
- Dans Word, un bandeau est insere ou mis a jour dans l'en-tete du document pour les quatre niveaux.
- Dans Word, les proprietes personnalisees `ClassificationLevel`, `ClassificationLabel`, `ClassificationUpdatedAt` et `ClassificationTool` sont mises a jour si l'API Word les accepte dans l'environnement Office utilise.
- Dans PowerPoint, un footer nomme `ClassifyMeFooter` est ajoute en bas de chaque slide existante.
- Dans PowerPoint, les anciens footers `ClassifyMeFooter` sont supprimes avant reapplication pour eviter les doublons.
- Dans Excel, une shape texte nommee `ClassifyMeBanner` est ajoutee en haut de chaque feuille existante.
- Dans Excel, les anciens bandeaux `ClassifyMeBanner` sont supprimes avant reapplication pour eviter les doublons.
- Dans Excel, un footer de classification est applique a chaque feuille existante pour les impressions et exports PDF.
- Dans Excel, les proprietes personnalisees `ClassificationLevel`, `ClassificationLabel`, `ClassificationUpdatedAt` et `ClassificationTool` sont mises a jour si l'API Excel les accepte dans l'environnement Office utilise.
- Dans Outlook compose mode, un bandeau HTML identifie par l'element `div#classifyme-classification-banner` est insere en haut du corps de l'email ou de la réunion ; les variantes d'identifiant prefixees par `x_` d'Outlook Web sont aussi reconnues.
- Dans Outlook compose mode, tous les bandeaux ClassifyMe identifies sont supprimes avant l'insertion d'un seul nouveau bandeau ; les marqueurs du premier pilote sont nettoyes lorsqu'ils sont encore presents.
- Dans Outlook compose mode, l'option de prefixe objet ajoute ou remplace le prefixe de classification si elle est cochee. Si elle est décochée, elle ne modifie pas l'objet de la réunion ; le retrait d'un préfixe connu reste conservé pour les emails afin de ne pas changer leur comportement existant.
- Dans Outlook compose mode, les proprietes personnalisees `ClassificationLevel`, `ClassificationLabel`, `ClassificationUpdatedAt` et `ClassificationTool` sont tentées pour les emails et les réunions organisées. Le bandeau visible reste appliqué si leur enregistrement échoue.
- Les manifestes Outlook de développement et de production déclarent une surface de commande `AppointmentOrganizerCommandSurface`, en plus de la surface email existante. La version du manifeste de production est `1.0.0.2`.

## Limites connues du MVP

- Le footer PowerPoint est applique uniquement aux slides existantes au moment du clic.
- Le footer PowerPoint utilise un positionnement fixe optimise pour les slides widescreen par defaut du MVP.
- Les nouvelles slides creees ensuite ne sont pas automatiquement marquees.
- Le MVP ne modifie pas encore les masques PowerPoint.
- Le MVP ne stocke pas encore de metadonnees personnalisees dans PowerPoint.
- Le bandeau Excel est une shape visible en vue normale, pas une ligne de cellules.
- Le bandeau Excel utilise une largeur fixe prevue pour tenir sur une page portrait standard.
- Le footer Excel est destine aux impressions et exports PDF.
- La classification Excel est appliquee uniquement aux feuilles existantes au moment du clic.
- Les nouvelles feuilles creees ensuite ne sont pas automatiquement marquees.
- L'utilisateur doit recliquer sur un niveau de classification pour mettre a jour le classeur apres creation de nouvelles feuilles.
- Les metadonnees Excel peuvent ne pas etre enregistrees si l'environnement Office ne supporte pas les proprietes personnalisees du classeur.
- Le panneau ne relit pas encore automatiquement une classification deja presente lors de l'ouverture d'un fichier.
- Outlook est limité au mode composition des emails et des réunions dont l'utilisateur est l'organisateur. Le mode lecture, les invitations reçues et les réponses à des invitations ne sont pas implémentés.
- Outlook utilise un manifeste separe, `manifest.outlook.xml`.
- Les proprietes personnalisees Outlook peuvent ne pas etre enregistrees de manière homogène selon le client, l'état réseau ou le type de compte, y compris sur une réunion organisée. Dans ce cas, le bandeau reste appliqué.
- Les proprietes personnalisees Outlook enregistrees en mode composition ne sont pas transmises aux destinataires.
- Le bandeau Outlook est appliqué au corps HTML courant du brouillon ou de la réunion. Les emails ou réunions au format texte brut peuvent refuser l'insertion HTML selon le client Outlook.
- Outlook peut reecrire le HTML d'un brouillon entre la lecture et l'ecriture. Le bandeau ne depend donc pas de commentaires HTML ; la compatibilite pilote ne peut nettoyer un ancien bandeau que si ses commentaires ou sa structure complete restent reconnaissables.
- Le prefixe d'objet Outlook n'est jamais conserve par ClassifyMe si l'option est decochee et si le prefixe existant fait partie des prefixes connus.
- Le MVP ne classifie pas automatiquement les reponses, les fils de conversation Outlook ou les réponses à des invitations.
- Le MVP ne lit pas et ne classe pas les emails reçus ou les invitations reçues.
- Le fichier n'est pas chiffre.
- L'add-in ne bloque pas l'enregistrement, l'envoi, le partage, la copie, l'impression ou le transfert.
- Aucun controle DLP n'est applique.
- Le contenu du fichier n'est pas analyse.
- Aucune IA n'est utilisee.
- Aucun reporting centralise n'est disponible.
- Aucune synchronisation avec Microsoft Purview n'est implementee.

## Prochaines etapes recommandees

1. Verifier manuellement le comportement dans Word pour les quatre niveaux.
2. Verifier manuellement le comportement dans PowerPoint sur une presentation de plusieurs slides.
3. Verifier manuellement le comportement dans Excel sur un classeur de plusieurs feuilles.
4. Verifier manuellement l'idempotence Outlook sur les trois clients cibles, pour les emails (y compris une signature et une réponse existante) et les réunions organisées.
5. Confirmer si PowerPoint doit utiliser les masques de slides dans une prochaine version.
6. Confirmer si Excel doit utiliser une zone reservee ou une position adaptee aux modeles internes.
7. Ajouter la lecture des proprietes ou marquages existants a l'ouverture du panneau.

## Hors perimetre

Les fonctionnalites suivantes restent explicitement exclues du MVP :

- chiffrement ;
- DLP ;
- restriction d'impression ;
- restriction de transfert ;
- analyse automatique du contenu ;
- suggestion par IA ;
- reporting centralise ;
- workflow de validation ;
- base de donnees ;
- API Graph ;
- authentification Entra ID ;
- integration Microsoft Purview.
