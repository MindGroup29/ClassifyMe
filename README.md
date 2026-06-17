# ClassifyMe

ClassifyMe est un MVP d'Office Add-in pour Microsoft Word, Microsoft PowerPoint, Microsoft Excel et Outlook en mode composition. Il permet a un utilisateur de choisir manuellement un niveau de classification, d'ajouter un marquage visible et de garder l'interface volontairement simple.

Le MVP ne chiffre pas les fichiers ou emails, ne bloque pas l'enregistrement ou l'envoi, n'analyse pas le contenu et ne remplace pas Microsoft Purview.

## Objectif actuel

Le projet couvre actuellement :

- Word : application d'un bandeau de classification dans l'en-tete du document ;
- PowerPoint : application d'un footer de classification sur les slides existantes ;
- Excel : application d'une shape de classification sur les feuilles existantes, avec un footer pour impression/PDF ;
- Outlook : application d'un bandeau HTML en haut d'un email en cours de redaction ;
- un meme panneau lateral `ClassifyMe` pour choisir `PUBLIC`, `RESTREINT`, `CONFIDENTIEL` ou `SECRET` ;
- une application directe de la classification lors du clic sur une carte ;
- un affichage du niveau selectionne dans le panneau avec les couleurs du niveau choisi ;
- une option Outlook, desactivee par defaut, pour ajouter un prefixe de classification a l'objet du courriel.

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
- `manifest.outlook.xml` pour Outlook en mode composition.

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
3. Creer un nouvel email.
4. Ouvrir le panneau ClassifyMe depuis le ruban du message en composition.
5. Verifier que l'option Ajouter un prefixe a l'objet du courriel est decochee.
6. Cliquer sur Confidentiel.
7. Verifier qu'un bandeau HTML ClassifyMe apparait en haut du corps du mail.
8. Verifier que l'objet du mail n'a pas ete modifie.
9. Cocher Ajouter un prefixe a l'objet du courriel.
10. Cliquer de nouveau sur Confidentiel.
11. Verifier que l'objet contient [CONFIDENTIEL] sans doublon.
12. Cliquer sur Secret.
13. Verifier que le bandeau est remplace sans doublon.
14. Verifier que le prefixe objet devient [SECRET] sans doublon.
```

L'option Outlook de prefixe d'objet est volontairement desactivee par defaut. Quand elle est cochee, ClassifyMe ajoute ou remplace uniquement les prefixes connus `[PUBLIC]`, `[RESTREINT]`, `[RESTRAINT]`, `[CONFIDENTIEL]` et `[SECRET]`. Quand elle est decochee, l'objet existant n'est pas modifie.

## Ce qui fonctionne

- Le panneau lateral `ClassifyMe` est disponible dans Word, PowerPoint et Excel via `manifest.xml`.
- Le panneau lateral `ClassifyMe` est disponible dans Outlook compose mode via `manifest.outlook.xml`.
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
- Dans Outlook compose mode, un bandeau HTML identifie par les marqueurs `ClassifyMe:BannerStart` et `ClassifyMe:BannerEnd` est insere en haut du corps du mail.
- Dans Outlook compose mode, un ancien bandeau ClassifyMe est remplace avant insertion pour eviter les doublons.
- Dans Outlook compose mode, l'option de prefixe objet ajoute ou remplace le prefixe de classification seulement si elle est cochee.
- Dans Outlook compose mode, les proprietes personnalisees `ClassificationLevel`, `ClassificationLabel`, `ClassificationUpdatedAt` et `ClassificationTool` sont enregistrees si l'API Outlook et le compte courant l'acceptent.

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
- Outlook est limite au mode composition de messages. Le mode lecture n'est pas implemente.
- Outlook utilise un manifeste separe, `manifest.outlook.xml`.
- Les proprietes personnalisees Outlook peuvent ne pas etre enregistrees selon le client, l'etat reseau ou le type de compte. Dans ce cas, le bandeau reste applique.
- Les proprietes personnalisees Outlook enregistrees en mode composition ne sont pas transmises aux destinataires.
- Le bandeau Outlook est applique au corps HTML courant du brouillon. Les emails au format texte brut peuvent refuser l'insertion HTML selon le client Outlook.
- Le prefixe d'objet Outlook n'est jamais ajoute si l'option est decochee.
- Le MVP ne classifie pas automatiquement les reponses ou les fils de conversation Outlook.
- Le MVP ne lit pas et ne classe pas les emails recus.
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
4. Verifier manuellement le comportement dans Outlook sur un nouvel email en mode composition.
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
