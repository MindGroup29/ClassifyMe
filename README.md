# ClassifyMe

ClassifyMe est un MVP d'Office Add-in pour Microsoft Word et Microsoft PowerPoint. Il permet a un utilisateur de choisir manuellement un niveau de classification, d'ajouter un marquage visible et de garder l'interface volontairement simple.

Le MVP ne chiffre pas les fichiers, ne bloque pas l'enregistrement, n'analyse pas le contenu et ne remplace pas Microsoft Purview.

## Objectif actuel

Le projet couvre actuellement :

- Word : application d'un bandeau de classification dans l'en-tete du document ;
- PowerPoint : application d'un footer de classification sur les slides existantes ;
- un meme panneau lateral `ClassifyMe` pour choisir `PUBLIC`, `RESTREINT`, `CONFIDENTIEL` ou `SECRET` ;
- une application directe de la classification lors du clic sur une carte ;
- un affichage du niveau selectionne dans le panneau avec les couleurs du niveau choisi.

## Cadre technique

Le projet est base sur le template Yeoman officiel pour les Office Add-ins :

- Office Add-in Task Pane ;
- Office.js ;
- TypeScript ;
- cibles Word et PowerPoint ;
- execution locale via Node.js et npm ;
- aucun backend ;
- aucune base de donnees ;
- aucune API Graph ;
- aucune authentification.

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

Arreter le serveur local et le debug Office :

```powershell
npm stop
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
11. Verifier que le bandeau ClassifyMe est retire.
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

## Ce qui fonctionne

- Le panneau lateral `ClassifyMe` est disponible dans Word et PowerPoint via le manifeste Office.
- Les quatre niveaux de classification sont affiches avec un libelle, un code et une courte description.
- Le niveau choisi est memorise dans l'etat local du panneau.
- L'encart du niveau selectionne reprend le fond et la couleur de texte du niveau choisi.
- Un clic sur une carte applique directement le niveau selectionne.
- Dans Word, un bandeau est insere ou mis a jour dans l'en-tete du document pour les niveaux non publics.
- Dans Word, le niveau `PUBLIC` retire le bandeau ClassifyMe.
- Dans Word, les proprietes personnalisees `ClassificationLevel`, `ClassificationLabel`, `ClassificationUpdatedAt` et `ClassificationTool` sont mises a jour si l'API Word les accepte dans l'environnement Office utilise.
- Dans PowerPoint, un footer nomme `ClassifyMeFooter` est ajoute en bas de chaque slide existante.
- Dans PowerPoint, les anciens footers `ClassifyMeFooter` sont supprimes avant reapplication pour eviter les doublons.

## Limites connues du MVP

- Le footer PowerPoint est applique uniquement aux slides existantes au moment du clic.
- Le footer PowerPoint utilise un positionnement fixe optimise pour les slides widescreen par defaut du MVP.
- Les nouvelles slides creees ensuite ne sont pas automatiquement marquees.
- Le MVP ne modifie pas encore les masques PowerPoint.
- Le MVP ne stocke pas encore de metadonnees personnalisees dans PowerPoint.
- Le panneau ne relit pas encore automatiquement une classification deja presente lors de l'ouverture d'un fichier.
- Excel et Outlook ne sont pas implementes.
- Le fichier n'est pas chiffre.
- L'add-in ne bloque pas l'enregistrement, le partage, la copie, l'impression ou le transfert.
- Aucun controle DLP n'est applique.
- Le contenu du fichier n'est pas analyse.
- Aucune IA n'est utilisee.
- Aucun reporting centralise n'est disponible.
- Aucune synchronisation avec Microsoft Purview n'est implementee.

## Prochaines etapes recommandees

1. Verifier manuellement le comportement dans Word pour les quatre niveaux.
2. Verifier manuellement le comportement dans PowerPoint sur une presentation de plusieurs slides.
3. Confirmer avec l'entreprise si le niveau `PUBLIC` doit rester sans bandeau dans Word.
4. Confirmer si PowerPoint doit utiliser les masques de slides dans une prochaine version.
5. Ajouter la lecture des proprietes ou marquages existants a l'ouverture du panneau.

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
