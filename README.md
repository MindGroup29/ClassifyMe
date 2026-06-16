# ClassifyMe

ClassifyMe est un MVP d'Office Add-in pour Microsoft Word. Il permet a un utilisateur de choisir manuellement un niveau de classification documentaire, d'ajouter un bandeau visible dans le document et de stocker des proprietes personnalisees simples.

Le MVP ne chiffre pas les documents, ne bloque pas l'enregistrement, n'analyse pas le contenu et ne remplace pas Microsoft Purview.

## Objectif actuel

Le premier tour du projet couvre uniquement Word avec un Task Pane Add-in Office.js en TypeScript.

L'add-in permet de :

- afficher un panneau lateral `ClassifyMe` dans Word ;
- choisir un niveau parmi `PUBLIC`, `RESTREINT`, `CONFIDENTIEL` et `SECRET` ;
- afficher le niveau selectionne dans le panneau ;
- appliquer directement la classification au document lors du clic sur un niveau ;
- creer ou mettre a jour un bandeau unique en haut du document pour `RESTREINT`, `CONFIDENTIEL` et `SECRET` avec un tableau Word pleine largeur a une cellule ;
- retirer le bandeau ClassifyMe lorsque le niveau `PUBLIC` est applique, car `classification-rules.md` indique qu'aucun bandeau n'est affiche par defaut pour ce niveau ;
- stocker les proprietes personnalisees `ClassificationLevel`, `ClassificationLabel`, `ClassificationUpdatedAt` et `ClassificationTool`.

## Cadre technique

Le projet est base sur le template Yeoman officiel pour les Office Add-ins :

- Office Add-in Task Pane ;
- Office.js ;
- TypeScript ;
- cible Word ;
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

Lancer l'add-in localement dans Word :

```powershell
npm start
```

Arreter le serveur local et le debug Office :

```powershell
npm stop
```

Selon le poste, Office ou le navigateur peut demander d'approuver un certificat de developpement local HTTPS.

## Verification manuelle dans Word

```text
1. Lancer npm start.
2. Ouvrir Word si le script ne l'ouvre pas automatiquement.
3. Ouvrir le panneau ClassifyMe depuis le ruban.
4. Cliquer sur Confidentiel.
5. Verifier que le panneau affiche Confidentiel (CONFIDENTIEL).
6. Verifier qu'un bandeau de classification apparait en haut du document.
7. Verifier que le fond du bandeau reste visible sans passer la souris dessus.
8. Cliquer sur Secret.
9. Verifier que le bandeau existant est mis a jour sans creer de doublon visible.
10. Cliquer sur Public.
11. Verifier que le bandeau ClassifyMe est retire.
```

## Ce qui fonctionne

- Le panneau lateral `ClassifyMe` est disponible dans Word.
- Les quatre niveaux de classification sont affiches avec un libelle, un code et une courte description.
- Le niveau choisi est memorise dans l'etat local du panneau.
- Un clic sur une carte de classification applique directement le niveau selectionne.
- Un bandeau unique est insere ou mis a jour en haut du document pour les niveaux non publics.
- Le fond du bandeau est applique a une cellule de tableau Word pour rester visible sans survol de la souris.
- Les anciens bandeaux ClassifyMe en doublon sont supprimes lors d'une reapplication.
- Les proprietes personnalisees du document sont mises a jour si l'API Word les accepte dans l'environnement Office utilise.

## Limites connues du MVP

- Le MVP cible Word uniquement.
- Excel, PowerPoint et Outlook ne sont pas implementes.
- Le document n'est pas chiffre.
- L'add-in ne bloque pas l'enregistrement, le partage, la copie, l'impression ou le transfert.
- Aucun controle DLP n'est applique.
- Le contenu du document n'est pas analyse.
- Aucune IA n'est utilisee.
- Aucun reporting centralise n'est disponible.
- Aucune synchronisation avec Microsoft Purview n'est implementee.
- Le panneau ne relit pas encore automatiquement une classification deja presente lors de l'ouverture d'un document.

## Prochaines etapes recommandees

1. Verifier manuellement le comportement dans Word pour les quatre niveaux.
2. Confirmer avec l'entreprise si le niveau `PUBLIC` doit rester sans bandeau.
3. Ajouter la lecture des proprietes existantes a l'ouverture du panneau.
4. Ajuster le style exact du bandeau selon la charte interne.
5. Etendre ensuite seulement vers Excel, puis PowerPoint, en conservant les memes constantes de classification.

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
