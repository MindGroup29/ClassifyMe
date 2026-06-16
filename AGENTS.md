# AGENTS.md — Instructions pour Codex

## Objectif du projet

Ce projet développe un add-in Office nommé provisoirement `ConfidentialiteOfficeAddin`.

L'objectif est de permettre aux utilisateurs de classifier manuellement des documents Office et des emails selon quatre niveaux :

- PUBLIC
- RESTREINT
- CONFIDENTIEL
- SECRET

Le projet est un MVP. Il ne doit pas chercher à reproduire Microsoft Purview.

## Cadre technologique obligatoire

Utiliser :

- Office Add-in ;
- Office.js ;
- TypeScript ;
- projet généré avec Yeoman `yo office` ;
- exécution locale via Node.js et npm ;
- interface simple HTML/CSS ou React léger si le template Yeoman le propose ;
- aucun backend dans le MVP ;
- aucune base de données dans le MVP.

Ne pas utiliser :

- .NET ;
- ASP.NET ;
- base SQL ;
- API Graph ;
- authentification Entra ID ;
- IA ;
- moteur DLP ;
- chiffrement ;
- intégration Purview.

## Périmètre fonctionnel MVP

L'add-in doit permettre de :

1. choisir un niveau de classification ;
2. appliquer un marquage visuel ;
3. stocker une métadonnée technique ;
4. modifier le niveau choisi ;
5. garder une interface simple et compréhensible.

## Applications cibles

Priorité de développement :

1. Word ;
2. Excel ;
3. PowerPoint ;
4. Outlook.

Ne pas commencer par Outlook, car les API et comportements sont plus spécifiques.

## Règles de développement

Le code doit être simple, lisible et maintenable.

Obligations :

- écrire des commentaires explicatifs clairs dans le code ;
- commenter les choix non évidents ;
- isoler les constantes de classification dans un fichier dédié ;
- éviter les duplications inutiles ;
- nommer clairement les fonctions ;
- limiter les dépendances ;
- ne pas ajouter de fonctionnalité non demandée ;
- ne pas ajouter de tests dans cette phase MVP.

## Documentation obligatoire

À chaque modification fonctionnelle ou technique, mettre à jour `README.md`.

Le `README.md` doit toujours expliquer :

- l'objectif actuel du projet ;
- l'état réel de ce qui fonctionne ;
- les commandes utiles ;
- les limites connues ;
- les prochaines étapes recommandées.

Ne jamais laisser le README promettre une fonctionnalité qui n'existe pas encore.

## Règles de prudence

Avant toute modification significative :

1. lire `classification-rules.md` ;
2. vérifier que la demande reste dans le périmètre MVP ;
3. refuser d'étendre implicitement le projet vers du chiffrement, DLP, reporting centralisé ou Purview ;
4. préférer une petite modification claire à une refonte large.

## Style attendu

Le code doit être compréhensible par un DSI ou un développeur reprenant le projet.

Favoriser :

- fonctions courtes ;
- noms explicites ;
- commentaires utiles ;
- architecture simple ;
- fichiers peu nombreux.

Éviter :

- surarchitecture ;
- abstraction prématurée ;
- frameworks inutiles ;
- logique cachée ;
- dépendances lourdes.

## Validation manuelle

Aucun test automatisé n'est demandé dans cette phase.

À chaque évolution, indiquer dans `README.md` comment vérifier manuellement le comportement dans Office.

Exemple :

```text
1. Lancer l'add-in.
2. Ouvrir Word.
3. Cliquer sur "Confidentiel".
4. Vérifier que le bandeau apparaît.
5. Vérifier que le niveau est affiché dans le panneau.
````