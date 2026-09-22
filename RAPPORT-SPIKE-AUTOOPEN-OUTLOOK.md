# Rapport de développement — Spike d’ouverture de ClassifyMe dans Outlook

**Date du débrief :** 22 septembre 2026  
**Statut :** spike techniquement concluant pour le repli par notification  
**Périmètre testé :** Outlook Classic Windows et Outlook sur le web avec Exchange Online

## 1. Résumé exécutif

Le spike devait déterminer si le panneau latéral ClassifyMe pouvait s’ouvrir
automatiquement lors de la création d’un message ou d’une réunion Outlook.

Le résultat comporte deux conclusions distinctes :

1. **L’ouverture automatique du task pane sans action utilisateur n’est pas officiellement
   supportée dans Outlook.** `Office.addin.showAsTaskpane()` dépend du requirement set
   `SharedRuntime 1.1`, alors que les runtimes partagés ne sont pas pris en charge par
   Outlook.
2. **Le meilleur repli officiellement supporté fonctionne :** l’activation événementielle
   affiche une notification Outlook contenant l’action **Ouvrir ClassifyMe**. Un clic sur
   cette action ouvre le task pane existant.

Le repli a été confirmé pour la création d’un nouveau message dans :

- Outlook Classic Windows ;
- Outlook sur le web.

Le nouvel Outlook Windows, les réponses, les transferts et les nouvelles réunions restent
à tester. Le spike ne peut donc pas encore conclure à une compatibilité complète sur toute
la matrice cible.

## 2. Décision proposée

| Question | Conclusion |
| --- | --- |
| Peut-on ouvrir automatiquement ClassifyMe sans clic ? | **Non**, pas avec une API officiellement supportée dans Outlook. |
| Peut-on réagir automatiquement à une nouvelle composition ? | **Oui**, avec Event-Based Activation et Mailbox 1.10. |
| Peut-on guider l’utilisateur vers ClassifyMe ? | **Oui**, avec une notification et une action `showTaskPane`. |
| Le repli est-il confirmé sur Classic ? | **Oui**, pour un nouveau message. |
| Le repli est-il confirmé sur Outlook Web ? | **Oui**, pour un nouveau message. |
| Le spike est-il prêt à passer tel quel en production ? | **Non**. La matrice restante et un pilote de déploiement doivent être exécutés. |

La recommandation est un **GO conditionnel** si le produit accepte le clic utilisateur sur
la notification. Si l’exigence reste une ouverture entièrement automatique, la conclusion
du spike est **NO-GO**.

## 3. Architecture expérimentée

```text
Création d’un message ou d’une réunion
                 │
                 ▼
Événement Outlook Mailbox 1.10
                 │
                 ▼
Runtime court ClassifyMe
                 │
                 ▼
Notification « Ouvrir ClassifyMe »
                 │
          clic utilisateur
                 │
                 ▼
Task pane ClassifyMe existant
```

Les événements déclarés sont :

- `OnNewMessageCompose` ;
- `OnNewAppointmentOrganizer`.

Les API et mécanismes utilisés sont :

- `Office.actions.associate()` pour associer les handlers ;
- `Office.context.mailbox.item.notificationMessages.addAsync()` ;
- `InsightMessage` avec l’action `showTaskPane` ;
- `event.completed()` sur tous les chemins de fin ;
- `Office.initialize` pour initialiser correctement le runtime HTML dans Outlook Web.

`OnNewMessageCompose` est également destiné à couvrir les réponses, réponses à tous et
transferts, mais pas la réouverture d’un brouillon. Ces variantes n’ont pas encore été
confirmées manuellement dans ce projet.

## 4. Modifications réalisées

### Fichiers créés

- `src/hosts/outlook/outlookLaunchEvents.ts` : handlers, notification de repli,
  notification diagnostique, instrumentation et appel garanti à `event.completed()` ;
- `src/hosts/outlook/outlookLaunchEvents.html` : runtime HTML utilisé par Outlook Web et
  initialisation Office.js ;
- `RAPPORT-SPIKE-AUTOOPEN-OUTLOOK.md` : présent rapport.

### Fichiers modifiés

- `manifest.outlook.xml` : ajout du `VersionOverridesV1_1`, du runtime et des deux
  `LaunchEvent` ;
- `webpack.config.js` : ajout du bundle `autoopen`, génération de la page HTML, chargement
  bloquant, cache-busting et désactivation du client HMR pour le runtime événementiel ;
- `README.md` : faisabilité, protocole, résultats, limites et procédure de retrait.

Le manifeste `manifest.outlook.production.xml`, le task pane et la logique de
classification n’ont pas été modifiés par le spike.

## 5. Historique du diagnostic

### Étape 1 — Implémentation initiale

Le manifeste DEV a reçu un `VersionOverridesV1_1` exigeant `Mailbox 1.10`. Les commandes
manuelles existantes ont été conservées dans cette branche du manifeste afin que le bouton
historique reste disponible.

Comme `showAsTaskpane()` n’est pas compatible avec le modèle de runtime Outlook, le handler
a été adapté pour afficher un `InsightMessage` proposant l’ouverture du panneau.

### Étape 2 — Échec initial dans Outlook Classic

Symptômes :

- aucune notification ;
- aucun panneau ;
- aucun log visible.

Cause identifiée : `webpack-dev-server` injectait son client WebSocket et le Hot Module
Replacement dans `autoopen.js`. Outlook Classic charge ce fichier dans un runtime
JavaScript-only ; le code du serveur de développement pouvait échouer avant
`Office.actions.associate()`.

Correction : désactivation de `client`, `hot` et `liveReload` dans le serveur de
développement. La notification a ensuite été confirmée dans Outlook Classic.

### Étape 3 — Échec initial dans Outlook Web

Les requêtes vers `autoopen.html` et `autoopen.js` étaient présentes, mais aucune
notification ne s’affichait.

Les corrections intermédiaires ont permis de fiabiliser le chargement :

- script généré sans `defer`, pour associer les handlers immédiatement ;
- versionnement du manifeste DEV ;
- paramètre de cache sur les ressources ;
- hash Webpack sur le bundle JavaScript ;
- traces `runtime chargé` et `handlers associés`.

Ces traces ont prouvé que le bundle était bien exécuté. La console a ensuite fourni la
cause finale : Office.js considérait que l’application n’avait appelé ni
`Office.onReady()` ni défini `Office.initialize`.

Correction finale : définition de `Office.initialize` dans `autoopen.html`, avant le
chargement du bundle événementiel. La notification est alors devenue fonctionnelle dans
Outlook Web.

## 6. Résultats au terme du spike

| Scénario | Outlook Classic | New Outlook | Outlook Web |
| --- | --- | --- | --- |
| Nouveau message | Notification fonctionnelle | À tester | Notification fonctionnelle |
| Réponse à un message | À confirmer | À tester | À tester |
| Transfert d’un message | À confirmer | À tester | À tester |
| Nouvelle réunion organisée | À confirmer | À tester | À tester |
| Bouton manuel ClassifyMe | Fonctionnel | À tester | À confirmer |

En l’état, les tests valident le mécanisme technique principal sur deux clients, mais pas
encore l’ensemble du périmètre demandé.

## 7. Instrumentation disponible

Les logs utilisent le préfixe suivant :

```text
[ClassifyMe][Spike AutoOpen]
```

Ils permettent de suivre :

- le chargement du runtime ;
- l’initialisation Office.js sur le Web ;
- l’association des handlers ;
- le déclenchement de l’événement ;
- le démarrage du handler ;
- la raison pour laquelle `showAsTaskpane()` n’est pas appelé ;
- le succès ou l’échec de la notification ;
- l’exécution de `event.completed()`.

Si l’`InsightMessage` échoue, une seconde tentative affiche une notification informative
minimale. Le traitement reste non bloquant et ne modifie ni le contenu ni la
classification de l’élément.

## 8. Validation technique effectuée

Les commandes suivantes réussissent :

```text
npm run build:dev
npm run validate:outlook
npm run lint
```

Le lint conserve deux avertissements préexistants dans la classification Word, sans erreur
liée au spike.

Le comportement réel a été confirmé manuellement pour un nouveau message sur Outlook
Classic et Outlook Web. Une compilation réussie et un manifeste valide ne sont pas, à eux
seuls, considérés comme une preuve fonctionnelle.

## 9. Limites et risques

- La notification demande un clic ; elle ne satisfait pas une exigence d’ouverture sans
  interaction.
- New Outlook Windows n’a pas encore été testé.
- Les réponses, transferts et réunions ne sont pas encore qualifiés.
- Outlook Web et New Outlook ne garantissent pas l’activation sur certaines surfaces de
  composition non standard.
- Le cache du manifeste et des ressources peut masquer une nouvelle version pendant les
  tests ; le versionnement ajouté réduit ce risque.
- La désactivation du HMR concerne le serveur DEV complet et retire le rechargement à chaud
  du task pane pendant ce spike. Elle n’a pas d’impact sur un build de production.
- Le runtime événementiel dépend d’une connexion réseau et doit terminer rapidement en
  appelant `event.completed()`.
- Le comportement doit encore être éprouvé après un déploiement centralisé à un groupe
  pilote, et pas uniquement par sideload.

## 10. Recommandations

1. Terminer la matrice sur New Outlook, puis sur les réponses, transferts et réunions.
2. Confirmer que le bouton manuel reste opérationnel sur Web et New Outlook.
3. Faire valider par le métier l’acceptabilité d’un clic sur la notification.
4. Si le repli est accepté, déployer le manifeste DEV à un petit groupe pilote avec des
   ressources HTTPS stables.
5. Avant industrialisation, réduire les logs au strict nécessaire et isoler si besoin le
   build événementiel afin de réactiver le HMR du task pane en développement.
6. Ne reporter les éléments du spike dans le manifeste de production qu’après cette
   validation.

## 11. Procédure de retrait

Si le spike est abandonné :

1. supprimer `src/hosts/outlook/outlookLaunchEvents.ts` ;
2. supprimer `src/hosts/outlook/outlookLaunchEvents.html` ;
3. retirer l’entrée et le plugin Webpack `autoopen` ;
4. rétablir les options habituelles de `webpack-dev-server` si nécessaire ;
5. retirer le `VersionOverridesV1_1` imbriqué de `manifest.outlook.xml` ;
6. retirer du README la section et les limites propres au spike.

La logique de classification et le manifeste Outlook de production ne nécessitent aucune
restauration.

## 12. Références Microsoft

- [Activation basée sur les événements](https://learn.microsoft.com/office/dev/add-ins/develop/event-based-activation)
- [Événements de nouvelle composition Outlook](https://learn.microsoft.com/office/dev/add-ins/outlook/on-new-compose-events-walkthrough)
- [Runtimes Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/testing/runtimes)
- [Afficher ou masquer un task pane](https://learn.microsoft.com/office/dev/add-ins/develop/show-hide-add-in)
- [API NotificationMessages](https://learn.microsoft.com/javascript/api/outlook/office.notificationmessages)
- [Exemple Outlook Event-Based Activation](https://learn.microsoft.com/samples/officedev/office-add-in-samples/outlook-add-in-set-signature/)

