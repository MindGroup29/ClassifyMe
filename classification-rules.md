# Règles de classification — ClassifyMe

## Objectif

Cet add-in permet aux utilisateurs de marquer manuellement les emails et documents Office selon le niveau de confidentialité défini par l'entreprise.

Le MVP ne vise pas à chiffrer, bloquer, auditer ou remplacer Microsoft Purview. Il vise uniquement à :
- responsabiliser l'utilisateur ;
- rendre le niveau de confidentialité visible ;
- stocker une métadonnée simple dans le document ou l'email ;
- préparer une éventuelle migration future vers des étiquettes de sensibilité Microsoft.

## Niveaux de classification

| Code technique | Libellé utilisateur | Usage attendu | Marquage visuel |
|---|---|---|---|
| PUBLIC | Public | Information destinée à être diffusée publiquement. | ~~Aucun marquage ou~~ bandeau discret. |
| RESTREINT | Restreint | Information réservée aux collaborateurs de l'entreprise. | Bandeau discret. |
| CONFIDENTIEL | Confidentiel | Information dont la diffusion non maîtrisée peut nuire à l'entreprise. | Bandeau visible. |
| SECRET | Secret | Information critique : RH, finance, juridique, stratégie, données clients sensibles, plans, savoir-faire. | Bandeau très visible. |

## Comportement général

L'utilisateur choisit volontairement le niveau de classification.

L'add-in doit :
- afficher les quatre niveaux ;
- permettre de sélectionner un niveau ;
- appliquer un marquage visuel ;
- stocker le niveau choisi dans une métadonnée technique ;
- permettre de modifier le niveau ultérieurement ;
- éviter toute action irréversible.

## Word / Excel / PowerPoint

Pour les documents Office, l'add-in doit :

1. insérer ou mettre à jour un bandeau de classification ;
2. stocker une propriété personnalisée nommée `ClassificationLevel` ;
3. stocker une propriété personnalisée nommée `ClassificationLabel`;
4. ne pas bloquer l'enregistrement du document dans le MVP ;
5. ne pas modifier le contenu métier du document en dehors du bandeau.

### Propriétés à stocker

| Propriété | Exemple |
|---|---|
| ClassificationLevel | CONFIDENTIEL |
| ClassificationLabel | Confidentiel |
| ClassificationUpdatedAt | 2026-06-16T10:30:00Z |
| ClassificationTool | ClassifyMe |

## Outlook

Pour les emails, l'add-in doit :

1. permettre à l'utilisateur de choisir un niveau avant l'envoi ;
2. insérer un bandeau dans le corps du message ;
3. stocker le niveau dans une propriété personnalisée Outlook si possible ;
4. ne pas ajouter automatiquement de préfixe dans l'objet par défaut ;
5. proposer éventuellement une option manuelle pour ajouter `[CONFIDENTIEL]` ou `[SECRET]`.

## Bandeaux proposés

### PUBLIC

Aucun bandeau par défaut.

### RESTREINT

> Classification : RESTREINT  
> Ce document est destiné à un usage restreint à l'entreprise.

### CONFIDENTIEL

> Classification : CONFIDENTIEL  
> Ce document contient des informations confidentielles. Sa diffusion doit être limitée aux personnes autorisées.

### SECRET

> Classification : SECRET  
> Ce document contient des informations secrètes. Sa diffusion, sa copie et son transfert doivent être strictement maîtrisés.

## Règles UX

L'add-in doit rester simple.

Principes :
- un panneau latéral unique ;
- quatre boutons ou cartes de classification ;
- un bouton "Appliquer" ;
- un message de confirmation clair ;
- aucun jargon Microsoft Purview ;
- aucun paramétrage avancé dans le MVP.

## Hors périmètre MVP

Les fonctionnalités suivantes sont explicitement exclues :

- chiffrement ;
- restriction d'impression ;
- restriction de transfert ;
- DLP ;
- analyse automatique du contenu ;
- suggestion par IA ;
- reporting centralisé ;
- workflow de validation ;
- intégration SIEM ;
- gestion par rôle ou service ;
- synchronisation avec Microsoft Purview.

## Compatibilité future

Les codes techniques doivent rester stables :

- PUBLIC
- RESTREINT
- CONFIDENTIEL
- SECRET

Ces codes pourront servir ultérieurement à migrer vers des étiquettes Microsoft Purview ou une autre solution de classification.

## Décisions ouvertes

À compléter par l'entreprise :

- [ ] Le marquage est-il recommandé ou obligatoire ?
- [x] Le niveau `PUBLIC` doit-il afficher un bandeau ? OUI
- [x] Le préfixe objet email est-il autorisé ? OUI
- [ ] Faut-il afficher une alerte si aucun niveau n'est choisi ?
- [ ] Le bandeau doit-il être inséré en haut ou en bas des documents ?
- [ ] Les documents existants doivent-ils être marqués lors de leur ouverture ?