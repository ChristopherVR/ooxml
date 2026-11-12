# Fonctionnalités prises en charge et garantie de préservation

Cette page est la liste écrite qu’exige [la politique de stabilité](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md)
pour un produit stable. Elle indique ce que l’éditeur Excel modélise et modifie, ce qu’il conserve intact lors d’une modification et d’un
enregistrement, et ce qu’il supprime ou refuse (avec le message affiché). [Fonctionnalités et limites](/fr/features) décrit
le comportement visible par l’utilisateur, et la [liste des fonctions](/functions) le moteur de calcul. Rien ici n’affirme que l’éditeur
correspond à Microsoft Excel ni qu’un fichier enregistré est identique à la source : l’affirmation est plus restreinte, voir ci-dessous.

## La garantie {#the-guarantee}

Le contenu que l’éditeur ne modélise pas (éléments inconnus, listes d’extensions, parties, relations,
XML personnalisé et équivalents) survit à l’ouverture d’un classeur, à sa modification et à son enregistrement, ou l’éditeur
signale qu’il ne peut pas le conserver. Une perte silencieuse est un bogue et est corrigée dans une version de correctif.

« Survit » signifie que le XML ou les octets reviennent inchangés, à une position valide. Cela ne signifie pas que
l’ensemble du paquet est identique : les parties que l’éditeur régénère à partir de son modèle (listées ci-dessous) sont
réécrites, et Excel peut les réécrire à nouveau lors de son prochain enregistrement.

## Ce qui est modélisé et modifié {#what-is-modelled-and-edited}

Ces éléments sont lus dans le modèle, modifiés par la session d’édition et réécrits à partir du modèle.
Les attributs et enfants qu’ils contiennent et que le modèle ne porte pas sont conservés là où le tableau l’indique.

| Domaine            | Modélisé et écrit à partir du modèle                                                                                                                                                                                                                                          |
| ------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Cellules           | Valeurs, formules (les formules partagées sont développées), formules matricielles et matricielles dynamiques avec leurs métadonnées, formules de table de données, chaînes de texte enrichi, erreurs, dates dans les deux systèmes de dates.                                 |
| Lignes et colonnes | Hauteurs, largeurs, masquage, niveaux de plan, réduction, formats de lignes et de colonnes, indications de bordure épaisse en haut et en bas de ligne.                                                                                                                        |
| Formats            | Polices, remplissages, bordures, formats de nombres, alignement, protection, styles de cellule nommés, formats différentiels.                                                                                                                                                 |
| Feuille            | Fusions, volets figés, affichages, zoom, quadrillage, couleur d’onglet, état (visible, masquée, très masquée), protection avec empreintes de mot de passe, mise en page, options d’impression, zone d’impression, en-têtes et pieds de page.                                  |
| Règles et listes   | Mise en forme conditionnelle (les types de règles pris en charge), validation des données, liens hypertextes, filtres automatiques, tableaux avec leurs filtres et leur ligne de total, noms définis.                                                                         |
| Annotations        | Notes, commentaires en fil avec leurs personnes.                                                                                                                                                                                                                              |
| Dessins            | Graphiques (par le modèle de graphique partagé ; un graphique que le modèle ne peut pas contenir est conservé en lecture seule), images, formes que l’éditeur comprend, SmartArt (dessin en cache régénéré, données source conservées), sparklines, segments et chronologies. |
| Analyse            | Tableaux croisés dynamiques sur plage de feuille pris en charge par l’éditeur, tables de données d’analyse de scénarios, scénarios du Gestionnaire de scénarios.                                                                                                              |
| Classeur           | Ordre et noms des feuilles, feuille active, noms définis (avec les attributs des noms de macros et de fonctions), paramètres de calcul, protection du classeur, système de dates, palette et polices du thème.                                                                |
| Propriétés         | Propriétés du document principales, étendues et personnalisées.                                                                                                                                                                                                               |

## Ce qui est conservé sans modification lors d’une modification et d’un enregistrement {#what-is-kept-unchanged-through-an-edit-and-a-save}

Les modifications de commentaires en fil conservent les auteurs, les dates de création, les identités des messages, les enregistrements de compte des personnes,
l’état résolu et le balisage de commentaire inconnu. Les mentions importées suivent le texte inséré ou supprimé
autour de leurs libellés inchangés ; modifier un libellé supprime la mention. La copie d’un fil attribue de
nouvelles identités de messages et de mentions tout en conservant les personnes auxquelles il renvoie. L’éditeur peut
résoudre et rouvrir les fils, mais ne peut pas créer de nouvelles @mentions.

L’écrivain copie une partie du paquet source octet pour octet (avec ses relations et chaque
partie qu’elles atteignent) lorsque rien dans le modèle ne la régénère. La source de ce comportement est
`PackageWriter.carry` dans `src/core/xlsx/write/package-writer.ts`, piloté par `saveXlsx`
(`src/core/xlsx/write/save.ts`) et `writeWorksheet` (`src/core/xlsx/write/worksheet.ts`).

| Conservé                                                                              | Méthode                                                                                                                                                                                                                                                                                                                                            |
| ------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Projet VBA (`xl/vbaProject.bin`) d’un `.xlsm`                                         | Conservé avec sa relation et le type de contenu avec macros. Les macros ne sont jamais exécutées. Enregistrer le classeur au format `.xlsx` supprime le projet, car ce format ne peut pas le contenir.                                                                                                                                             |
| XML personnalisé (`customXml/*`) et leurs parties de propriétés                       | Conservés avec la relation du classeur et les relations entre les parties.                                                                                                                                                                                                                                                                         |
| Liens externes (`xl/externalLinks/*`)                                                 | Conservés avec leurs cibles de relation (externes) ; l’élément `externalReferences` reste dans `workbook.xml`. Les liens ne sont pas mis à jour (signalés au chargement).                                                                                                                                                                          |
| Connexions, tables de requête et autres parties du classeur                           | Toute relation de classeur que l’écrivain ne régénère pas est conservée avec sa partie cible, quel que soit son type.                                                                                                                                                                                                                              |
| Paramètres d’imprimante, propriétés de contrôle, propriétés personnalisées            | Les relations de feuille que l’écrivain ne régénère pas sont conservées avec leurs parties cibles (`printerSettings`, `ctrlProps`, `customProperty`, objets intégrés, images).                                                                                                                                                                     |
| Éléments de feuille que le modèle ne représente pas                                   | `sheetCalcPr`, `phoneticPr`, `rowBreaks`, `colBreaks`, `customProperties`, `cellWatches`, `ignoredErrors`, `smartTags`, `oleObjects`, `controls`, `picture`, `webPublishItems`, `extLst` et les autres du schéma sont conservés en XML et réécrits dans l’ordre du schéma. Les entrées d’extension inconnues dans `extLst` restent telles quelles. |
| Attributs source des éléments de feuille modélisés                                    | `sheetPr`, `sheetViews`, `sheetFormatPr`, `sheetProtection`, `autoFilter`, `pageMargins`, `pageSetup` et `headerFooter` conservent les attributs que le modèle ne porte pas lorsque la partie modélisée n’a pas changé.                                                                                                                            |
| Éléments du classeur que le modèle ne représente pas                                  | `fileVersion`, `fileSharing`, `functionGroups`, `externalReferences`, `oleSize`, `customWorkbookViews`, `smartTagPr`, `smartTagTypes`, `webPublishing`, `fileRecoveryPr`, `webPublishObjects`, `extLst`, chaque fenêtre de `bookViews` et les attributs non modélisés de `workbookPr`, `workbookView` et `calcPr`.                                 |
| Feuilles de graphique, feuilles de boîte de dialogue, feuilles de macro               | Conservées entières (affichées comme feuilles vides et signalées au chargement) ; leurs références suivent les renommages et suppressions de feuilles.                                                                                                                                                                                             |
| Parties de tableaux croisés dynamiques non pris en charge, caches de tableaux croisés | Conservés avec leurs relations ; les définitions de cache suivent les renommages de feuilles.                                                                                                                                                                                                                                                      |
| Dessins et graphiques que l’éditeur n’a pas touchés                                   | Les parties de dessin, de graphique, de style, de couleur et de médias d’un objet non modifié sont conservées octet pour octet.                                                                                                                                                                                                                    |
| Thème                                                                                 | Conservé, sauf si la palette du thème a été modifiée.                                                                                                                                                                                                                                                                                              |

`src/core/xlsx/write/preservation.test.ts` vérifie les premières lignes de ce tableau : il construit un paquet
avec un projet VBA, du XML personnalisé, un lien externe avec une cible externe, une partie de connexions,
des paramètres d’imprimante, une partie de propriétés de contrôle, une relation inconnue, des entrées `extLst` inconnues
dans le classeur et dans la feuille, chaque élément de feuille listé ci-dessus et une seconde fenêtre de classeur, modifie
une cellule, enregistre, et vérifie que chaque partie est identique octet pour octet, que chaque relation et chaque type de contenu
est présent, et que chaque fragment XML est revenu dans l’ordre du schéma. Un second enregistrement du résultat est stable.
`save.test.ts` fait de même avec les fichiers propres à Excel et à openpyxl.

## Ce qui est supprimé, et pourquoi {#what-is-dropped-and-why}

| Supprimé                                                                                                 | Motif                                                                                                                                                                                                                  |
| -------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `xl/calcChain.xml`                                                                                       | Reconstruit par Excel ; le classeur est enregistré avec `fullCalcOnLoad` afin que chaque valeur soit recalculée à l’ouverture.                                                                                         |
| Signatures numériques (l’origine, les parties de signature et leurs relations)                           | Toute modification les invalide. Signalé au chargement : « Le classeur est signé numériquement ... l’enregistrement le supprime ». Excel les supprime aussi.                                                           |
| Le projet VBA lorsqu’il est enregistré en `.xlsx`                                                        | Ce format ne peut pas le contenir. Le format `.xlsm` est conservé tel qu’il a été chargé.                                                                                                                              |
| `x15ac:absPath` (emplacement du dernier enregistrement) et `xr:revisionPtr`                              | Indices de rédaction pour Excel, régénérés par lui. Le chemin absolu divulguerait en outre l’organisation des dossiers de l’auteur.                                                                                    |
| Attributs `xr:uid`, `spans` des lignes, `x14ac:dyDescent`                                                | Identifiants de révision et indices de mise en page que Excel régénère.                                                                                                                                                |
| `ph` des lignes et des cellules (afficher les phonétiques), `cm` des cellules sans partie de métadonnées | Non modélisés. Les guides phonétiques restent dans les chaînes partagées ; seul l’indicateur qui les affiche est perdu.                                                                                                |
| Formules partagées et chaînes en ligne en tant que telles                                                | Écrites comme les formules simples équivalentes et les chaînes partagées (mêmes valeurs).                                                                                                                              |
| Types de règles de mise en forme conditionnelle que l’éditeur ne prend pas en charge                     | Signalés au chargement : `Conditional formatting rule type "..." is not supported and was dropped.` Les autres règles de la feuille sont conservées.                                                                   |
| Une partie de tableau ou de graphique qui ne peut pas être lue                                           | Signalé au chargement (« Une partie de tableau n’a pas pu être lue et a été supprimée », « Le graphique ... n’a pas pu être lu ; il est conservé mais non affiché » : la partie du graphique elle-même est conservée). |
| Un élément de feuille hors du schéma                                                                     | Il ne peut pas être placé à une position valide ; il est donc signalé au chargement : `the element "..." is not part of the worksheet schema and is dropped on save.`                                                  |

## Ce qui est refusé ou signalé {#what-is-refused-or-reported}

Rien n’est refusé en silence. L’éditeur signale les situations par `workbook.warnings` (affiché par les hôtes sous forme d’avis)
et par une exception :

- un paquet chiffré sans mot de passe demande un mot de passe (`load/detect.ts`) ; un mot de passe erroné est une erreur ;
- une partie non prise en charge ou endommagée est signalée par son nom (« ... est conservé mais non affiché », « ... n’a pas pu être lu ») ;
- un classeur contenant des macros VBA ou des liens externes indique qu’ils sont conservés, mais ni exécutés ni mis à jour ;
- un classeur signé numériquement indique que la signature est supprimée à l’enregistrement ;
- un graphique qui ne peut pas être modifié sans perdre de détails est marqué `editBlocked` et affiché et enregistré sans modification ;
- `.xlsb` et `.ods` ne sont pas ouverts ; le format `.xls` ancien est converti en `.xlsx` à l’enregistrement, avec un avertissement pour chaque fonctionnalité que le lecteur binaire ne sait pas décoder ;
- modifier une partie d’une table de données est refusé (« Cannot change part of a data table. »).

## Limites connues de la garantie {#known-limits-of-the-guarantee}

- Une partie conservée octet pour octet n’est pas réécrite lorsque les cellules qu’elle référence se déplacent. Les graphiques, les feuilles de graphique, les caches de tableaux croisés et les autres parties connues de l’écrivain sont corrigés pour les renommages et suppressions de feuilles et pour les déplacements de leurs plages ; une partie XML personnalisé, de connexion ou externe qui mentionne une plage de cellules conserve l’ancien texte.
- La modification concurrente (le domaine de collaboration) transporte le sous-ensemble modélisé dans le document partagé ; un pair qui ne possède pas le paquet d’origine ne peut pas réécrire les parties non modélisées.
- La vérification native qu’Excel ouvre un fichier enregistré sans réparation est manuelle (`src/core/scripts/xlsx-excel-acceptance.ps1`) ; la date de la dernière exécution réussie appartient à l’état du projet, et non à cette liste.
