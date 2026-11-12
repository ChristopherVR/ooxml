# Unités de mesure dans le modèle de document

Le `DocumentModel` stocke les mesures dans les unités qu'utilise WordprocessingML sur disque. Depuis l'introduction des marques d'unités, ces unités font partie des types TypeScript : passer une taille de police là où une longueur est attendue, ou une valeur en pixels non arrondie là où des twips entiers sont requis, est une erreur de compilation au lieu d'un fichier corrompu.

Les marques n'existent qu'à la compilation. Un `Twips` est un simple nombre JavaScript à l'exécution, et reste assignable à `number` : l'arithmétique, les comparaisons et la sérialisation JSON ne changent pas.

## Marques {#brands}

Toutes les marques sont exportées depuis `docx-core`.

| Type           | Unité                               | Type de schéma          | Constructeur (strict) | Constructeur d'arrondi |
| -------------- | ----------------------------------- | ----------------------- | --------------------- | ---------------------- |
| `Twips`        | 1/20 de point, non négatif          | `ST_TwipsMeasure`       | `twips(n)`            | `roundTwips(n)`        |
| `SignedTwips`  | 1/20 de point, peut être négatif    | `ST_SignedTwipsMeasure` | `signedTwips(n)`      | `roundSignedTwips(n)`  |
| `HalfPoints`   | 1/2 point (taille `w:sz`)           | `ST_HpsMeasure`         | `halfPoints(n)`       | `roundHalfPoints(n)`   |
| `EighthPoints` | 1/8 de point (épaisseur de bordure) | `ST_EighthPointMeasure` | `eighthPoints(n)`     | `roundEighthPoints(n)` |
| `Emu`          | 1/914400 de pouce (DrawingML)       | `ST_Coordinate`         | `emu(n)`              | `roundEmu(n)`          |

`Twips` est un sous-type de `SignedTwips` : une mesure non signée peut être assignée à un champ signé, mais jamais l'inverse.

**Les constructeurs stricts** lèvent une `RangeError` sauf si la valeur est déjà un entier entier et sûr (et, pour `twips`, non négatif). Les analyseurs et les tests les utilisent.

**Les constructeurs d'arrondi** servent aux valeurs calculées (déplacements de règle, calculs de zoom, conversions en pixels). Ils arrondissent à la moitié en s'éloignant de zéro et ne lèvent une `RangeError` que pour une entrée `NaN` ou infinie. `roundTwips` et `roundEighthPoints` ramènent les résultats négatifs à zéro.

**Les conversions depuis des unités réelles** arrondissent à l'unité cible entière la plus proche : `twipsFromPoints`, `signedTwipsFromPoints`, `twipsFromPixels` (15 twips par pixel CSS), `signedTwipsFromPixels`, `twipsFromInches`, `halfPointsFromPoints`, `eighthPointsFromPoints`, `emuFromPixels`.

**Les conversions vers des unités réelles** renvoient des nombres simples : `twipsToPoints`, `twipsToPixels`, `twipsToInches`, `halfPointsToPoints`, `eighthPointsToPoints`, `emuToPixels`.

Il n'y a pas de transtypage du genre `as Twips` en dehors de `units.ts` ; utilisez un constructeur.

## Quelles propriétés du modèle portent quelle marque {#which-model-fields-carry-which-brand}

Les noms des champs ne changent pas.

- **Paragraphe** (`Paragraph`, `ParagraphFormatting`, définitions de style) : `spacingBeforeTwips`, `spacingAfterTwips`, `firstLineTwips`, `hangingTwips` sont des `Twips` ; `lineSpacingTwips`, `indentLeftTwips`, `indentRightTwips`, `indentStartTwips`, `indentEndTwips` sont des `SignedTwips`. `TabStop.posTwips` est un `SignedTwips`.
- **Niveaux de numérotation et étiquettes de liste** : `indentLeftTwips` est un `SignedTwips` ; `hangingTwips` et `firstLineTwips` sont des `Twips`.
- **Sections** : `pageWidthTwips`, `pageHeightTwips`, `marginLeftTwips`, `marginRightTwips`, `headerDistanceTwips`, `footerDistanceTwips`, `gutterTwips` et les `widthTwips`/`spacingTwips` des colonnes sont des `Twips` ; `marginTopTwips` et `marginBottomTwips` sont des `SignedTwips` (Word autorise des marges haute et basse négatives).
- **Tableaux** : `Table.grid` est un `Twips[]` ; `Table.widthTwips`, `TableCell.widthTwips` et `TableRowProperties.heightTwips` sont des `Twips` ; `Table.indentTwips` est un `SignedTwips` (Word écrit couramment des retraits de tableau négatifs) ; `TableCellMargins.top/bottom/left/right` sont des `SignedTwips`.
- **Bordures** : `TableBorderSide.sizeEighthPoints` est un `EighthPoints`.
- **Segments de texte** : `TextRun.characterSpacingTwips` est un `SignedTwips`.
- **Entrée de mise en page** (`ooxml-core/docx/layout`) : les champs `*Twips` de paragraphe de `LayoutParagraph` utilisent les mêmes marques que le modèle.

### Champs volontairement sans marque {#fields-that-are-deliberately-not-branded}

- `TextRun.fontSize` est exprimé en **points** et peut être fractionnaire (`10.5`). `HalfPoints` est la valeur `w:sz` telle qu'elle est stockée sur disque ; l'analyseur la convertit. Utilisez `halfPointsFromPoints` pour réécrire une taille.
- `TableBorderSide.spacePoints` est en **points** entiers (`ST_PointMeasure`).
- `widthPx`, `heightPx`, `offsetXPx`, `offsetYPx` et tous les champs `*Px` de mise en page sont des **pixels CSS à 96 dpi**. Ce sont des valeurs dérivées de l'espace du périphérique, souvent fractionnaires : elles restent donc des nombres simples. Les dimensions des images sont stockées en pixels (dérivés des EMU) plutôt qu'en `Emu`.
- `DocumentModel.page` (`width`, `height`, marges) est en pixels CSS ; `SectionProperties` reste la source faisant autorité pour les twips.

## Construire des valeurs de modèle {#building-model-values}

Les champs du modèle prennent des nombres marqués ; le code qui construit ou modifie des objets du modèle doit donc les construire :

```ts
import { twips, eighthPoints, roundSignedTwips, twipsFromPixels } from 'docx-core';
paragraph.spacingAfterTwips = twips(240);
table.grid = [twips(3000), twips(3000)];
side.sizeEighthPoints = eighthPoints(8);

// Computed values: round explicitly at the boundary
paragraph.indentLeftTwips = roundSignedTwips(dragPx * 15);
section.pageWidthTwips = twipsFromPixels(page.width);
```

La lecture n'est pas affectée : les valeurs marquées sont des nombres, donc `paragraph.spacingAfterTwips ?? 0` et `twipsToPixels(x)` fonctionnent comme d'habitude. Lorsque vous avez besoin d'un nombre simple dans un emplacement marqué, à partir d'une source dont vous êtes sûr qu'elle est entière (par exemple une valeur relue depuis votre propre JSON), passez-la par `twips()` ou `signedTwips()` pour qu'une valeur incorrecte échoue bruyamment. `twips(n)` rejette les nombres négatifs ; utilisez `signedTwips(n)` pour les valeurs signées.

## L'analyse suit la signature du schéma {#parsing-follows-the-schema-signedness}

Comme le modèle distingue désormais les longueurs signées et non signées, les analyseurs appliquent le type du schéma attribut par attribut. Les valeurs négatives de `w:spacing/@before|@after`, `w:ind/@firstLine|@hanging`, `w:pgMar/@left|@right|@header|@footer|@gutter`, de la grille de tableau, des largeurs et des hauteurs de ligne sont rejetées (le champ reste indéfini), tandis que `w:ind/@left|@right|@start|@end`, `w:spacing/@line`, `w:tab/@pos`, `w:pgMar/@top|@bottom`, `w:tblInd` et `w:spacing` (espacement des caractères) sont conservés. `validateModel` (le validateur exécuté avant l'enregistrement) applique la même classification et rejette les valeurs non entières.

## Frontière des attributs de l'éditeur {#editor-attribute-boundary}

Les attributs de nœud ProseMirror ne sont pas typés. L'éditeur les convertit en modèle en un seul endroit (`attr-units.ts` dans `web-component`) : les valeurs sont arrondies à des twips entiers, et les valeurs non numériques, non finies ou (pour les champs non signés) négatives sont écartées au lieu d'être écrites.
