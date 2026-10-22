# Architecture et responsabilités

## Sens des dépendances

```text
published, one per framework (self-contained bundles)
  docx-react-viewer | -vue-viewer | -angular-viewer | -svelte-viewer | -solid-viewer | -vanilla-viewer
  ------------------------------------------------------------------------------------------------
  internal, private, inlined into every bundle:
        bindings (lifecycle and events only)
                         |
                  web-component editor (canvas text measurer)
                         |
        ooxml-core (external, published)
   /docx/layout   /docx/load (format detection, legacy .doc;
   /collab        ole2 codecs inlined by the core)
        ^
   docx-core (external): thin re-export of ooxml-core/docx and /docx/embedded
```

Seuls `docx-core` et les six paquets `*-viewer` sont publiés. `web-component` et `bindings` sont des paquets d'espace de travail `private` : ils constituent la source partagée de l'éditeur, jamais une cible d'installation npm. `scripts/build-packages.mjs` les regroupe dans chaque paquet de framework, de sorte qu'une archive n'importe que `docx-core`, `ooxml-core`, les bibliothèques ProseMirror et ses dépendances de framework (peers). `scripts/check-published-refs.mjs` et le test de fumée d'empaquetage font échouer la construction si un paquet interne ou `ole2` fuit dans une archive.

Le modèle de document canonique, l'analyseur DOCX et le sérialiseur résident dans la zone `docx` du paquet `ooxml-core` ; `docx-core` est un fin point d'entrée publié qui le réexporte. Le modèle modifiable est une projection volontairement réduite d'OOXML. Une session chargée détient le paquet d'origine ainsi que le modèle de référence, ce qui permet à l'enregistrement des modifications prises en charge de préserver les parties extérieures à cette projection. Un enregistrement sans changement renvoie les octets d'origine.

`ooxml-core/docx/load` détecte les signatures de conteneur et choisit entre DOCX et DOC hérité. Son lecteur d'ancien format extrait les paragraphes du corps principal des fichiers Word 97-2003, s'appuie sur les codecs ole2 (intégrés au cœur) pour le CFB et les structures binaires de Word, et traduit le texte extrait vers le modèle commun. Sa voie d'édition n'accepte le texte brut des paragraphes que si les contrôles de sécurité sont satisfaits. `ooxml-core/docx/layout` est le moteur de pagination sans DOM ; l'éditeur lui injecte son `TextMeasurer` fondé sur le canvas. Les utilitaires de collaboration en temps réel (validation, identité, règles d'ordonnancement) proviennent de `ooxml-core/collab` ; le client et l'autorité `prosemirror-collab`, le plugin de présence et les décorations restent dans l'éditeur (`ooxml-ui/docx`).

L'élément `<docx-editor>` (source dans `ooxml-ui/docx`, que `web-component` réexporte) est la seule interface d'édition. Il convertit le modèle en ProseMirror, applique les transactions d'édition et répare les identités de blocs, émet les changements de modèle et appelle la fonction d'enregistrement de la session chargée. La surface de commandes actuelle prend en charge l'édition de paragraphes, la mise en forme directe de base du texte, l'alignement, les tableaux simples et le zoom sur une surface d'édition continue. La mise en page Impression est une vue paginée approximative issue de `ooxml-core/docx/layout`, et non la pagination de Word. Le Shadow DOM isole les styles de la barre d'outils et des pages. L'enregistrement n'a lieu que dans le navigateur et est idempotent. Des tests dans le navigateur exercent le même contrat à travers chaque adaptateur.

`bindings` ne gère que l'intégration aux frameworks : montage/démontage, mises à jour de propriétés, transmission d'événements et références impératives. Les nouvelles fonctionnalités de l'éditeur relèvent du composant web ; une modification d'un adaptateur ne se justifie que par une contrainte de cycle de vie ou d'API du framework.

## Partage entre les deux visionneuses

Le lecteur/écrivain CFB et les structures binaires internes de Word sont canoniques dans le paquet `ole2` (`@christophervr/ole2`). Le modèle DOCX, l'analyseur, l'écrivain et l'API de DOCX intégré (`docx-core/embedded`) sont canoniques dans `ooxml-core/docx` (réexporté par `docx-core`). Aucune implémentation DOCX moderne n'a sa place dans `ole2`.

Le paquet partagé `ole2` n'a aucune dépendance envers React, Word ou l'interface de PowerPoint, et ne gère que le CFB et les formats binaires hérités de Word. Le DOCX est de l'OOXML et relève de `ooxml-core/docx`.

## Frontière de construction et de publication

Sept paquets sont publiés (`docx-core` et les six paquets de framework ; voir la [publication](/releasing)), avec des publications npm automatisées. Le paquet CFB partagé ne contient aucun codec DOCX.

`docx-core` reste une véritable dépendance de chaque paquet de framework au lieu d'être regroupé : les applications importent directement `createDocument` et les types du modèle depuis ce paquet, de sorte qu'une copie partagée garantit que `DocumentModel` est identique des deux côtés. Les bibliothèques ProseMirror sont elles aussi de véritables dépendances, non regroupées, afin qu'un gestionnaire de paquets puisse dédupliquer une seule copie pour l'éditeur et pour tout autre code ProseMirror de l'application. `ooxml-core` est également une véritable dépendance (il apporte la mise en page, le chargement et la collaboration). Tout le reste, interne, est regroupé afin que les consommateurs en aval n'aient besoin d'aucune extraction de sources voisines ni d'aucun paquet non publié ; ole2 est intégré dans le cœur et n'est une dépendance d'aucun élément ici.

Consultez la feuille de route de la prise en charge avant de considérer ceci comme une parité fonctionnelle avec Word. Des constructions de paquets réussies et des routes de démo ne démontrent ni le rendu de Word, ni sa pagination, ni la parité d'aller-retour.
