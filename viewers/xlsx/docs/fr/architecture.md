# Architecture et responsabilités

## Sens des dépendances

```text
published, one per framework (self-contained bundles)
  @christophervr/xlsx-react-viewer | xlsx-vue-viewer | xlsx-angular-viewer
  xlsx-svelte-viewer | xlsx-solid-viewer | xlsx-vanilla-viewer
  ----------------------------------------------------------------------------
  internal, private, inlined into every bundle:
        bindings (lifecycle and events only)
                         |
               <xlsx-editor> web component (grid painter, ribbon, dialogs)
                         |
        ooxml-core (external, published)                 ooxml-ui (external)
   /xlsx       model, formula engine, layout, edit,      office-ui-* controls
               read/write .xlsx
   /xlsx/load  format detection, .xls (ole2 codecs inlined), CSV
        ^
   @christophervr/xlsx-core (external): thin re-export of ooxml-core/xlsx and /xlsx/load
```

Seuls `@christophervr/xlsx-core` et les six paquets de framework sont publiés. `web-component` et `bindings` sont des paquets d’espace de travail `private` : ils constituent la source partagée de l’éditeur, jamais une cible d’installation npm. `scripts/build-packages.mjs` les regroupe dans chaque paquet de framework, de sorte qu’une archive n’importe que `@christophervr/xlsx-core`, `ooxml-core`, `ooxml-ui` et sa dépendance homologue de framework. `scripts/check-published-refs.mjs` et le test de fumée du paquet font échouer la compilation si un paquet interne ou `ole2` fuit dans une archive.

## Le cœur porte la logique

Tout ce qui n’est pas de l’interface vit dans la zone `xlsx` de `ooxml-core` (source dans le dépôt `ChristopherVR/ooxml`) :

- le modèle de classeur (`Workbook`, `Worksheet`, `Cell`, styles de cellule résolus et dédupliqués),
- les adresses et les plages (`parseAddress`, `formatRange`, ...),
- les formats numériques (`formatValue`, `parseCellInput`),
- l’analyseur de formules et le moteur de calcul (`createCalcEngine`, suivi des dépendances, tableaux dynamiques),
- les commandes d’édition avec annuler et rétablir (`createEditSession`),
- la logique de vue sans DOM : métriques de la grille, cellules visibles, vues de cellule, mises en forme conditionnelles, modèles de vue des graphiques,
- le lecteur et l’écrivain `.xlsx` et, sous `/xlsx/load`, la détection de format, le lecteur `.xls` ancien format (via les codecs ole2 partagés) et le CSV.

Un classeur chargé conserve dans `workbook.source` les parties que le modèle ne représente pas, afin que l’enregistrement les restitue.

## Le composant web dessine

`packages/web-component` est la seule interface d’édition : il crée une session d’édition sur le classeur, dessine ce que renvoient les fonctions de mise en page du cœur, achemine le clavier, la souris et le ruban vers les commandes d’édition, et gère les styles, les langues et les boîtes de dialogue. Le Shadow DOM limite la portée de ses styles ; l’enregistrement n’a lieu que dans le navigateur et est idempotent. Les tests de navigateur exercent le même contrat à travers chaque liaison.

`packages/bindings` ne gère que l’intégration aux frameworks : montage et démontage, mises à jour des propriétés, transmission des événements et handles impératifs. Les nouvelles fonctionnalités de l’éditeur appartiennent au composant web (ou au cœur) ; une modification d’une liaison ne se justifie que par une contrainte de cycle de vie ou d’API du framework.

## Frontière de compilation et de publication

`@christophervr/xlsx-core` reste une vraie dépendance de chaque paquet de framework plutôt que d’être regroupé : les applications y importent directement `createWorkbook` et les types du modèle, si bien qu’une seule copie partagée garde `Workbook` identique des deux côtés. `ooxml-core` et `ooxml-ui` sont aussi de vraies dépendances. Tout le reste de l’interne est regroupé ; ole2 est intégré au cœur et n’est une dépendance de rien ici. Voir [publication](/releasing).
