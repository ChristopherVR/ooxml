# Prise en main {#getting-started}

Ouvrez un diagramme en local, comprenez les limites actuelles et intégrez le même visualiseur dans votre application.

::: warning Prise en charge initiale du visualiseur, pas une certification Visio
Il s'agit d'une version bêta publique. Ne considérez pas l'ouverture réussie d'un fichier comme la preuve que chaque forme, formule ou fonctionnalité du format est représentée fidèlement. Consultez le [registre des capacités](/parity).
:::

## Essayer le terrain de jeu local {#try-the-local-playground}

1. [Ouvrez le terrain de jeu](/demo/){target="_self"} pour voir un document d'exemple original.
2. Choisissez un fichier `.vsdx` local avec le sélecteur de fichiers ou la zone de dépôt.
3. Changez de page, ajustez le zoom, ajustez la page à la fenêtre et sélectionnez une forme.
4. Consultez les notes de compatibilité. Comparez les fichiers importants avec leur rendu dans Microsoft Visio.

Aucun document n'est envoyé. Les commandes permettent la recherche, l'affichage des calques, l'édition expérimentale de texte brut, l'annulation/rétablissement bornés, l'export SVG et le téléchargement explicite de copies VSDX. Les instantanés d'impression ne sont qu'une préparation disponible par l'API ; il n'y a pas de boîte de dialogue d'impression. Le dessin général, l'export PDF et la plupart des fonctionnalités des anciens fichiers binaires `.vsd` ne font pas partie de cette version.

Les autres versions du même terrain de jeu pour chaque framework sont répertoriées sur [la page des démos](/demos).

## Installer un paquet pour votre framework {#install-a-framework-package}

Installez un paquet de visualiseur pour votre framework. Chacun inclut le visualiseur partagé, son adaptateur, ses workers et l'API de document.

::: code-group

```bash [React]
npm install visio-react-viewer react
```

```bash [Vue]
npm install visio-vue-viewer vue
```

```bash [Angular]
npm install visio-angular-viewer
```

```bash [Svelte]
npm install visio-svelte-viewer svelte
```

```bash [Solid]
npm install visio-solid-viewer solid-js
```

```bash [Vanilla JS]
npm install visio-vanilla-viewer
```

:::

| Framework | Paquet                 | Export                   |
| --------- | ---------------------- | ------------------------ |
| Vanilla   | `visio-vanilla-viewer` | `mountViewer`            |
| React     | `visio-react-viewer`   | `VisioViewer`            |
| Vue 3     | `visio-vue-viewer`     | `VisioViewer`            |
| Angular   | `visio-angular-viewer` | `VisioViewerComponent`   |
| Svelte 5  | `visio-svelte-viewer`  | `VisioViewer` par défaut |
| Solid     | `visio-solid-viewer`   | `VisioViewer`            |

Les six adaptateurs transmettent le même contrôleur, les mêmes propriétés, les mêmes événements et le même handle impératif via le visualiseur partagé. Consultez les [guides par framework](/frameworks/react) et le [contrat des liaisons](/bindings).

## Monter un visualiseur {#mount-a-viewer}

Le conteneur doit avoir une hauteur. Un `File`, un `Blob`, un `Uint8Array` ou un `ArrayBuffer` du navigateur peut être passé à `load`.

```ts
import { mountViewer } from 'visio-vanilla-viewer';

const viewer = mountViewer(host, {
	events: { 'shape-select': (shape) => console.log(shape) },
});

await viewer.load(file); // A local File or bytes
viewer.fit();

// When your view unmounts:
viewer.destroy();
```

## Lancer l'espace de développement {#run-the-development-workspace}

Utilisez Node.js 22.12 ou une version plus récente. Les builds normaux installent depuis npm l'API `ooxml-core/visio` publiée.

```bash
npm ci --ignore-scripts
npm ci --prefix packages/bindings --ignore-scripts
npm run check
npm run dev
```

Utilisez l'URL locale affichée par Vite. Cette commande démarre un serveur de développement ; elle ne publie pas le site.

## Édition expérimentale de texte brut {#editing}

Chargez un paquet source local, sélectionnez une forme prise en charge et utilisez ses commandes de texte. `replacePlainText(pageId, shapeId, text)`, `undo()` et `redo()` exposent les mêmes opérations via chaque handle. Un modèle affecté avec `document` n'a pas d'octets de paquet et reste en lecture seule.

L'édition refuse le texte lié à un masque, le texte enrichi, les champs, les paquets signés et le contenu de macros. L'historique est borné et peut écarter d'anciens états d'annulation. Les caches de formules enregistrés ne sont pas recalculés. Le dessin général et l'édition des styles ne sont pas implémentés.

`exportVsdx()` renvoie les octets du paquet d'origine ou modifié, ainsi que des diagnostics, pour un téléchargement explicite de copie. Il n'écrase jamais le fichier d'entrée. La réouverture dans Visio natif et la fidélité d'un aller-retour sans perte restent non vérifiées.

## Exporter un SVG statique {#export}

`viewer.exportSvg()` renvoie la chaîne SVG de la page courante, ses dimensions, sa taille en octets et les diagnostics de compatibilité. Il laisse inchangés la sélection, le zoom et le document source. Le bouton Export SVG du terrain de jeu effectue un téléchargement local explicite.

Le contenu d'arrière-plan pris en charge et les images raster validées sont intégrés. Les polices ne sont pas intégrées ; le retour à la ligne et les autres approximations restent visibles dans la description et les métadonnées du SVG. L'export est limité à 16 Mio, avec des contrôles précoces prudents des ressources. Les appelants peuvent abaisser la limite avec `{ maxBytes }`. Ce n'est pas une sortie VSDX modifiable. Les rastériseurs en aval doivent plafonner indépendamment les dimensions et la surface en pixels.

## Préparer des instantanés de pages {#prepare-page-snapshots}

`createPrintSnapshot()` renvoie par l'API des artefacts de pages SVG immuables et des diagnostics. Il préserve les dimensions du dessin, mais n'applique pas les réglages d'imprimante natifs, n'ouvre pas de boîte de dialogue d'impression et n'implémente pas l'export PDF. Cela prépare les données d'un futur flux d'impression.

## Propriétés et événements {#properties-and-events}

| Surface      | Contrat actuel                                                                                                                                 |
| ------------ | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| Propriétés   | `document`, `pageIndex`, `zoom`, `showToolbar`, `ribbonAddIns`                                                                                 |
| Événements   | `document-load`, `document-change`, `document-error`, `page-change`, `zoom-change`, `shape-select`, `selection-change`, `office-ribbon-add-in` |
| Cycle de vie | `update`, `load`, `fit`, `exportSvg`, `createPrintSnapshot`, `replacePlainText`, `undo`, `redo`, `cancelEdit`, `exportVsdx`, `destroy`         |

Montez le visualiseur une fois l'hôte disponible et appelez `destroy()` lorsque la vue du framework est démontée. Détruire plusieurs fois est sans danger. Les mises à jour partielles laissent inchangées les propriétés omises. La référence complète est l'[API du visualiseur](/fr/api).

## Vérification avant adoption {#verification-before-adoption}

Exécutez `npm run check` et `npm run test:browser` dans la copie du visualiseur. Les tests de l'analyseur du cœur relèvent du moteur frère. Ces vérifications n'établissent que les comportements qu'elles affirment ; elles n'établissent pas la parité visuelle avec Microsoft Visio.
