# Liaisons par framework

Six intégrations natives légères utilisent un seul élément de navigateur et une seule implémentation de `mountViewer`. Chacune est publiée comme son propre paquet autonome ; aucune n'analyse de documents, ne dessine de diagrammes ni ne possède un autre moteur de rendu.

Chaque handle natif expose `createBlankDrawing(options?: CreateVsdxOptions)`.
Le contrôleur partagé crée et accepte la nouvelle source par son cycle de chargement
existant ; les hooks d'état et les événements de document se mettent à jour normalement. Le déplacement du pointeur et la sélection par rectangle s'exécutent dans le même composant web partagé pour toutes les liaisons.

| Framework | Paquet                 | Export                                       |
| --------- | ---------------------- | -------------------------------------------- |
| Vanilla   | `visio-vanilla-viewer` | `mountViewer`                                |
| React     | `visio-react-viewer`   | `VisioViewer`                                |
| Vue 3     | `visio-vue-viewer`     | `VisioViewer`                                |
| Angular   | `visio-angular-viewer` | `VisioViewerComponent` (`visio-viewer-host`) |
| Svelte 5  | `visio-svelte-viewer`  | `VisioViewer` par défaut                     |
| Solid     | `visio-solid-viewer`   | `VisioViewer`                                |

Importez le paquet d'un framework précis pour éviter de charger des environnements d'exécution de frameworks sans rapport. La source des adaptateurs se trouve dans `packages/bindings` du dépôt, et son README répertorie les vérifications qui les protègent.

La mise en place est commune. Le contrat (`VIEWER_CONTRACT`, les clés de propriétés et d'événements, les membres du handle) et la transmission des mises à jour de `mountFrameworkViewer`, qui tient compte des révisions, se trouvent dans `packages/bindings/src/common.ts`. Le mécanisme générique des liaisons (le handle différé, la sélection des propriétés et les fabriques de composants React et Vue) est le kit `ooxml-ui/bindings`. React et Vue sont construits à partir de ces fabriques ; Angular installe les membres du handle à partir de `VIEWER_CONTRACT` ; Svelte est un composant ; Solid est un composant JSX qui monte directement l'élément partagé. `packages/bindings/tests/contract-parity.test.ts` vérifie chaque adaptateur par rapport au contrat : mêmes propriétés, chaque événement passé par la table de rappels et la sortie native, et chaque membre du handle.

## Contrat partagé {#shared-contract}

Tous les adaptateurs acceptent les mêmes propriétés :

- `document` : un `VisioDocument` ou `null` pour le vider
- `pageIndex` : index de page à partir de zéro
- `zoom` : échelle, où `1` correspond à 100 %
- `showToolbar` : visibilité de la barre d'outils
- `canEdit` : autorisation de modification. Le visualiseur complet est modifiable par défaut ; `false` le passe en lecture seule (la sélection, la recherche, le zoom et l'export continuent de fonctionner). Les blocs de construction sont en lecture seule par défaut, voir [blocs de construction](/fr/building-blocks)
- `readOnly` : une restriction propre. `true` rend le dessin en lecture seule et l'emporte sur `canEdit` ; `false` ou omis ne lève que cette restriction
- `ribbonAddIns` : onglets que l'hôte ajoute après Aide (voir [Onglets complémentaires du ruban](/fr/api#ribbon-add-in-tabs))
- `collaboration` : partage le dessin ouvert dans une salle tant qu'il est défini ; `null` le quitte (voir [collaboration](/fr/collaboration))
- `storage` et `customization` : enregistrement et ouverture gérés par l'hôte, et l'objet de personnalisation partagé (voir [Personnalisation](/fr/api#customization) et [Stockage géré par l'hôte](/api#host-owned-storage-and-file-commands))
- `fileProviders` : emplacements cloud pour Fichier > Ouvrir, comme OneDrive (voir [OneDrive](/fr/onedrive))
- `events` : table de rappels typée utilisant des noms d'événements littéraux

La table d'événements contient actuellement `document-load`, `document-change`, `document-error`, `page-change`, `zoom-change`, `shape-select`, `selection-change`, `office-ribbon-add-in`, `collaboration-change` et `file-command`. Elle est dérivée du contrat canonique. Tous les adaptateurs transmettent chaque événement ; aucun n'analyse de document ni ne possède un autre moteur de rendu.

Les mises à jour de propriétés sont partielles. Les instantanés natifs des frameworks sont comparés, de sorte que les rendus sans rapport ne défont pas un chargement impératif ni un cadrage modifié par l'utilisateur. Les propriétés omises ou `undefined` conservent l'état courant du visualiseur. `null` vide explicitement le document. Retirer la propriété `events` supprime les gestionnaires de la table de rappels. Fournissez un nouvel objet document pour un remplacement externe ; la mutation sur place n'est pas un mécanisme de réactivité pris en charge. Les interactions de page et de zoom peuvent modifier l'état interne du visualiseur ; les propriétés demandent des changements lorsque le framework les met à jour.

`document-change` transporte `{ document, dirty, kind }`, où `kind` vaut `edit`, `undo`, `redo`, `remote` (une modification d'un collaborateur) ou `save` (après `markClean()`). Le handle impératif `ViewerHandle` expose `element`, `controller`, `load(bytesOrBlob)` et `fit()`. Les méthodes d'édition comprennent `applyEdits(edits)`, `replacePlainText(pageId, shapeId, text)`, `duplicateSelection()`, `copySelection()`, `cutSelection()`, `pasteSelection()`, `undo()`, `redo()`, `cancelEdit()` et `exportVsdx()`. Une fois les octets exportés enregistrés par l'hôte, celui-ci appelle `controller.markClean()`. La collaboration se pilote par la propriété ; `element.collaborationState`, `element.reconnectCollaboration()` et `element.stopCollaboration()` sont disponibles sur l'`element` du handle. L'état du presse-papiers est accessible via `controller.state.clipboard` à travers chaque abonnement réactif natif ; les méthodes du presse-papiers exigent un geste utilisateur dans le navigateur et prennent en charge le format de forme borné du visualiseur. Elles délèguent à la même implémentation adossée à la source et isolée dans un worker. `ViewerEditState` et `VsdxExportResult` sont des types partagés exportés ; l'état est disponible via `controller.state.edit`. Angular expose la sortie native `documentChange` ; chaque adaptateur prend en charge l'événement `document-change` de la table de rappels. L'accès en lecture et les opérations via un handle de framework échouent clairement avant le montage ou après le démontage. Un contrôleur ou un élément conservé est également détruit lorsque son propriétaire est démonté. Vanilla renvoie la liaison partagée complète avec `update()` et `destroy()` idempotent.

L'éditeur lui-même se monte côté client. Des tests Node séparés vérifient que les six points d'entrée s'importent sans globales du navigateur ; React, Vue, Svelte et Solid rendent aussi des hôtes natifs vides côté serveur. React et Vue disposent de tests qui hydratent leur balisage serveur réel avant de monter l'élément partagé. Le rendu serveur d'Angular et l'hydratation de Solid et Svelte restent non vérifiés, de même que le montage inter-documents. Ces contrôles utilisent les versions installées et n'impliquent pas une couverture complète d'intégration aux frameworks à rendu serveur.

## Utilisation native {#native-usage}

React utilise une référence transmise (forwarded ref) et des propriétés ordinaires :

```tsx
const ref = useRef<ViewerHandle>(null);
<VisioViewer ref={ref} document={diagram} events={{ 'document-error': reportError }} />;
```

Vue expose le même handle sur la ref de son composant et émet aussi les événements natifs en kebab-case :

```vue
<VisioViewer ref="viewer" :document="diagram" @document-error="reportError" />
```

Angular expose des entrées et sorties natives. Ajoutez le composant autonome aux imports de l'application :

```html
<visio-viewer-host [document]="diagram" (documentError)="reportError($event)" />
```

Les sorties Angular sont `documentLoad`, `documentChange`, `documentError`, `pageChange`, `zoomChange`, `shapeSelect`, `selectionChange`, `ribbonAddIn`, `collaborationChange` et `fileCommand`. `fileCommand` se contente de signaler la commande ; pour l'annuler, un hôte Angular renvoie `false` depuis le rappel `events['file-command']`. La table de rappels `events` complète est aussi acceptée comme entrée. Vue et Angular appellent le gestionnaire de la table de rappels puis émettent leur sortie native ; utilisez l'une ou l'autre API, sauf si les deux notifications sont voulues.

Pour les grands modèles de document immuables, utilisez `shallowRef`/`markRaw` avec Vue et `$state.raw` avec Svelte, plutôt que de placer dans un proxy profond un objet d'options nouvellement construit à chaque mise à jour. Réenvelopper un document dans un nouveau proxy réactif change son identité et demande un remplacement externe. La mutation sur place n'est pas prise en charge.

Svelte utilise des props de rappel natives via `events`. Son composant exporte `load`, `fit` et `getHandle`, afin que le handle typé complet soit accessible sans exposer le DOM privé :

```svelte
<VisioViewer bind:this={viewer} document={diagram}
  events={{ 'document-error': reportError }} />
```

Solid fournit un rappel de ref qui reçoit le handle au montage et `undefined` au démontage ou au remplacement de la ref :

```tsx
<VisioViewer
	document={diagram()}
	viewerRef={setViewer}
	events={{ 'document-error': reportError }}
/>
```

Vanilla ne gère que l'enfant qu'il crée :

```ts
const viewer = mountViewer(host, { document: diagram, zoom: 1 });
viewer.update({ events: { 'document-error': reportError } });
// Later:
viewer.destroy();
```

Les points d'accroche de style sont natifs à chaque framework : `className`/`style` pour React, attributs hérités pour Vue, attributs de l'hôte pour Angular, `class`/`style` pour Svelte, et `class`/`style` pour Solid. L'application doit donner à l'hôte une hauteur utilisable. Le style partagé du visualiseur et l'accessibilité sont portés par l'élément personnalisé.
