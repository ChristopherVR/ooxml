# API du visualiseur {#viewer-api}

L'API indépendante du framework qui se trouve derrière chaque paquet de visualiseur. Chaque paquet de framework
(`visio-react-viewer`, `visio-vue-viewer`, `visio-angular-viewer`,
`visio-svelte-viewer`, `visio-solid-viewer`, `visio-vanilla-viewer`) transmet
les mêmes méthodes de handle et les mêmes événements ; consultez `packages/bindings/README.md` pour les
props, événements et cycles de vie natifs de chaque adaptateur. Le rendu et l'édition sont
des fonctionnalités bêta ; le [registre des capacités](/parity) répertorie ce qui manque.

## Monter et charger {#mount-and-load}

```ts
import { mountViewer } from 'visio-vanilla-viewer';
const viewer = mountViewer(container, {
	zoom: 1,
	events: { 'document-error': (error) => console.error(error) },
});
await viewer.load(file); // Blob, Uint8Array or ArrayBuffer
viewer.fit();
viewer.update({ pageIndex: 1 });
const snapshot = viewer.exportSvg(); // SVG string, dimensions, byteLength and diagnostics
const prepared = viewer.createPrintSnapshot(); // Frozen current-page artifacts, no printing
viewer.destroy(); // idempotent; later commands reject
```

L'import du module est sans danger pour le SSR ; le montage et l'enregistrement de l'élément ne
fonctionnent que dans le navigateur. `ViewerController` n'utilise pas le DOM, si bien qu'un autre moteur de rendu peut s'abonner
au même état et aux mêmes événements. `renderPage` est le port du moteur de rendu SVG. Consultez
l'[architecture](/fr/architecture) pour la répartition des modules.

## Propriétés et événements {#properties-and-events}

Chaque binding accepte les mêmes propriétés et signale les mêmes événements (le contrat est `ViewerProperties` et `ViewerEvents` dans `ooxml-ui/visio` ; les [bindings](/bindings) listent les formes natives).

| Surface    | Contrat                                                                                                                                                                |
| ---------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Propriétés | `document`, `pageIndex`, `zoom`, `showToolbar`, `canEdit`, `ribbonAddIns`, `collaboration`                                                                             |
| Événements | `document-load`, `document-change`, `document-error`, `page-change`, `zoom-change`, `shape-select`, `selection-change`, `office-ribbon-add-in`, `collaboration-change` |

- `canEdit` (booléen) est l'autorisation de modification. Le visualiseur complet est modifiable par défaut et `canEdit: false` le met en lecture seule ; les [blocs de construction](/building-blocks) sont en lecture seule tant que l'hôte ne passe pas `true`. Un dessin sans octets de paquet (un modèle passé comme `document`, un ancien `.vsd`) est en lecture seule quelle que soit la valeur de `canEdit`.
- `document-change` fournit `{ document, dirty, kind }`. `kind` vaut `edit`, `undo` ou `redo` pour les modifications de cette fenêtre, `remote` quand la modification d'un collaborateur a été appliquée, et `save` après `markClean()`.
- `markClean()` (sur `viewer.controller` et sur les blocs de construction) déclare enregistré le contenu actuel. Un hôte qui enregistre `exportVsdx()` lui-même l'appelle ensuite : `edit.dirty` passe à `false`, un `document-change` de type `save` est émis, et les modifications et annulations suivantes sont mesurées à partir de ce point. Il ne fait rien pour un dessin qui ne peut pas être réenregistré.
- `collaboration` (`VisioCollaborationOptions | null`) partage le dessin ouvert dans une salle tant qu'elle est définie. `collaboration-change` fournit un `CollaborationState` quand le partage démarre ou s'arrête et quand l'état, la synchronisation ou les personnes changent. Sur l'élément (`viewer.element` d'un handle), `collaborationState` lit l'état actuel, et `reconnectCollaboration()` et `stopCollaboration()` agissent dessus. Voir [Collaboration](/fr/collaboration).

## Sélectionner des formes {#select-shapes}

Chaque handle monté ou natif expose `selectShapes(targets)`, `selectAll()` et
`clearSelection()`. Une cible a la forme `{ id, name, pageId? }` ; l'omission de `pageId` désigne
la page courante. `controller.state.selectedShapes` est un tableau ordonné et immuable,
dont le premier élément est le `selectedShape` principal. L'événement existant
`shape-select` signale cet élément principal ; `selection-change` signale le
tableau complet. Ces états sont disponibles via le hook, le composable, le signal
ou le store natif de chaque framework. Vue émet `selection-change` ; Angular expose
`selectionChange` ; les six adaptateurs acceptent `events['selection-change']`.

Un clic avec Maj, Ctrl ou Meta bascule l'état d'une forme. Les clics sur l'arrière-plan et la touche Échap effacent
la sélection. Les cibles masquées ou supprimées sont retirées, la sélection d'un groupe exclut
ses descendants, et le changement de page ou de document efface la sélection. Tout sélectionner inclut
les formes visibles du premier plan. Les volets de texte et de géométrie à cible unique exigent
exactement une forme sélectionnée. La mise en forme de l'onglet Accueil et la suppression s'appliquent de façon atomique à
plusieurs formes sélectionnées ; la protection de la source de n'importe quelle cible fait rejeter le lot.

## Créer un dessin vierge {#create-a-blank-drawing}

Chaque handle monté ou natif et l'élément partagé exposent
`createBlankDrawing(options?: CreateVsdxOptions): Promise<void>`. Le contrôleur
propose la même méthode. `width` et `height`, facultatifs, utilisent des pouces physiques ;
les valeurs par défaut sont 8,5 sur 11 avec une page vide et une échelle de dessin de 1:1. Fichier > Nouveau >
Dessin vierge et Ctrl+N appellent le même cycle de vie. Les commandes modifiables conservent leur
comportement de raccourci natif.

Une demande Nouveau acceptée remplace le document, la sélection et l'historique, et l'élément
le nomme `New drawing.vsdx`. Les octets sources peuvent être exportés immédiatement ;
`dirty` enregistre les modifications ultérieures. Les chargements remplacés, l'annulation et la libération
ne peuvent pas accepter un résultat tardif. Les changements d'état passent par l'API d'état réactif
existante de chaque liaison.

## Déplacement au pointeur et sélection par rectangle {#pointer-movement-and-marquee-selection}

Faire glisser une forme 2D locale éligible déplace la sélection courante d'une translation
commune. Un seuil de quatre pixels client préserve les clics ordinaires. L'aperçu
ne modifie que la présentation SVG ; le relâchement soumet une seule modification atomique de la source. Échap,
la perte de capture du pointeur, les changements de source, de page ou de sélection et les changements d'outil l'annulent.
Le déplacement utilise des coordonnées de page physiques et les convertit en unités de dessin dans le cœur,
en préservant la rotation enregistrée et les points d'ancrage locaux. Les protections de la source restent déterminantes.

Faire glisser sur un espace vide de la page sélectionne les formes visibles du premier plan entièrement incluses,
selon leur largeur/hauteur transformées. Maj, Ctrl ou Meta ajoute à la
sélection courante. Le déplacement au pointeur exclut actuellement les masques, groupes, calques, objets étrangers
et connecteurs collés. Les poignées de redimensionnement et le routage des connecteurs sont des
capacités distinctes ; les poignées de redimensionnement sont décrites ci-dessous, tandis que le routage reste en dehors de
la portée de ce geste.

## Poignées de redimensionnement et Taille et position {#resize-handles-and-size-position}

L'outil pointeur partagé affiche huit poignées de redimensionnement pour une forme 2D locale
éligible. Les poignées de bord modifient une dimension ; les poignées de coin modifient les deux. Le
bord ou coin opposé reste fixe dans les axes locaux, pivotés et réfléchis, enregistrés de la forme.
Les points d'ancrage locaux décentrés enregistrés sont mis à l'échelle proportionnellement lorsque la preuve de source
le permet. Pendant le glissement, seul le cadre de sélection affiche un aperçu des nouvelles limites ; le relâchement
soumet une modification de la source et une étape d'annulation. Franchir la poignée opposée plafonne
la dimension à la plus petite valeur entre sa taille d'origine et 1/16 de pouce physique, sans
changer les indicateurs de symétrie. L'aimantation, les gestes avec modificateur de proportions, l'aperçu complet du contenu et
le redimensionnement de plusieurs formes ne sont pas implémentés.

Affichage > Volets de tâches > Taille et position ouvre le volet numérique partagé. X/Y sont le
point de rotation en pouces de dessin depuis l'origine en bas à gauche, Y augmentant
vers le haut. La largeur et la hauteur utilisent aussi des pouces de dessin, indépendamment de l'échelle
physique de la page. L'angle est en degrés, positif dans le sens antihoraire. Entrée ou la sortie d'un champ
valide ce champ une seule fois ; Échap restaure sa valeur enregistrée. Les changements de largeur/hauteur
dans ce volet maintiennent fixe le point de rotation, contrairement aux poignées ancrées du pointeur.
Une saisie identique à la valeur actuelle laisse inchangés les octets sources, les formules et l'historique.
Les brouillons survivent au zoom et aux autres modifications sur la même source et la même sélection ; le remplacement ou
une sélection différente les efface.

Les deux commandes exigent un VSDX adossé à la source et une forme 2D locale ordinaire, sans calque,
non collée, avec des points d'ancrage enregistrés. Les groupes, masques, objets étrangers et la sélection
multiple ne sont pas disponibles. L'éligibilité de la scène est préliminaire : le cœur vérifie
la géométrie de la source, les protections, les formules et les dépendances à la soumission. Un refus
affiche une erreur sans modifier les octets sources ni l'historique. L'annulation, la libération
et une intention plus récente de source, de page ou de sélection empêchent l'acceptation d'une modification en attente.

L'assistant du cœur `visioSizePositionState(page, shapeId)` fournit un état numérique
figé. `visioSizePositionCommand(page, shapeId, field, value)` renvoie des
modifications ordinaires, un tableau vide pour une saisie sans effet exacte, ou `undefined` pour un candidat invalide ou
non pris en charge. Les commandes personnalisées peuvent utiliser ces assistants avec l'API
d'édition atomique existante. Les chaînes d'unités, la saisie de formules et le changement de la position
du point d'ancrage ne sont pas pris en charge par le volet numérique.

## Rechercher du texte dans le diagramme {#find-diagram-text}

Les commandes de recherche partagées trouvent du texte littéral dans les formes et pages visibles.
Les hôtes programmatiques utilisent `controller.setSearchQuery(query)`, `nextSearchResult()`,
`previousSearchResult()` et `selectSearchResult(index)` ; la requête et les résultats
immuables sont disponibles dans `controller.state.search`. Interroger laisse la
sélection inchangée jusqu'à une navigation explicite. Un nouveau document efface la recherche, et
les actions réentrantes plus récentes de l'hôte remplacent l'ancienne navigation dans les résultats.

La recherche renvoie un résultat par forme correspondante, avec un aperçu de texte borné.
Elle utilise la comparaison en minuscules d'ECMAScript, pas d'expression régulière ni de collation
sensible à la langue. Les limites sont 256 caractères de requête, 500 formes correspondantes, 32 768 caractères indexés
par forme et 2 millions de caractères indexés au total. Le statut
identifie explicitement les résultats partiels. Les sous-arbres masqués et le texte de groupe supprimé sont
exclus ; les pages d'arrière-plan sont indexées une seule fois comme des pages distinctes. La mise en évidence
par occurrence et la recherche dans les données de forme ne sont pas implémentées.

## Afficher les calques {#display-layers}

Utilisez les commandes Calques partagées ou la méthode `setLayerVisibility(pageId, layerId, visible)`
de n'importe quel handle monté ou natif pour remplacer la visibilité d'affichage.
`visible: null` rétablit un indicateur enregistré ; `resetLayerVisibility(pageId?)` réinitialise
une page source ou le document entier. Les identifiants de page source distinguent les calques
de premier plan et d'arrière-plan. Les remplacements figés sont disponibles dans
`controller.state.layerVisibilityOverrides`.

Les changements de calques synchronisent le rendu, la sélection et la recherche sans modifier le
document. Les repères, NoShow et les formes masquées héritées décrites de façon incomplète ne peuvent pas être
révélés par un remplacement de calque. Le panneau accessible affiche au plus 200 calques uniques
et signale la troncature ; l'API autorise au plus 25 000 remplacements
à l'échelle du document. Les artefacts SVG et d'impression continuent d'utiliser la visibilité enregistrée. Les couleurs de
calque, l'édition et la politique d'impression enregistrée ne sont pas implémentées.

## Onglets de compléments du ruban {#ribbon-add-in-tabs}

Un hôte peut ajouter des onglets après Aide, comme le fait un complément dans Visio. `ribbonAddIns` est une
propriété de l'élément et une prop de chaque liaison de framework, et
`office-ribbon-add-in` fait partie de la table des événements :

```tsx
const tabs = [
	{
		id: 'reports',
		label: 'Reports',
		groups: [
			{
				label: 'Export',
				commands: [{ id: 'export', label: 'Export', icon: 'save', run: () => exportReport() }],
			},
		],
	},
];

<VisioViewer
	ribbonAddIns={tabs}
	events={{ 'office-ribbon-add-in': ({ tab, command }) => track(tab, command) }}
/>;
```

Vue prend `:ribbon-add-ins` et émet `office-ribbon-add-in` ; Angular prend
`[ribbonAddIns]` et dispose de la sortie `(ribbonAddIn)` ; la liaison vanilla prend
`mountViewer(host, { ribbonAddIns })` et `viewer.update({ ribbonAddIns })`. Sur
l'élément lui-même, définissez `element.ribbonAddIns` et écoutez l'événement DOM.

Une commande exécute son rappel `run` puis déclenche l'événement. Omettre la prop
conserve les onglets actuels ; un tableau vide les supprime. Repasser un descripteur
égal, comme le fait un framework à chaque rendu, conserve les panneaux et
ne remplace que les rappels. Le descripteur (`RibbonAddInTab`, issu de
`ooxml-core/ribbon`, également exporté par `ooxml-ui` et par chaque paquet de framework) est
partagé avec les éditeurs Word, Excel et PowerPoint. Un onglet ne peut pas prendre l'identifiant d'un
onglet intégré (`home`, `insert`, `design`, `data`, `process`, `review`, `view`,
`help`).

## Export SVG statique {#static-svg-export}

`exportPageSvg(model, pageIndex, { maxBytes })` et la méthode `exportSvg()` de chaque handle monté ou natif
renvoient un instantané de la page courante sans modifier le document, la
sélection ni la fenêtre d'affichage. L'export inclut les arrière-plans pris en charge, les ressources
raster intégrées et des notes de compatibilité visibles dans la description et les métadonnées du SVG. L'export SVG
n'a aucune ressource distante ni téléchargement automatique. La démo ne télécharge qu'après
l'appui explicite sur un bouton.

Le plafond UTF-8 est de 16 Mio ; les appelants peuvent l'abaisser. Des contrôles préalables prudents
peuvent rejeter un contenu dont la sérialisation finale serait plus petite. Les polices ne sont pas
intégrées et le texte et la mise en page restent approximatifs. Les dimensions physiques de la page sont
préservées ; les rastériseurs en aval doivent plafonner leurs dimensions de sortie et la surface en pixels
avant d'allouer une surface d'image. L'export PDF, l'export bitmap et la mise en page d'impression restent
non implémentés.

Des tests de pixels secondaires facultatifs sous Linux s'exécutent avec
`node scripts/test-svg-rasterization.mjs` après le build, en utilisant Python 3 du système,
librsvg/Cairo et la dépendance canvas native du cœur. Ces vérifications sont distinctes
de la validation dans le navigateur ou dans Microsoft Visio.

## Préparer des artefacts de pages de dessin {#prepare-drawing-page-artifacts}

`createPrintSnapshot(model, { pageIndices: [0, 2] })` préserve l'ordre explicite des pages
et renvoie des enregistrements SVG distincts, profondément figés, avec les dimensions du dessin et
tous les diagnostics d'export. La méthode `createPrintSnapshot()` de chaque handle monté ou natif
capture sa page courante. Elle rejette le remplacement du document pendant la préparation
sans modifier l'état de la vue. Il n'y a ni cadres, ni boîtes de dialogue, ni téléchargements, ni commandes
d'imprimante.

L'instantané utilise l'apparence d'affichage enregistrée. Les réglages d'impression de calque, NonPrinting et d'imprimante
enregistrés ne sont pas appliqués ; les dimensions du dessin ne correspondent ni au format de papier de l'imprimante ni à
l'échelle d'impression. Le résultat inclut ces limitations. Des garde-fous bornent les pages sélectionnées,
les octets agrégés, les arrière-plans répétés, les dimensions des images raster, le travail sur le texte et les tracés, et
les validations répétées du modèle entier. Les appelants peuvent abaisser les limites. Les SVG distincts doivent
rester isolés, car leurs identifiants de ressources internes peuvent se répéter. Il s'agit d'une
préparation à un futur flux d'impression, pas d'une prise en charge de l'impression native de Visio.

## Rechercher et remplacer avec portée {#scoped-find-and-replace}

Accueil > Rechercher > Remplacer et Ctrl+H ouvrent les commandes de remplacement partagées.
Suivant, Précédent, Remplacer et Tout remplacer utilisent des occurrences complètes, littérales et
sensibles à la casse. Les portées sont Sélection, Page courante et Toutes les pages. La sélection inclut les
sous-arbres de groupes sélectionnés dans l'ordre de sélection ou de la source, les cibles qui se recouvrent étant
dédupliquées. Les portées de page suivent l'ordre enregistré de la source, y compris le texte source masqué.
La navigation boucle et sélectionne une cible visible lorsque c'est possible. Respecter la casse est
coché et actuellement fixe ; les mots entiers, le pliage de casse et les codes de recherche spéciaux
ne sont pas implémentés. Ces règles de portée et de parcours sont une politique d'application
bornée, pas une prétention d'équivalence avec la boîte de dialogue Rechercher native.

Chaque liaison native expose le même `controller` : les hôtes peuvent utiliser
`captureTextReplaceToken(scope)`, `getTextReplaceOccurrences(token, query)`,
`selectTextReplaceOccurrence(token, occurrence, query)`,
`planTextReplacement(token, { query, replacement, mode, current? })` et
`applyTextReplacePlan(plan, token)`. La navigation et l'application renvoient un nouveau
jeton opaque. Conservez ce jeton pour garder la portée sélectionnée d'origine tout en passant
d'une correspondance à l'autre. Les changements manuels de sélection ou de page, le remplacement de la source ou une autre
opération invalident les jetons précédents. `cancelTextReplace(token)` n'annule que l'application
active de ce jeton, en laissant s'exécuter les modifications plus récentes sans rapport. Les plans et les jetons appartiennent à leur contrôleur
et ne peuvent pas être falsifiés ni réutilisés entre instances.

Le remplacement utilise une seule modification atomique de la source et une seule étape d'annulation, y compris
lors de la navigation entre résultats de plusieurs pages. Les événements natifs `document-change`, `selection-change` et `shape-select`
restent disponibles via les rappels du framework. Tout remplacer préserve la
sélection ; le remplacement courant passe à l'occurrence suivante. Un remplacement
par le même texte conserve exactement les octets sources et n'ajoute rien à l'historique tout en continuant de naviguer.
Un texte correspondant enrichi, contenant des champs, lié à un masque, protégé ou dépendant,
et non pris en charge, fait refuser le lot entier. Le brouillon affiche l'erreur et reste modifiable.
Fermer les commandes ou changer d'intention annule son opération en attente.

## Édition locale expérimentale de texte brut {#experimental-local-plain-text-editing}

Après `load(bytesOrBlob)`, sélectionnez une forme locale et utilisez les commandes de texte partagées,
ou appelez `replacePlainText(pageId, shapeId, text)` sur n'importe quel handle monté ou natif.
`undo()` et `redo()` utilisent l'historique borné du document ; `cancelEdit()` annule
le travail en attente. Les six adaptateurs transmettent les mêmes méthodes et le même événement
`document-change`, dont le détail est `{ document, dirty, kind }` (kind `edit`, `undo` ou `redo`, plus `remote` et `save` ; voir ci-dessus).
Inspectez `controller.state.edit` pour la disponibilité de la source, l'état occupé ou modifié,
la disponibilité de l'annulation/rétablissement, la troncature de l'historique, les erreurs et les diagnostics.

Les documents VSDX adossés à la source, chargés ou créés par le visualiseur, peuvent être modifiés ou
exportés au format VSDX.
Les imports VSD binaires v11 utilisent le codec ole2 publié via ooxml-core : les transformations
explicites de pages et de formes prises en charge, la géométrie de déplacement et de ligne et le texte brut sont affichés en aperçu
avec des styles de repli diagnostiqués. Les octets hérités d'origine sont conservés en privé ;
l'édition, l'annulation/rétablissement et l'export VSDX sont désactivés pour les sources héritées. Les versions
binaires antérieures, les groupes et masques, la géométrie complexe, les gabarits de formes, les modèles et les formats
avec macros restent non pris en charge. Affecter un modèle avec `document` ne fournit pas d'octets de paquet
modifiables. La validation du cœur rejette les formes liées à un masque, le texte enrichi, les
champs, les paquets signés et le contenu avec macros. Les modifications du XML et du paquet s'exécutent dans un
worker isolé ; les wrappers de framework n'implémentent aucune logique de format. Les modifications
de texte brut refusent les caches de formules dépendant du texte, touchés ou inconnus, au lieu de les laisser
périmés, et respectent la protection `LockTextEdit` effective. Le texte logique VSDX
exclut son marqueur de paragraphe final stocké tout en conservant les caractères LF finaux
intentionnels et les paragraphes vides. La sémantique du texte des VSD hérités
est inchangée. L'édition du contenu en texte enrichi reste non prise en charge.

`exportVsdx()` renvoie `{ bytes, dirty, diagnostics }` pour une copie téléchargée explicitement.
Il n'écrase pas le fichier source, ne l'envoie nulle part et ne revendique pas de fidélité d'aller-retour
native. L'interface partagée ne télécharge qu'après une action explicite de l'utilisateur. L'historique
est borné et peut écarter d'anciens états d'annulation, ce que signale `historyTruncated`.
Le remplacement de texte brut reste limité au sous-ensemble admis ci-dessus. Les opérations de
mise en forme et de géométrie ci-dessous ont leurs propres règles d'admission. Les preuves de réouverture
native couvrent des cas enregistrés précis, décrits dans le
[registre de vérification](/verification) ; elles n'établissent pas une fidélité générale.

## Mise en forme et disposition expérimentales {#experimental-formatting-and-arrangement}

Le ruban Accueil et `applyEdits()` partagent les commandes `format-text`, `format-shape` et
`reorder-shape`. La mise en forme du texte comprend les familles de polices enregistrées, la taille
en points, gras/italique/souligné/barré, la couleur, les puces, le retrait et
l'alignement horizontal/vertical. La mise en forme des formes comprend le remplissage uni, la couleur et l'épaisseur
du trait, les motifs de trait 0-23, les motifs de remplissage 0-24, la couleur d'arrière-plan et
la transparence. La mise en forme exige des formes feuilles locales ordinaires. La mise en forme du texte de toute la forme
met à jour chaque ligne de caractères ou de paragraphes effective tout en préservant
les marqueurs d'exécution, le texte, les bits de style sans rapport et les champs en cache pris en charge. La mise en forme des caractères
et celle des paragraphes ont une admission de source distincte ; l'alignement vertical
ne réécrit aucune des deux sections. Les lignes protégées ou ambiguës et les dépendances de formules affectées
font rejeter la transaction entière. L'augmentation et la diminution de la taille de police
exigent une taille uniforme ; le sélecteur de taille peut définir une sélection de tailles mixtes.
Les cellules locales Character/Paragraph `F="Inh"` sont modifiables lorsque leur ancêtre et
leur valeur ou unités en cache peuvent être prouvés. La délégation sans cache exige un véritable ancêtre ;
un choix de mise en forme explicite crée un remplacement local même si sa valeur est inchangée.
Les cellules héritées de protection et de transformation de premier niveau restent protégées.
Les polices doivent déjà exister dans le dessin source. Les commandes d'empilement prennent en charge
premier plan/arrière-plan/avancer/reculer pour une forme ordinaire de premier niveau dans la bande d'affichage zéro.

`format-shape` accepte `linePattern`, `fillPattern`, `fillBackgroundColor`
(`'#RRGGBB'`), `lineTransparency` et `fillTransparency` (pourcentage de 0 à 100).
La transparence est arrondie par pas d'un demi-pour-cent, les égalités vers le haut. La transparence de remplissage
définit ensemble le premier plan et l'arrière-plan. Le motif 0 signifie aucune peinture ; le motif 1 est
uni. `linePattern: 1` explicite active un contour indépendamment de la couleur du trait.
Un motif de remplissage explicite remplace le motif uni impliqué par `fillColor` ;
`fillColor: 'none'` avec un motif de remplissage non nul est rejeté. Une transparence de dégradé active
exige un remplacement explicite de la peinture, tandis que les peintures de thème unies héritées ordinaires sont
prises en charge.

Le menu Trait partagé expose Aucun trait et les motifs intégrés. Les options supplémentaires des deux menus de peinture
et le menu contextuel de la forme ouvrent le volet de tâches Format de la forme, qui
applique chaque changement immédiatement, comme une seule transaction de sélection. Les valeurs sources mixtes
s'affichent vides jusqu'à modification. `fillColorTheme`, `lineColorTheme` et (sur
`format-text`) `fontColorTheme` nomment l'échantillon de couleurs de thème d'où provient une couleur,
afin qu'elle soit enregistrée sous forme de formule de thème Visio ; `beginArrow`, `endArrow`,
`beginArrowSize`, `endArrowSize`, `lineCap` et `rounding` définissent les extrémités du trait,
l'extrémité et l'arrondi des coins. Les six liaisons natives utilisent
l'`applyEdits()` existant et les événements réactifs de document et de sélection ; aucun moteur de peinture propre à un framework n'est ajouté.
`visioShapeFormattingState(shapes)` fournit les valeurs de peinture sources agrégées pour
les commandes personnalisées. Ces valeurs sont distinctes des multiplicateurs d'opacité de dégradé ou de mosaïque.

Six commandes Aligner alignent sur la première forme sélectionnée. Position fournit un espacement égal
des bords horizontaux et verticaux pour trois formes ou plus. L'assistant du cœur
`visioArrangeCommands(page, selectedIds, action)` produit des modifications atomiques
pour les hôtes avec des commandes personnalisées. Il mesure les boîtes de largeur/hauteur pivotées, ordonne
la distribution selon les centres et préserve l'ordre de sélection en cas d'égalité. Les groupes, masques,
calques et connecteurs collés restent hors de sa portée d'admission. Ces actions
partagent l'historique, la conservation de la sélection et le téléchargement/la réouverture VSDX avec les autres modifications.

## Duplication expérimentale {#experimental-duplication}

`duplicateSelection(): Promise<void>` sur le contrôleur et sur chaque handle monté ou natif
duplique les formes sélectionnées admises de la page courante. Le menu Coller de l'onglet Accueil, le menu contextuel de la forme,
Dites-moi et Ctrl+D appellent la même méthode. Les champs de texte
conservent leur comportement clavier natif. Dupliquer n'utilise pas le presse-papiers système.

L'assistant du cœur `visioDuplicateCommand(page, selectedIds)` produit une commande source
`duplicate-shapes` avec de nouveaux identifiants locaux à la page. Les copies conservent le XML source,
le texte mixte et l'ordre d'empilement, reçoivent de nouveaux noms sans identifiants uniques copiés,
et se déplacent de 0,33 pouce de dessin vers la droite et vers le bas. La validation de la source
admet actuellement les formes feuilles 2D locales non collées, à l'exclusion des masques, groupes, contenus étrangers,
calques, identités ambiguës et dépendances non prises en charge. La protection des points d'ancrage et
les caches de formules affectés sont vérifiés avant d'accepter toute modification de la source.

Les copies obtenues deviennent la sélection en une seule étape d'historique atomique. Annuler
restaure la sélection d'origine et rétablir les copies, sauf si une sélection ou
une action de page plus récente la remplace. Les opérations échouées, annulées ou remplacées ne peuvent pas
modifier partiellement les octets sources ni la sélection. Les événements ordinaires `selection-change` et
`document-change` exposent le résultat accepté à l'état du framework natif.

## Presse-papiers de formes expérimental {#experimental-shape-clipboard}

Chaque handle monté ou natif expose `copySelection()`, `cutSelection()` et
`pasteSelection()`, chacun renvoyant `Promise<void>`. Appelez-les depuis un geste de l'utilisateur
afin que le navigateur puisse accorder l'accès au presse-papiers. L'onglet Accueil, les menus contextuels, Dites-moi
et Ctrl/Meta+C, X et V utilisent le même adaptateur partagé. Les champs de texte conservent le
comportement natif du presse-papiers. Les échecs du presse-papiers apparaissent dans le visualiseur sans modifier
les octets sources, l'historique ni la sélection. Couper écrit dans le presse-papiers avant de supprimer ;
si la protection empêche la suppression, les formes copiées restent dans le presse-papiers et
le visualiseur signale qu'elles n'ont pas pu être coupées.

La charge utile est du `text/plain` borné avec le marqueur `OOXML-VISIO-SHAPES/1`. Elle
contient le XML des formes capturées et le contexte source nécessaire, de sorte que modifier ou couper
les originaux n'invalide pas le contenu copié. Coller lit toujours le presse-papiers système
actuel ; il ne réutilise pas silencieusement une copie interne précédente. Les formats de presse-papiers natifs de
Microsoft Visio, le texte arbitraire, les images et le Collage spécial ne sont pas pris en charge. Les formes collées reçoivent de nouveaux identifiants et noms et un décalage
de 0,33 pouce de dessin vers la droite et vers le bas. Les collages répétés utilisent ce même décalage ;
la cascade de collage native et le placement entre pages ne sont pas établis.

La première portée prise en charge est celle des feuilles 2D locales ordinaires non collées, avec des définitions
de police, de style et de thème du document et des paramètres de page, un ordinal, un contexte d'arrière-plan
et un rapport de dessin identiques. Les références doivent rester à l'intérieur de la sélection copiée.
Les groupes, masques, calques, contenus étrangers, fragments porteurs de relations, formules non prises en charge
et l'importation de ressources restent refusés. Le cœur valide toute la charge utile avant une modification
atomique de la source. Coller sélectionne ses copies et partage les règles d'historique borné et d'intention
de sélection de Dupliquer.

`controller.state.clipboard` expose l'état immuable `ready`, `preparing` et `error`
via l'abonnement réactif existant de chaque framework. La capture de la source
s'exécute dans un worker annulable et est invalidée par les changements de sélection, de page, de visibilité
ou de document. Les hôtes fournissant leur propre transport de presse-papiers peuvent utiliser
`captureClipboardToken()`, `prepareClipboardSelection(token)`,
`getPreparedClipboard(token)`, `cutPreparedSelection(token)` et
`pasteClipboardText(text, token)` sur le contrôleur. Les jetons rejettent les actions périmées.
Le cœur sans DOM exporte `captureVisioClipboard`, `serializeVisioClipboard`,
`deserializeVisioClipboard`, la modification `paste-shapes` et le constructeur de commandes
`visioPasteCommand(page, clipboard, offset?)`.

## Édition expérimentale de la géométrie {#experimental-geometry-editing}

Les commandes de géométrie partagées exposent la création d'un rectangle, le déplacement, le redimensionnement et la suppression
sûre. Chaque handle de framework natif ou monté transmet aussi
`applyEdits(edits: readonly VisioEdit[])` pour un lot atomique via le même
worker, le même historique et le même chemin d'annulation.

Accueil > Outil Texte et Insertion > Zone de texte utilisent le même brouillon de zone de texte
de taille fixe. Faites glisser sur un espace vide de la page, saisissez du texte brut, puis utilisez Ajouter une zone de texte ou
Ctrl/Meta+Entrée ; Échap l'annule. Les brouillons vides ne créent pas de formes sources ;
les espaces et sauts de ligne intentionnels sont préservés. Cliquer sur une forme existante avec l'outil Texte la sélectionne et révèle
les commandes de modification de texte existantes. Le dessin de rectangles, d'ellipses et de lignes utilise le
même cycle de vie de création protégé. Leur transaction interne du contrôleur capture
la source, la page, la sélection ordonnée, le zoom et l'intention de visibilité avant de dessiner,
vérifie cette intention avant d'accepter la source, et sélectionne les formes créées dans la
même étape d'annulation. Une intention plus récente ou la libération remplace une création en attente. Les jetons
de création internes ne constituent pas une API de liaison native supplémentaire.

Les six handles publics acceptent la nouvelle commande ordinaire via `applyEdits()` :

```ts
await handle.applyEdits([
	{
		type: 'create-text-box',
		pageId: '0',
		shapeId: '10',
		x: 4,
		y: 5,
		width: 3,
		height: 2,
		text: 'Text with an intentional trailing blank paragraph\n',
	},
]);
```

Le texte est obligatoire. Le style de texte par défaut enregistré fournit la mise en forme initiale des
caractères. `FillPattern=0` et `LinePattern=0` explicites rendent la nouvelle zone
sans peinture ; changer seulement la couleur du trait n'active pas de contour. La création d'un brouillon
de taille fixe n'établit ni les valeurs par défaut de l'outil Texte de l'interface native, ni le dimensionnement
automatique du texte, ni l'édition de contenu en texte enrichi. `applyEdits()` générique conserve son comportement
de sélection existant ; l'outil de dessin partagé gère la sélection de la forme créée.
Les preuves natives couvrent 24 cas d'API de zones fixes enregistrées à trois échelles de dessin,
avec les styles par défaut et personnalisés Text Only/Arial 18 pt, les sorties du cœur étant rouvertes
dans Visio. Ces cas de style enregistré ne mesurent pas les valeurs par défaut interactives de l'interface ; consultez le
[registre de vérification](/verification).

Insertion > Nouvelle page passe par la commande d'insertion de page adossée à la source existante,
puis sélectionne la nouvelle page. Elle partage le comportement d'insertion de la barre de pages, l'annulation/rétablissement
et l'admission de la source ; cela n'ajoute pas de création de page à partir d'un modèle.

```ts
await handle.applyEdits([
	{
		type: 'create-rectangle',
		pageId: '0',
		shapeId: '9',
		x: 4,
		y: 5,
		width: 3,
		height: 2,
		text: 'New shape',
	},
	{ type: 'move-shape', pageId: '0', shapeId: '1', x: 6, y: 7 },
]);
```

Les coordonnées sont des positions du point de rotation en pouces de dessin, origine en bas à gauche,
positif vers le haut. Le redimensionnement maintient fixe le point d'ancrage sauf si la commande inclut
`anchor: { x, y }`. Chaque coordonnée d'ancrage vaut `0`, `0.5` ou `1`, et identifie le
point fixe dans les limites locales normalisées de largeur/hauteur de la forme, orientées vers le haut. Par
exemple, `{ x: 0, y: 0 }` maintient fixe son coin local inférieur gauche. Le redimensionnement
ancré a une preuve de source distincte pour les points d'ancrage locaux, les caches affectés et les valeurs fixes de
rotation et de symétrie. Les identifiants sont explicites. Les formes 2D de premier niveau locales
admises existantes peuvent être modifiées, y compris celles importées d'autres producteurs.
Le cœur évalue les dépendances ShapeSheet numériques affectées prises en charge et
préserve les charges ZIP non touchées. Le visualiseur ne modifie jamais le XML lui-même.

Il s'agit d'un sous-ensemble étroit et admis. Les masques, groupes et formes étrangères,
connecteurs et collage, protections ou redirections dangereuses, formules affectées non prises en charge,
dépendances de paquet ambiguës et suppressions référencées échouent avec une erreur
visible. La géométrie de ligne relative est mise à l'échelle ; la géométrie de ligne absolue exige une dépendance de
dimension prise en charge. Consultez le
[`src/core/visio/README.md`](https://github.com/ChristopherVR/ooxml/blob/main/src/core/visio/README.md)
du cœur pour le contrat des commandes, les limites numériques, les raisons d'exclusion et les prochaines extensions.
Le [registre des capacités](/parity) consigne l'ensemble admis actuel, qui inclut désormais
aussi les déplacements bornés d'instances de masques primitifs avec des points d'ancrage locaux explicites.

Les tests de formes simples générées ou importées et un flux de travail Chrome local réussissent. Les 19
diagrammes du corpus public acceptés n'admettent actuellement aucune modification de géométrie, car la portée des
dépendances hors page ne peut pas encore être prouvée indépendante. Treize n'ont de plus aucune forme locale
éligible. La couverture d'édition pratique de ce corpus reste bloquée ;
la suite consiste en un graphe de dépendances de paquet, de masque et de thème à portée définie et en l'édition
d'instances de masques, puis en un routage des extrémités collées. La réouverture native a depuis été
vérifiée pour les captures précises de mise en forme et de disposition consignées dans le
[registre de vérification](/verification) ; la fidélité générale reste non vérifiée.
