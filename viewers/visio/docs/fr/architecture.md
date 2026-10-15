# Architecture {#architecture}

## Responsabilités {#ownership}

- `ooxml-core/visio` dans le dépôt frère `ooxml` : accueil des paquets, XML, modèle, héritage, géométrie en cache, diagnostics structurés. Aucune interface de visualiseur n'y a sa place.
- `src/controller.ts` : état du document et de la vue, sélection, événements, comportement « le dernier chargement l'emporte » et orchestration transactionnelle de l'édition.
- `src/document-history.ts` : historique d'édition borné adossé à la source et état d'export de copie VSDX. La mutation du XML et du paquet reste dans le cœur et s'exécute dans un worker dédié.
- `src/document-text-search.ts` : correspondance littérale bornée sur le texte visible normalisé des formes, avec des résultats immuables limités à chaque page. Le contrôleur partagé gère la navigation et les contrôles de réentrance.
- `src/render-svg.ts` et `src/render-text.ts` : présentation SVG et avertissements de rendu.
- `src/export-svg.ts` : sérialisation statique et bornée de la page courante par le même moteur de rendu SVG, avec ressources raster intégrées et métadonnées de compatibilité.
- `src/print-snapshot.ts` : artefacts de pages sélectionnées sans effet de bord, budgets agrégés et diagnostics préservés. Il ne crée pas de cadres d'impression et n'implémente aucune politique d'imprimante.
- `src/viewer-element.ts` : surface navigateur partagée.
- `src/contract.ts` et `src/binding.ts` : propriétés, événements, cycle de vie mount/update/load/fit/exportSvg/destroy côté client uniquement. Une mise à jour réentrante plus récente remplace le reste du correctif plus ancien.
- Wrappers de frameworks : cycle de vie du framework et transmission des événements et propriétés uniquement. Aucun fork d'analyseur, de rendu ou de géométrie par framework.

## Limites {#boundaries}

Le moteur utilise actuellement les valeurs ShapeSheet en cache. Il ne s'agit pas d'une évaluation complète des formules. Les métriques de police du navigateur, le retour à la ligne, l'espacement et l'indentation mixtes des paragraphes, le RTL et les puces sont implémentés avec un travail borné ; cela n'établit pas une fidélité exacte du texte Visio.

Le XML/HTML du document n'est pas injecté dans le DOM. Les nœuds SVG sont créés via les API du DOM. Aucun envoi de document, aucune télémétrie ni aucune récupération de document distant ne sont proposés. Les vérifications du paquet et des relations relèvent du cœur. Le composant navigateur analyse dans un worker dédié, avec annulation et un délai parent de 15 secondes. Des budgets indépendants pour la scène, les métadonnées, le travail sur le texte et les images raster décodées protègent le rendu. Les ressources d'image sont partagées au sein de chaque rendu et révoquées lors du remplacement, de la déconnexion ou de la libération. Les écouteurs de chargement de police déclenchent un recalcul de la mise en forme du texte et sont retirés à la déconnexion. Ces protections ne constituent pas une certification de sécurité.

Le remplacement expérimental de texte brut adossé à la source prend en charge l'annulation/rétablissement bornés et l'export explicite de copie VSDX. Le cœur refuse les formes liées à un masque, le texte enrichi, les champs, les paquets signés et le contenu de macros. Les documents composés d'un modèle seul sont en lecture seule. Les caches de formules ne sont pas recalculés et la vérification de la réouverture dans Visio natif reste en attente ; ce n'est pas une garantie d'aller-retour d'édition et d'enregistrement. Les méthodes des handles, les événements et les limites figurent dans l'[API du visualiseur](/fr/api).

## Sécurité et limites {#safety-and-limits}

Le cœur valide les relations OPC, rejette les DTD et entités ainsi que les chemins ZIP suspects, et borne l'entrée de l'archive, le nombre d'entrées, l'expansion déclarée, la sortie réellement diffusée en flux et la complexité du XML. Le visualiseur vérifie la taille du fichier avant d'allouer un Blob et valide les scènes fournies de l'extérieur avant de dessiner, y compris la structure réelle des images raster, le type MIME détecté et les dimensions intrinsèques. La validation des images raster vérifie la structure et les sommes de contrôle, pas le décodage entropique complet ; la robustesse des décodeurs natifs reste donc pertinente. Les relations externes ne sont jamais récupérées, les chaînes du document sont des nœuds de texte et les fonctionnalités non prises en charge sont signalées.

L'élément personnalisé du navigateur utilise un worker d'analyse dédié, avec annulation et une limite de 15 secondes côté parent. Le contrôleur sans interface et les environnements sans prise en charge de Worker utilisent les limites coopératives du cœur. Le rendu SVG et du texte s'exécute toujours sur le thread principal, avec des limites distinctes pour la scène, les images raster, les métadonnées et le travail sur le texte. Le projet n'est pas encore certifié sur le plan de la sécurité ni durci pour la production.

## Métafichiers intégrés {#embedded-metafiles}

Les métafichiers améliorés (EMF) intégrés font l'objet de diagnostics de compatibilité bornés au niveau des enregistrements. Le vérificateur rejette les formats, styles et états de mappage non pris en charge, ainsi que l'amplification de ressources, avant toute conversion. Il n'exécute ni les données WMF intégrées, ni les polices, images, scripts ou références externes. Le worker d'analyse jetable ne convertit qu'un sous-ensemble étroit admis de lignes, rectangles, ellipses et objets stock, via le paquet de conversion publié, avec des budgets par document et une échéance stricte gérée par le parent. L'analyse directe du cœur reste limitée à l'inspection, sauf si un convertisseur de confiance est explicitement fourni dans un hôte isolé. Les cinq médias EMF réels inspectés précédemment restent tous non pris en charge ; l'admission structurelle seule ne promet ni rendu ni fidélité à Visio.

Le cœur sans interface contient un assainisseur de vecteurs neutres et un validateur de transport testés. Le worker d'analyse utilise le convertisseur publié pour un sous-ensemble borné de primitives EMF, partagé par le rendu en direct, l'export SVG et le rendu d'impression immuable. Le repli d'analyse hors worker reste sans convertisseur. Les tests de découpe générés à partir de chemins seuls concordent avec le backend SVG Skia local ; le backend librsvg installé ignore les intersections imbriquées de clip-path. Les comparaisons dans le navigateur et avec Visio natif restent nécessaires. Consultez [l'examen détaillé de l'adoption](research/emf-adoption-review.md).

## Vérification {#verification}

Les tests unitaires et de fixtures du cœur établissent les cas d'analyse et de modèle qu'ils affirment. Les tests du visualiseur établissent le cycle de vie et le comportement de la vue. Les tests dans le navigateur établissent l'interaction réelle. Des fixtures authentiques créées dans Visio, des rendus de référence connus et des comparaisons mesurées restent nécessaires pour toute affirmation de fidélité. Les tests d'exécution des frameworks sont distincts de la compilation des wrappers.

## Site web {#website}

Build Vite statique multipage : accueil, guide, registre des capacités, architecture et un terrain de jeu séparé. Les liens sont relatifs et les ressources locales. Aucune police ni aucun script tiers. Aucun bundle d'analyseur sur la page d'accueil.

La conception de la documentation s'inspire de la [référence pptx-viewer](https://github.com/ChristopherVR/pptx-viewer). L'examen approfondi se trouve dans `research/pptx-pages-design-review.md`.

## Publication {#publication}

Sept paquets npm exposent l'API Visio sans interface et le visualiseur partagé avec six adaptateurs de frameworks natifs. Ils dépendent de `ooxml-core/visio` publié. Des versions horaires fondées sur les commits conventionnels construisent et vérifient les artefacts empaquetés avant la publication de confiance npm. GitHub Pages déploie la documentation et la démo après les vérifications.
