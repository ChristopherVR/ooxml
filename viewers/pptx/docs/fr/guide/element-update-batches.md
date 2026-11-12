---
title: Mises à jour d'éléments par lot
description: Mettre à jour des éléments sur plusieurs diapositives comme une seule opération annulable, via l'API publique de la visionneuse.
---

# Mises à jour d'éléments par lot

Utilisez `updateElements` lorsqu'une action de l'hôte doit mettre à jour des éléments sur plusieurs diapositives. Par exemple, un bouton « Ajuster les titres du rapport » peut repositionner les titres de deux pages sans que l'utilisateur quitte la page qu'il est en train de modifier.

La méthode est disponible sur le handle public de la visionneuse dans React, Vue, Angular, Svelte et Vanilla JS. `ElementUpdate` et `ElementUpdateOptions` sont exportés depuis le package de composant de chaque binding.

```ts
updateElements(
	updates: readonly ElementUpdate[],
	options?: ElementUpdateOptions,
): Promise<void>;

interface ElementUpdate {
	slideId: string;
	elementId: string;
	patch: Partial<PptxElement>;
}

interface ElementUpdateOptions {
	label?: string;
}
```

Ici `viewer` est le handle de composant public actuel, et les deux diapositives contiennent déjà un élément titre :

```ts
const [first, second] = viewer.getSlides();

await viewer.updateElements(
	[
		{
			slideId: first.id,
			elementId: first.elements[0].id,
			patch: { x: 84 },
		},
		{
			slideId: second.id,
			elementId: second.elements[0].id,
			patch: { x: 90 },
		},
	],
	{ label: 'Adjust report titles' },
);

viewer.undo(); // Restore both titles in one step.
viewer.redo(); // Reapply both updates in one step.
```

## Comportement de validation et d'historique {#commit-and-history-behavior}

- Chaque lot effectif crée une seule étape d'annulation, indépendante des modifications précédentes et suivantes, y compris un autre lot soumis dans le même tour JavaScript.
- Appliquer un lot conserve la diapositive active et la sélection d'éléments. Il ne navigue pas vers les diapositives cibles. Annuler et rétablir utilisent le mécanisme existant de restauration par instantané de la visionneuse.
- Attendez la fin de l'opération avant de lire le résultat ou d'en lancer une opération dépendante. Dans React, récupérez de nouveau `ref.current` lorsque vous lisez via le handle public.
- Le texte en ligne en attente est validé séparément avant chaque lot effectif. Soumettez les lots entre deux interactions de pointeur terminées, et non pendant un glisser-déposer ou un redimensionnement : les cinq bindings rejettent un lot tant qu'une interaction de pointeur est active. Un lot rejeté laisse intacts le texte en attente, l'état du document et l'historique d'annulation.
- Les lots vides et les lots sans modification nette ne changent ni l'historique ni l'état modifié. Les lots rejetés préservent la pile du document et les piles annuler/rétablir.
- La méthode `updateElement` existante conserve sa sémantique actuelle, individuelle et limitée à la diapositive active.

## Cibles, correctifs et erreurs {#targets-patches-and-errors}

Les lots exigent un document chargé et modifiable, en mode d'édition ordinaire des diapositives. Les contextes en lecture seule, d'aperçu, de présentation, de modification des masques et de modèles rejettent la promesse retournée. Les identifiants cibles doivent se résoudre de façon unique : `slideId` identifie une diapositive, et `elementId` identifie un élément de premier niveau de cette diapositive. Les enfants de groupes imbriqués et les éléments de masque ou de disposition sortent du périmètre de cette méthode.

Chaque cible et chaque correctif sont vérifiés avant qu'une modification ne soit validée. Si une cible est absente ou ambiguë, le lot entier est rejeté. Les correctifs doivent être des objets et ne peuvent pas remplacer l'`id` ou le `type` d'un élément. Les champs de position, de taille et de rotation doivent être des nombres finis ; la largeur et la hauteur ne peuvent pas être négatives. Les autres champs conservent le contrat existant `Partial<PptxElement>` ; il ne s'agit pas d'un validateur complet du schéma des éléments.

Les correctifs utilisent un remplacement superficiel, comme `updateElement` : fournissez une valeur complète de propriété imbriquée lorsque vous la remplacez. Plusieurs entrées visant le même élément s'appliquent dans l'ordre de saisie. Les entrées sont copiées défensivement : modifier un objet de correctif après l'appel de la méthode ne change pas le lot soumis.
