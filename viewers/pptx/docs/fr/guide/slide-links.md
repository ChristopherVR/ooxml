# Liens vers les diapositives

Les liens hypertexte d'un passage de texte ou d'une forme vers une diapositive naviguent à l'intérieur de la visionneuse en mode lecture seule et en mode aperçu, ainsi que pendant un diaporama. Les liens internes `ppaction://` n'ouvrent jamais un nouvel onglet du navigateur. En édition, utilisez Ctrl+clic (Cmd+clic sur macOS) pour suivre un lien interne vers une diapositive ; un clic ordinaire continue de sélectionner ou de modifier l'élément.

Les visionneuses en lecture seule et en aperçu masquent l'indicateur d'action orange et son info-bulle de création. Les visionneuses en édition continuent de les afficher.

## Fonction de rappel de l'hôte {#host-callback}

Chaque liaison propose une fonction de rappel avant de suivre un lien de passage de texte ou de forme. Renvoyez `false` pour annuler la navigation par défaut, y compris la navigation externe.

| Liaison            | Fonction de rappel                            |
| ------------------ | --------------------------------------------- |
| Options Vanilla    | `onHyperlinkClick`                            |
| Prop React         | `onHyperlinkClick`                            |
| Prop de rappel Vue | `onHyperlinkClick` (ou `:on-hyperlink-click`) |
| Entrée Angular     | `[onHyperlinkClick]`                          |
| Prop Svelte        | `onhyperlinkclick`                            |

La fonction reçoit un `HyperlinkClickEvent` issu de `ooxml-ui/pptx` : l'`url` externe, l'`action` interne, un `targetSlideIndex` à partir de zéro lorsqu'il est résolu, et l'`elementId` de la forme lorsqu'il est disponible. Les champs sans valeur sont undefined.

```ts
const viewer = createPptxViewer(container, {
	onHyperlinkClick: (link) => {
		if (link.targetSlideIndex !== undefined) {
			selectSlideInHost(link.targetSlideIndex);
			return false;
		}
	},
});
```

Sans fonction de rappel, la visionneuse suit les liens normalement. Les liens externes conservent le cadre cible de la présentation, les contrôles de sécurité des URL et le paramètre de confirmation du Centre de gestion de la confidentialité.
