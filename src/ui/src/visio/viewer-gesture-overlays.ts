/** Restore only captured nodes; a replacement render owns its newly created overlays. */
export function hideGestureOverlays(viewport: HTMLElement, selector: string): () => void {
	const entries = Array.from(viewport.querySelectorAll<HTMLElement | SVGElement>(selector)).map(
		(element) => {
			const visibility = element.style.visibility;
			element.style.visibility = 'hidden';
			return { element, visibility };
		},
	);
	return () => {
		for (const { element, visibility } of entries) element.style.visibility = visibility;
	};
}
