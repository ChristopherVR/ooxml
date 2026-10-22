/**
 * The Theme Colors swatches with this label in every `pptx-ui-color-grid` under `root`, read from
 * the grids' open shadow roots. Colour field tests use it to pick a theme colour by name.
 */
export function themeSwatches(root: ParentNode, title: string): HTMLButtonElement[] {
	return Array.from(root.querySelectorAll('pptx-ui-color-grid')).flatMap((grid) =>
		Array.from(
			grid.shadowRoot?.querySelectorAll<HTMLButtonElement>(`button[title="${title}"]`) ?? [],
		),
	);
}

/** The first Theme Colors swatch with this label under `root`, or null. */
export function themeSwatch(root: ParentNode, title: string): HTMLButtonElement | null {
	return themeSwatches(root, title)[0] ?? null;
}
