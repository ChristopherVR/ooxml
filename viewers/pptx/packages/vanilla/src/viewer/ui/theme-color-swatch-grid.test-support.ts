/** Reading the Theme Colors swatches of the shared `pptx-ui-color-grid` through its open shadow root. */

/** Every Theme Colors grid under `scope` (or `scope` itself, when it is one), in DOM order. */
export function themeGrids(scope: ParentNode): HTMLElement[] {
	const own = scope instanceof HTMLElement && scope.localName === 'pptx-ui-color-grid';
	return [
		...(own ? [scope] : []),
		...Array.from(scope.querySelectorAll<HTMLElement>('pptx-ui-color-grid')),
	];
}

/** Every swatch button of those grids, in DOM order. */
export function themeSwatches(scope: ParentNode): HTMLButtonElement[] {
	return themeGrids(scope).flatMap((grid) =>
		Array.from(grid.shadowRoot?.querySelectorAll<HTMLButtonElement>('button.swatch') ?? []),
	);
}

/** The first swatch with the given label (its tooltip and accessible name), or null. */
export function themeSwatch(scope: ParentNode, label: string): HTMLButtonElement | null {
	return themeSwatches(scope).find((button) => button.title === label) ?? null;
}
