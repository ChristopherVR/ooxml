/** Reading the Theme Colors swatches of the shared `pptx-ui-color-grid` through its open shadow root. */

/** The Recent Colors rows are the same element, marked with the row's test id. */
const THEME_GRID = 'pptx-ui-color-grid:not([data-testid="pptx-color-recent"])';

/** Every Theme Colors grid under `scope` (or `scope` itself, when it is one), in DOM order. */
export function themeGrids(scope: ParentNode): HTMLElement[] {
	const own = scope instanceof HTMLElement && scope.matches(THEME_GRID);
	return [...(own ? [scope] : []), ...Array.from(scope.querySelectorAll<HTMLElement>(THEME_GRID))];
}

/** Every swatch button of those grids, in DOM order. */
export function themeSwatches(scope: ParentNode): HTMLButtonElement[] {
	return themeGrids(scope).flatMap((grid) =>
		Array.from(grid.shadowRoot?.querySelectorAll<HTMLButtonElement>('button.swatch') ?? []),
	);
}

/** The swatches of a Recent Colors row (`row` itself or the first one under it). */
export function recentSwatches(row: ParentNode): HTMLButtonElement[] {
	const grid =
		row instanceof HTMLElement && row.dataset.testid === 'pptx-color-recent'
			? row
			: row.querySelector<HTMLElement>('[data-testid="pptx-color-recent"]');
	return Array.from(grid?.shadowRoot?.querySelectorAll<HTMLButtonElement>('button') ?? []);
}

/** The first swatch with the given label (its tooltip and accessible name), or null. */
export function themeSwatch(scope: ParentNode, label: string): HTMLButtonElement | null {
	return themeSwatches(scope).find((button) => button.title === label) ?? null;
}
