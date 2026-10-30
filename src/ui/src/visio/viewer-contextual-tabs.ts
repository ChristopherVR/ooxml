import type { ViewerState } from './controller';

const MIRRORED = ['disabled', 'main-disabled', 'checked', 'pressed', 'title'] as const;

/**
 * The state of the contextual Shape Format tab. The tab exists only while shapes are selected, and
 * its Shape Styles, Arrange and Size controls are the Home controls built again: every module that
 * syncs a control by its id updates the first match (Home), so this copies that state onto the
 * contextual twin instead of teaching each module about two copies.
 */
export function renderContextualTabs(toolbar: HTMLElement, state: ViewerState): void {
	const ribbon = toolbar as HTMLElement & { setTabHidden?(t: string, r: string, h: boolean): void };
	const hidden = state.selectedShapes.length === 0;
	ribbon.setTabHidden?.('shape-format', 'contextual', hidden);
	const panel = toolbar.querySelector<HTMLElement>('[data-ribbon-tab="shape-format"]');
	if (!panel || hidden) return;
	for (const twin of panel.querySelectorAll<HTMLElement>('[command], [data-menu]')) {
		const id = twin.getAttribute('command') ?? twin.dataset.menu!;
		const source = [
			...toolbar.querySelectorAll<HTMLElement>(`[command="${id}"], [data-menu="${id}"]`),
		].find((candidate) => !panel.contains(candidate) && candidate.localName === twin.localName);
		if (!source) continue;
		for (const name of MIRRORED) {
			const value = source.getAttribute(name);
			if (value === null) twin.removeAttribute(name);
			else if (twin.getAttribute(name) !== value) twin.setAttribute(name, value);
		}
		if (source.localName === 'office-ui-gallery')
			(twin as HTMLElement & { state?: unknown }).state = (
				source as HTMLElement & { state?: unknown }
			).state;
	}
}
