// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { parseVsdx } from 'ooxml-core/visio';
import { VISIO_TEMPLATES, demoDocument } from 'ooxml-core/visio/ui';
import { mountViewer } from './binding';

afterEach(() => {
	vi.restoreAllMocks();
	document.body.replaceChildren();
});
function setup() {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: demoDocument });
	vi.spyOn(viewer.controller, 'parser').mockImplementation(parseVsdx);
	return { viewer, root: viewer.element.shadowRoot! };
}

it('shows a tile per template, Blank Drawing first', () => {
	const { viewer, root } = setup();
	viewer.element.openBackstage('new');
	const tiles = [...root.querySelectorAll<HTMLButtonElement>('.backstage-templates button')];
	expect(tiles.map((tile) => tile.querySelector('strong')!.textContent)).toEqual(
		VISIO_TEMPLATES.map((template) => template.name),
	);
	expect(tiles[0]!.dataset.backstageAction).toBe('new-blank');
	viewer.destroy();
});

it.each([
	['basic-flowchart', 4],
	['cross-functional-flowchart', 3],
	['cross-functional-flowchart-vertical', 3],
	['organization-chart', 5],
] as const)('creates the %s template from its tile', async (id, links) => {
	const { viewer, root } = setup();
	viewer.element.openBackstage('new');
	root.querySelector<HTMLButtonElement>(`[data-backstage-action="new-template:${id}"]`)!.click();
	await vi.waitFor(() => expect(viewer.element.fileName).toBe('New drawing.vsdx'));
	const page = viewer.controller.state.document!.pages[0]!;
	expect(page.connectors).toHaveLength(links * 2);
	expect(viewer.exportVsdx().dirty).toBe(false);
	viewer.destroy();
});
