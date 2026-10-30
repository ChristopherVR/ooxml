// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import type { VisioShape } from 'ooxml-core/visio';
import { createRibbon } from './ribbon';
import { contextualTabFor, renderContextualTabs, selectionKind } from './viewer-contextual-tabs';
import { routeRibbonAction, type RibbonTargets } from './ribbon-router';
import type { ViewerState } from './controller';

const shape = (id: string, extra: Partial<VisioShape> = {}) =>
	({ id, name: id, kind: 'shape', children: [], ...extra }) as unknown as VisioShape;

/** A state with one page holding `shapes`, of which `selected` ids are selected. */
function stateOf(shapes: VisioShape[], selected: string[]): ViewerState {
	const selection = selected.map((id) => ({ id, name: id, pageId: 'p1' }));
	return {
		document: { pages: [{ id: 'p1', shapes }] },
		pageIndex: 0,
		selectedShapes: selection,
		selectedShape: selection[0] ?? null,
	} as never as ViewerState;
}

const picture = shape('1', { kind: 'foreign', image: { bytes: new Uint8Array() } as never });
const container = shape('2', { structure: { type: 'container', memberIds: ['4'] } });
const chart = shape('3', { embeddedObject: { bytes: new Uint8Array() } as never });
const plain = shape('4');
const all = [picture, container, chart, plain];

function ribbonWithTabs() {
	const ribbon = createRibbon(document);
	document.body.append(ribbon);
	Object.assign(ribbon, {
		setTabHidden: (tab: string, _reason: string, hidden: boolean) =>
			ribbon
				.querySelector(`[data-ribbon-tab="${tab}"]`)!
				.toggleAttribute('data-tab-hidden', hidden),
	});
	return ribbon;
}
const visibleTabs = (ribbon: HTMLElement) =>
	[...ribbon.querySelectorAll<HTMLElement>('[data-contextual]')]
		.filter((panel) => !panel.hasAttribute('data-tab-hidden'))
		.map((panel) => panel.dataset.ribbonTab);

describe('contextual tab per selection kind', () => {
	it('classifies pictures, containers and charts', () => {
		expect(selectionKind(picture)).toBe('picture');
		expect(selectionKind(container)).toBe('container');
		expect(selectionKind(chart)).toBe('chart');
		expect(selectionKind(plain)).toBe('shape');
		expect(selectionKind(undefined)).toBe('shape');
	});

	it('shows exactly one tab: the specialised one for a lone object, else Shape Format', () => {
		const ribbon = ribbonWithTabs();
		const expected: [string[], string[]][] = [
			[[], []],
			[['1'], ['picture-format']],
			[['2'], ['container-format']],
			[['3'], ['chart-design']],
			[['4'], ['shape-format']],
			[['1', '4'], ['shape-format']],
		];
		for (const [selected, tabs] of expected) {
			renderContextualTabs(ribbon, stateOf(all, selected));
			expect(visibleTabs(ribbon), selected.join()).toEqual(tabs);
		}
		expect(contextualTabFor(stateOf(all, ['1']))).toBe('picture-format');
		ribbon.remove();
	});
});

describe('Picture Format tab', () => {
	it('lists Adjust commands disabled with a reason and keeps Arrange mirrored', () => {
		const ribbon = ribbonWithTabs();
		const panel = ribbon.querySelector<HTMLElement>('[data-ribbon-tab="picture-format"]')!;
		expect(
			[...panel.querySelectorAll('office-ui-ribbon-group')].map((g) => g.getAttribute('label')),
		).toEqual(['Adjust', 'Picture Styles', 'Arrange', 'Size']);
		for (const id of [
			'picture-brightness',
			'picture-contrast',
			'picture-auto-balance',
			'picture-compression',
		]) {
			const el = panel.querySelector<HTMLElement>(`[data-menu="${id}"]`)!;
			expect(el.hasAttribute('disabled'), id).toBe(true);
			expect(el.title).toContain('not available yet');
		}
		const crop = panel.querySelector<HTMLElement>('[command="picture-crop"]')!;
		expect(crop.hasAttribute('disabled')).toBe(true);
		const home = ribbon.querySelector<HTMLElement>('#home-panel [command="bring-forward"]')!;
		home.setAttribute('disabled', '');
		renderContextualTabs(ribbon, stateOf(all, ['1']));
		expect(panel.querySelector('[command="bring-forward"]')!.hasAttribute('disabled')).toBe(true);
		ribbon.remove();
	});
});

describe('Container Format tab', () => {
	it('has Size, Container Styles, Membership and Arrange', () => {
		const ribbon = ribbonWithTabs();
		const panel = ribbon.querySelector<HTMLElement>('[data-ribbon-tab="container-format"]')!;
		expect(
			[...panel.querySelectorAll('office-ui-ribbon-group')].map((g) => g.getAttribute('label')),
		).toEqual(['Size', 'Container Styles', 'Membership', 'Arrange']);
		for (const id of ['container-margins', 'container-resize', 'container-lock'])
			expect(panel.querySelector(`[data-menu="${id}"]`)!.hasAttribute('disabled'), id).toBe(true);
		ribbon.remove();
	});

	it('routes Fit, Select Contents and Disband to the container target', () => {
		const calls: string[] = [];
		const targets = {
			controller: { state: {} },
			container: (command: string) => calls.push(command),
		} as never as RibbonTargets;
		for (const command of ['fit', 'select-contents', 'disband'] as const)
			routeRibbonAction(targets, { type: 'container', command });
		expect(calls).toEqual(['fit', 'select-contents', 'disband']);
	});
});

describe('Chart Design tab', () => {
	it('exposes the chart editor entry points as insert actions', () => {
		const ribbon = ribbonWithTabs();
		const panel = ribbon.querySelector<HTMLElement>('[data-ribbon-tab="chart-design"]')!;
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		for (const id of ['chart-tab-edit', 'chart-tab-data', 'chart-tab-format'])
			panel
				.querySelector(`[command="${id}"]`)!
				.dispatchEvent(new CustomEvent('office-command', { bubbles: true }));
		expect(seen.map((detail) => (detail as { item?: string }).item ?? detail)).toEqual([
			'chart-edit',
			'chart-data',
			'chart-format',
		]);
		ribbon.remove();
	});
});
