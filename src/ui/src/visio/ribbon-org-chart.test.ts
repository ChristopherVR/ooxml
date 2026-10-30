// @vitest-environment jsdom
import { beforeAll, describe, expect, it } from 'vitest';
import { parseVsdx, type VisioDocument } from 'ooxml-core/visio';
import { createTemplateVsdx } from 'ooxml-core/visio/ui';
import { createRibbon } from './ribbon';
import { renderContextualTabs } from './viewer-contextual-tabs';
import { orgChartSelection } from './viewer-org-chart';
import type { ViewerState } from './controller';

let document_: VisioDocument;
beforeAll(async () => {
	document_ = await parseVsdx(await createTemplateVsdx('organization-chart'));
});
const stateWith = (ids: string[]) =>
	({
		document: document_,
		pageIndex: 0,
		selectedShapes: ids.map((id) => ({ id, pageId: document_.pages[0]!.id })),
	}) as never as ViewerState;

describe('Org Chart contextual tab', () => {
	it('is a hidden contextual tab with Insert, Position and Arrange groups', () => {
		const ribbon = createRibbon(globalThis.document);
		globalThis.document.body.append(ribbon);
		const panel = ribbon.querySelector<HTMLElement>('[data-ribbon-tab="org-chart"]')!;
		expect(panel.dataset.label).toBe('Org Chart');
		expect(panel.dataset.contextual).toBeDefined();
		expect(panel.hasAttribute('data-tab-hidden')).toBe(true);
		expect(
			[...panel.querySelectorAll('office-ui-ribbon-group')].map((g) => g.getAttribute('label')),
		).toEqual(['Insert', 'Position', 'Arrange']);
		const labels = [...panel.querySelectorAll('[label]')].map((el) => el.getAttribute('label'));
		for (const name of ['Subordinate', 'Coworker', 'Assistant', 'Delete Position', 'Layout'])
			expect(labels).toContain(name);
		ribbon.remove();
	});

	it('shows beside the format tab for a position, and not for a connector or nothing', () => {
		const ribbon = createRibbon(globalThis.document);
		globalThis.document.body.append(ribbon);
		const setTabHidden = (tab: string, _reason: string, hidden: boolean) =>
			ribbon
				.querySelector(`[data-ribbon-tab="${tab}"]`)!
				.toggleAttribute('data-tab-hidden', hidden);
		Object.assign(ribbon, { setTabHidden });
		const panel = ribbon.querySelector<HTMLElement>('[data-ribbon-tab="org-chart"]')!;
		const format = ribbon.querySelector<HTMLElement>('[data-ribbon-tab="shape-format"]')!;
		renderContextualTabs(ribbon, stateWith(['2']));
		expect(panel.hasAttribute('data-tab-hidden')).toBe(false);
		expect(format.hasAttribute('data-tab-hidden')).toBe(false);
		renderContextualTabs(ribbon, stateWith(['100']));
		expect(panel.hasAttribute('data-tab-hidden')).toBe(true);
		renderContextualTabs(ribbon, stateWith([]));
		expect(panel.hasAttribute('data-tab-hidden')).toBe(true);
		ribbon.remove();
	});

	it('finds the position a selection is', () => {
		expect(orgChartSelection(stateWith(['5']))).toEqual({ pageId: '0', id: '5' });
		expect(orgChartSelection(stateWith(['100']))).toBeUndefined();
	});
});
