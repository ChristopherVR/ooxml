// @vitest-environment jsdom
import { beforeAll, describe, expect, it } from 'vitest';
import { parseVsdx, type VisioDocument } from 'ooxml-core/visio';
import { createTemplateVsdx } from 'ooxml-core/visio/ui';
import { createRibbon } from './ribbon';
import { renderContextualTabs } from './viewer-contextual-tabs';
import { orgChartSelection } from './viewer-org-chart';
import { planDeleteSelection } from './viewer-delete';
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
	it('is a hidden contextual tab with Insert, Position, Import and Arrange groups', () => {
		const ribbon = createRibbon(globalThis.document);
		globalThis.document.body.append(ribbon);
		const panel = ribbon.querySelector<HTMLElement>('[data-ribbon-tab="org-chart"]')!;
		expect(panel.dataset.label).toBe('Org Chart');
		expect(panel.dataset.contextual).toBeDefined();
		expect(panel.hasAttribute('data-tab-hidden')).toBe(true);
		expect(
			[...panel.querySelectorAll('office-ui-ribbon-group')].map((g) => g.getAttribute('label')),
		).toEqual(['Insert', 'Position', 'Import', 'Arrange']);
		const labels = [...panel.querySelectorAll('[label]')].map((el) => el.getAttribute('label'));
		for (const name of [
			'Subordinate',
			'Coworker',
			'Assistant',
			'Delete Position',
			'Layout',
			'Import Org Chart',
		])
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
		renderContextualTabs(ribbon, stateWith(['17']));
		expect(panel.hasAttribute('data-tab-hidden')).toBe(false);
		expect(format.hasAttribute('data-tab-hidden')).toBe(false);
		renderContextualTabs(ribbon, stateWith(['97']));
		expect(panel.hasAttribute('data-tab-hidden')).toBe(true);
		renderContextualTabs(ribbon, stateWith([]));
		expect(panel.hasAttribute('data-tab-hidden')).toBe(true);
		ribbon.remove();
	});

	it('finds the position a selection is', () => {
		expect(orgChartSelection(stateWith(['65']))).toEqual({ pageId: '0', id: '65' });
		expect(orgChartSelection(stateWith(['97']))).toBeUndefined();
	});
	it('turns Delete on a lone position into Delete Position, and leaves other selections alone', () => {
		const plan = planDeleteSelection(stateWith(['17']));
		expect(plan?.message).toBe('Deleted the position and its connectors.');
		// The two subordinates are joined to the manager in the same batch.
		expect(plan?.edits.filter((edit) => edit.type === 'drop-stencil-master')).toHaveLength(2);
		expect(plan?.edits.filter((edit) => edit.type === 'delete-shape').length).toBeGreaterThan(1);
		// A connector, or several shapes, are deleted as shapes.
		expect(planDeleteSelection(stateWith(['97']))?.edits).toEqual([
			{ type: 'delete-shape', pageId: '0', shapeId: '97' },
		]);
		expect(
			planDeleteSelection(stateWith(['17', '33']))?.edits.every(
				(edit) => edit.type === 'delete-shape',
			),
		).toBe(true);
	});
});
