// @vitest-environment jsdom
import { beforeAll, describe, expect, it } from 'vitest';
import { parseVsdx, type VisioDocument } from 'ooxml-core/visio';
import { createTemplateVsdx } from 'ooxml-core/visio/ui';
import { createRibbon } from './ribbon';
import { renderContextualTabs } from './viewer-contextual-tabs';
import { swimlaneResizeEdits, swimlaneSelection } from './viewer-swimlanes';
import type { ViewerState } from './controller';

let document_: VisioDocument;
beforeAll(async () => {
	document_ = await parseVsdx(await createTemplateVsdx('cross-functional-flowchart'));
});
const stateWith = (ids: string[]) =>
	({
		document: document_,
		pageIndex: 0,
		selectedShapes: ids.map((id) => ({ id, pageId: document_.pages[0]!.id })),
	}) as never as ViewerState;

describe('Cross-Functional Flowchart contextual tab', () => {
	it('is a hidden contextual tab with Insert, Arrange and Design groups', () => {
		const ribbon = createRibbon(globalThis.document);
		globalThis.document.body.append(ribbon);
		const panel = ribbon.querySelector<HTMLElement>('[data-ribbon-tab="cross-functional"]')!;
		expect(panel.dataset.label).toBe('Cross-Functional Flowchart');
		expect(panel.dataset.contextual).toBeDefined();
		expect(panel.hasAttribute('data-tab-hidden')).toBe(true);
		const groups = [...panel.querySelectorAll('office-ui-ribbon-group')].map((g) =>
			g.getAttribute('label'),
		);
		expect(groups).toEqual(['Insert', 'Arrange', 'Design']);
		const labels = [...panel.querySelectorAll('[label]')].map((el) => el.getAttribute('label'));
		for (const name of [
			'Swimlane',
			'Separator',
			'Direction',
			'Rotate Lane Labels',
			'Style',
			'Insert Swimlane Before',
			'Delete Swimlane',
		])
			expect(labels).toContain(name);
		// Direction is listed, disabled, with the reason it cannot change an existing flowchart.
		const direction = panel.querySelector<HTMLElement>('[data-menu="swimlane-direction"]')!;
		expect(direction.hasAttribute('disabled')).toBe(true);
		expect(direction.title).toContain('not available yet');
		ribbon.remove();
	});

	it('shows for the frame, a lane or a shape in a lane, and not for other shapes', () => {
		const ribbon = createRibbon(globalThis.document);
		globalThis.document.body.append(ribbon);
		const setTabHidden = (tab: string, _reason: string, hidden: boolean) =>
			ribbon
				.querySelector(`[data-ribbon-tab="${tab}"]`)!
				.toggleAttribute('data-tab-hidden', hidden);
		Object.assign(ribbon, { setTabHidden });
		const panel = ribbon.querySelector<HTMLElement>('[data-ribbon-tab="cross-functional"]')!;
		for (const id of ['50', '10', '2']) {
			renderContextualTabs(ribbon, stateWith([id]));
			expect(panel.hasAttribute('data-tab-hidden'), id).toBe(false);
		}
		renderContextualTabs(ribbon, stateWith([]));
		expect(panel.hasAttribute('data-tab-hidden')).toBe(true);
		renderContextualTabs(ribbon, stateWith(['100']));
		expect(panel.hasAttribute('data-tab-hidden')).toBe(true);
		ribbon.remove();
	});

	it('finds the lane a selection belongs to', () => {
		expect(swimlaneSelection(stateWith(['11']))).toMatchObject({ lane: '11', isLane: true });
		expect(swimlaneSelection(stateWith(['2']))).toMatchObject({ lane: '11', isLane: false });
		expect(swimlaneSelection(stateWith(['50']))).toMatchObject({ isLane: false });
		expect(swimlaneSelection(stateWith(['100']))).toBeUndefined();
	});

	it('turns a lane resize into a stacking edit and refuses a frame resize', () => {
		const page = document_.pages[0]!;
		const edits = swimlaneResizeEdits(page, ['11'], { width: 8.4, height: 3 });
		expect(edits?.[0]).toMatchObject({ type: 'resize-swimlane', laneId: '11', size: 3 });
		expect(edits?.some((edit) => edit.type === 'move-shape')).toBe(true);
		expect(swimlaneResizeEdits(page, ['50'], { width: 9, height: 6 })).toBeUndefined();
		expect(swimlaneResizeEdits(page, ['2'], { width: 2, height: 1 })).toBeNull();
		expect(swimlaneResizeEdits(page, ['100'], { width: 2, height: 1 })).toBeNull();
	});
});
