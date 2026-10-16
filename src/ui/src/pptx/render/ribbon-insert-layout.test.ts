import { describe, expect, it } from 'vitest';

import { RIBBON_CONTROL_IDS, RIBBON_GROUP_IDS } from './customization/ribbon-control-ids';
import { INSERT_GROUPS, insertControlStatus } from './ribbon-insert-layout';
import type { InsertControlSpec } from './ribbon-insert-layout';
import { INSERT_GROUP_CAPTIONS } from './ribbon-insert-state';
import type { RibbonInsertState } from './ribbon-insert-state';

const state: RibbonInsertState = {
	editable: true,
	hasSelection: true,
	shapeType: 'rect',
	chartKind: 'column',
};
const controls = INSERT_GROUPS.flatMap((group) => group.items.flat()) as InsertControlSpec[];
const find = (id: string) => controls.find((item) => item.control === id)!;

describe('Insert tab layout', () => {
	it('follows PowerPoint: the groups in order, every one captioned and catalogued', () => {
		expect(INSERT_GROUPS.map((group) => group.id)).toStrictEqual([
			'insert.slides',
			'insert.tables',
			'insert.images',
			'insert.camera',
			'insert.illustrations',
			'insert.forms',
			'insert.addIns',
			'insert.links',
			'insert.comments',
			'insert.text',
			'insert.symbols',
			'insert.media',
		]);
		for (const group of INSERT_GROUPS) {
			expect(INSERT_GROUP_CAPTIONS[group.id], group.id).toBeDefined();
			expect(RIBBON_GROUP_IDS as readonly string[], group.id).toContain(group.id);
		}
	});

	it('gives every control a catalogue id and a ScreenTip description', () => {
		expect(new Set(controls.map((item) => item.control)).size).toBe(controls.length);
		for (const item of controls) {
			expect(RIBBON_CONTROL_IDS as readonly string[], item.control).toContain(item.control);
			expect(item.tip[1].length, item.control).toBeGreaterThan(10);
		}
	});

	it('orders the controls inside each group as PowerPoint does', () => {
		const ids = (group: string) =>
			INSERT_GROUPS.find((item) => item.id === group)!
				.items.flat()
				.map((item) => (item as InsertControlSpec).control);
		expect(ids('insert.illustrations').map((id) => id.split('.')[2])).toStrictEqual([
			'shapes',
			'icons',
			'threeDModels',
			'smartArt',
			'chart',
		]);
		expect(ids('insert.text').map((id) => id.split('.')[2])).toStrictEqual([
			'textBox',
			'headerFooter',
			'wordArt',
			'dateTime',
			'slideNumber',
			'object',
		]);
		expect(ids('insert.media').map((id) => id.split('.')[2])).toStrictEqual([
			'media',
			'audio',
			'screenRecording',
		]);
	});

	it('marks controls without a backing feature unsupported rather than hiding them', () => {
		for (const id of ['insert.camera.cameo', 'insert.forms.forms', 'insert.images.screenshot']) {
			expect(insertControlStatus(state, find(id)), id).toStrictEqual({
				hidden: false,
				disabled: true,
				reason: 'unsupported',
			});
		}
		expect(insertControlStatus(state, find('insert.slides.newSlide')).reason).toBe('unsupported');
		expect(
			insertControlStatus({ ...state, newSlideAvailable: true }, find('insert.slides.newSlide')),
		).toStrictEqual({ hidden: false, disabled: false, reason: null });
		expect(
			insertControlStatus({ ...state, hasSelection: false }, find('insert.links.link')),
		).toStrictEqual({
			hidden: false,
			disabled: true,
			reason: 'selection',
		});
		expect(
			insertControlStatus({ ...state, editable: false }, find('insert.tables.table')).reason,
		).toBe('readOnly');
	});
});
