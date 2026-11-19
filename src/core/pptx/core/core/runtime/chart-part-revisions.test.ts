import { describe, expect, it } from 'vitest';

import type { PptxChartData } from '../../types';
import { cloneSlide } from '../../utils/clone-utils';
import { ChartPartRevisions } from './chart-part-revisions';

const PATH = 'ppt/charts/chart1.xml';
const data = (): PptxChartData => ({ chartType: 'bar', categories: ['A'], series: [] });
/** Chart data copied as the editor's history copies it. */
const historyCopy = (chart: PptxChartData): PptxChartData =>
	(
		cloneSlide({
			id: 's',
			rId: 'r',
			slideNumber: 1,
			elements: [{ type: 'chart', id: 'c', x: 0, y: 0, width: 1, height: 1, chartData: chart }],
		}).elements[0] as { chartData: PptxChartData }
	).chartData;

describe('the saved states of a chart part', () => {
	it('saves data from the part as it is now until the part is saved', () => {
		const revisions = new ChartPartRevisions();
		const loaded = data();
		revisions.start(PATH, loaded);
		expect(revisions.source(PATH, loaded)).toBeUndefined();
		expect(revisions.source(PATH, data())).toBeUndefined();
	});

	it('gives history data edited before a save the part as it was', () => {
		const revisions = new ChartPartRevisions();
		const live = data();
		revisions.start(PATH, live);
		const before = historyCopy(live);
		revisions.record(PATH, '<loaded/>', live);
		expect(revisions.source(PATH, live)).toBeUndefined();
		expect(revisions.source(PATH, before)).toBe('<loaded/>');
		const middle = historyCopy(live);
		revisions.record(PATH, '<first/>', live);
		expect(revisions.source(PATH, middle)).toBe('<first/>');
		expect(revisions.source(PATH, before)).toBe('<loaded/>');
	});

	it('reaches the state through a reactive proxy of the data', () => {
		// As Vue's `reactive` does: every extensible object read through the proxy is wrapped too.
		const reactive = <T>(value: T): T =>
			value && typeof value === 'object' && Object.isExtensible(value)
				? (new Proxy(value, { get: (target, key) => reactive(Reflect.get(target, key)) }) as T)
				: value;
		const revisions = new ChartPartRevisions();
		const live = data();
		revisions.start(PATH, live);
		const before = historyCopy(live);
		revisions.record(PATH, '<loaded/>', live);
		expect(revisions.source(PATH, reactive(before))).toBe('<loaded/>');
	});

	it('serializes a state as its id only', () => {
		const revisions = new ChartPartRevisions();
		const live = data();
		revisions.start(PATH, live);
		revisions.record(PATH, '<loaded/>', live);
		expect(JSON.parse(JSON.stringify(live.savedRevision))).toEqual({ id: live.savedRevision!.id });
	});

	it('finds the state for a copy that kept only its id while other data holds it', () => {
		const revisions = new ChartPartRevisions(0);
		const live = data();
		revisions.start(PATH, live);
		const held = historyCopy(live);
		const cloned = structuredClone(live);
		revisions.record(PATH, '<loaded/>', live);
		expect(revisions.source(PATH, structuredClone(live))).toBeUndefined();
		expect(
			revisions.source(PATH, JSON.parse(JSON.stringify(live)) as PptxChartData),
		).toBeUndefined();
		// The budget of 0 released the state: the copy and the data holding it save from the part
		// as it is now.
		expect(revisions.source(PATH, cloned)).toBeNull();
		expect(revisions.source(PATH, held)).toBeNull();
		const kept = new ChartPartRevisions();
		const other = data();
		kept.start(PATH, other);
		const holder = historyCopy(other);
		const copy = JSON.parse(JSON.stringify(other)) as PptxChartData;
		kept.record(PATH, '<loaded/>', other);
		expect(kept.source(PATH, copy)).toBe('<loaded/>');
		expect(holder.savedRevision).toBeDefined();
	});

	it('does not apply a state of one part to another', () => {
		const revisions = new ChartPartRevisions();
		const first = data();
		const second = data();
		revisions.start(PATH, first);
		revisions.start('ppt/charts/chart2.xml', second);
		const old = historyCopy(first);
		revisions.record(PATH, '<one/>', first);
		revisions.record('ppt/charts/chart2.xml', '<two/>', second);
		expect(revisions.source('ppt/charts/chart2.xml', old)).toBeNull();
	});

	it('releases the oldest states past the budget, compressed', () => {
		let seed = 7;
		const part = (save: number) =>
			`<c:chartSpace s="${save}">${Array.from({ length: 2000 }, (_, i) => {
				seed = (seed * 48271) % 2147483647;
				return `<c:pt idx="${i}"><c:v>${seed}</c:v></c:pt>`;
			}).join('')}</c:chartSpace>`;
		const parts = Array.from({ length: 6 }, (_, save) => part(save));
		const revisions = new ChartPartRevisions(40_000);
		const live = data();
		revisions.start(PATH, live);
		const held: PptxChartData[] = [];
		for (const xml of parts) {
			held.push(historyCopy(live));
			revisions.record(PATH, xml, live);
		}
		expect(revisions.size).toBeGreaterThan(0);
		expect(revisions.size).toBeLessThanOrEqual(40_000);
		expect(revisions.source(PATH, held[0]!)).toBeNull();
		expect(revisions.source(PATH, held[5]!)).toBe(parts[5]);
		expect(parts.reduce((total, xml) => total + xml.length, 0)).toBeGreaterThan(2 * 40_000);
	});

	it('shares the current state with a copied part', () => {
		const revisions = new ChartPartRevisions();
		const live = data();
		revisions.start(PATH, live);
		revisions.record(PATH, '<loaded/>', live);
		const copy = historyCopy(live);
		revisions.copy(PATH, 'ppt/charts/chart9.xml');
		expect(revisions.source('ppt/charts/chart9.xml', copy)).toBeUndefined();
		const before = historyCopy(copy);
		revisions.record('ppt/charts/chart9.xml', '<copied/>', copy);
		expect(revisions.source('ppt/charts/chart9.xml', before)).toBe('<copied/>');
		expect(revisions.source(PATH, live)).toBeUndefined();
	});
});
