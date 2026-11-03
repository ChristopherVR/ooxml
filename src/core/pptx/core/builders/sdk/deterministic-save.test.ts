/**
 * `save(slides, { fixedDate })` gives the same bytes for the same input,
 * whenever it runs, and the dates and revision it writes load back. Without
 * `fixedDate` a save is dated and counted as before.
 */
import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxHandlerSaveOptions } from '../../core/types';
import { createShapeElement } from './ElementFactory';

const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const FIXED = new Date('2026-01-02T03:04:05Z');

async function build(options?: PptxHandlerSaveOptions): Promise<Uint8Array> {
	const { handler, createSlide } = await PptxHandler.create({ title: 'Report' });
	const first = createSlide('Title and Content')
		.addText('Hello', { fontSize: 24 })
		.addShape('rect', { fill: { type: 'solid', color: '#336699' } })
		.addImage(PNG)
		.addTable({ rows: [{ cells: [{ text: 'a' }, { text: 'b' }, { text: 'c' }] }] })
		.addChart('bar', { series: [{ name: 'S', values: [1, 2] }], categories: ['x', 'y'] })
		.addConnector({})
		.addGroup([createShapeElement('ellipse')]);
	const second = createSlide('Blank').addFreeform('M0 0 L10 10 L0 10 Z');
	return handler.save([first.build(), second.build()], options);
}

async function coreXml(bytes: Uint8Array): Promise<string> {
	const zip = await JSZip.loadAsync(bytes);
	return (await zip.file('docProps/core.xml')?.async('string')) ?? '';
}

afterEach(() => {
	vi.useRealTimers();
});

describe('fixedDate', () => {
	it('produces identical bytes for saves made at different times', async () => {
		vi.useFakeTimers({ toFake: ['Date'] });
		vi.setSystemTime(new Date('2026-05-05T10:00:00Z'));
		const first = await build({ fixedDate: FIXED });
		vi.setSystemTime(new Date('2026-06-07T11:00:03.500Z'));
		const second = await build({ fixedDate: FIXED });
		expect(Buffer.from(second).equals(Buffer.from(first))).toBe(true);
	});

	it('dates the package, takes revision 1 and reads both back', async () => {
		const bytes = await build({ fixedDate: FIXED });
		const core = await coreXml(bytes);
		expect(core).toContain('<cp:revision>1</cp:revision>');
		expect(core).toMatch(/<dcterms:created [^>]*>2026-01-02T03:04:05Z<\/dcterms:created>/u);
		expect(core).toMatch(/<dcterms:modified [^>]*>2026-01-02T03:04:05Z<\/dcterms:modified>/u);
		const zip = await JSZip.loadAsync(bytes);
		// ZIP stores DOS time to 2 s, written in UTC and read back as local time.
		for (const file of Object.values(zip.files)) {
			expect(Math.abs(file.date.getTime() - FIXED.getTime())).toBeLessThanOrEqual(
				14 * 3600_000 + 2000,
			);
		}
		const loaded = await new PptxHandler().load(bytes);
		expect(loaded.coreProperties?.revision).toBe('1');
		expect(loaded.coreProperties?.created).toBe('2026-01-02T03:04:05Z');
		expect(loaded.coreProperties?.modified).toBe('2026-01-02T03:04:05Z');
	});

	it('lets coreProperties win over the fixed date and revision', async () => {
		const core = await coreXml(
			await build({
				fixedDate: FIXED,
				coreProperties: { revision: '7', created: '2020-02-02T00:00:00Z' },
			}),
		);
		expect(core).toContain('<cp:revision>7</cp:revision>');
		expect(core).toMatch(/<dcterms:created [^>]*>2020-02-02T00:00:00Z</u);
		expect(core).toMatch(/<dcterms:modified [^>]*>2026-01-02T03:04:05Z</u);
	});

	it('names a new group after its shape id, not its time-based element id', async () => {
		const zip = await JSZip.loadAsync(await build({ fixedDate: FIXED }));
		const slide = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
		const name = /<p:grpSp><p:nvGrpSpPr><p:cNvPr id="(\d+)" name="([^"]*)"/u.exec(slide);
		expect(name?.[2]).toBe(`Group ${name?.[1]}`);
	});

	it('dates and counts as before without it, also on the same handler', async () => {
		const { handler, createSlide } = await PptxHandler.create();
		const slides = [
			createSlide('Blank')
				.addGroup([createShapeElement('rect')])
				.build(),
		];
		await handler.save(slides, { fixedDate: FIXED });
		const bytes = await handler.save(slides);
		const core = await coreXml(bytes);
		// The first save created the package on the fixed date; the second one
		// modifies it now and counts the revision up from there.
		expect(core).toMatch(/<dcterms:created [^>]*>2026-01-02T03:04:05Z</u);
		expect(core).not.toMatch(/<dcterms:modified [^>]*>2026-01-02T03:04:05Z</u);
		expect(core).toContain('<cp:revision>2</cp:revision>');
		const zip = await JSZip.loadAsync(await build());
		const slide = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
		expect(slide).toMatch(/<p:grpSp><p:nvGrpSpPr><p:cNvPr id="\d+" name="grp_/u);
	});

	it('rejects an invalid date', async () => {
		await expect(build({ fixedDate: new Date('nope') })).rejects.toThrow(RangeError);
	});
});
