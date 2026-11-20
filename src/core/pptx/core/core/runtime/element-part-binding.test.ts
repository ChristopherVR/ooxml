import { readFileSync } from 'node:fs';
import path from 'node:path';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { findSharedOwnedParts } from '../../../../opc/part-ownership';
import {
	parseRelationships,
	relationshipsPartFor,
	resolvePartPath,
} from '../../../../opc/relationships';
import { duplicateElement } from '../../builders/sdk/slide-operations';
import { PptxHandler } from '../../PptxHandler';
import type { ChartPptxElement, PptxElement, PptxSlide, SmartArtPptxElement } from '../../types';
import { cloneSlide } from '../../utils/clone-utils';

const FIXTURES = path.resolve(__dirname, '../../../__tests__/fixtures');
const CHARTS = new Uint8Array(readFileSync(path.join(FIXTURES, 'e2e/chart-data-fidelity.pptx')));
const SMART_ART = new Uint8Array(
	readFileSync(path.join(FIXTURES, 'corpus/smartart-chart-table-mix.pptx')),
);

async function load(bytes: Uint8Array) {
	const handler = new PptxHandler();
	const data = await handler.load(bytes.slice().buffer as ArrayBuffer);
	return { handler, data };
}

const firstOf = <T extends PptxElement['type']>(slides: readonly PptxSlide[], type: T) => {
	for (const slide of slides) {
		const element = slide.elements.find((candidate) => candidate.type === type);
		if (element) return { slide, element: element as Extract<PptxElement, { type: T }> };
	}
	throw new Error(`no ${type}`);
};

/** The part a slide's relationship `id` names. */
async function target(zip: JSZip, slide: string, id: string): Promise<string> {
	const rels = parseRelationships(await zip.file(relationshipsPartFor(slide))!.async('string'));
	return resolvePartPath(slide, rels.get(id)!.target);
}

const chartId = (element: ChartPptxElement) => element.chartData!.chartRelationshipId!;

describe('element part binding', () => {
	it('gives a duplicated chart its own chart part and workbook, unchanged until edited', async () => {
		const { handler, data } = await load(CHARTS);
		const { slide, element } = firstOf(data.slides, 'chart');
		const copy = duplicateElement(element) as ChartPptxElement;
		slide.elements.push(copy);
		const zip = await JSZip.loadAsync(await handler.save(data.slides));

		expect(await findSharedOwnedParts(zip)).toStrictEqual([]);
		const originalPart = await target(zip, slide.id, chartId(element));
		const copyPart = await target(zip, slide.id, chartId(copy));
		expect(copyPart).not.toBe(originalPart);
		expect(copy.chartData!.chartPartPath).toBe(copyPart);
		expect(element.chartData!.chartPartPath).toBe(originalPart);
		// Nothing changed in the copy, so its part is the original's bytes.
		expect(await zip.file(copyPart)!.async('string')).toBe(
			await zip.file(originalPart)!.async('string'),
		);

		copy.chartData!.series[0]!.values[0] = 31415926;
		const edited = await JSZip.loadAsync(await handler.save(data.slides));
		expect(await edited.file(copyPart)!.async('string')).toContain('31415926');
		expect(await edited.file(originalPart)!.async('string')).not.toContain('31415926');
		expect(await findSharedOwnedParts(edited)).toStrictEqual([]);
	});

	it('pastes a chart on another slide as a copy of the source chart, not of that slide own', async () => {
		const { handler, data } = await load(CHARTS);
		const [source, destination] = data.slides.filter((s) =>
			s.elements.some((e) => e.type === 'chart'),
		) as [PptxSlide, PptxSlide];
		const chart = source.elements.find((e) => e.type === 'chart') as ChartPptxElement;
		const pasted = { ...structuredClone(chart), id: 'pasted' } as ChartPptxElement;
		destination.elements.push(pasted);
		const zip = await JSZip.loadAsync(await handler.save(data.slides));

		expect(await findSharedOwnedParts(zip)).toStrictEqual([]);
		const pastedPart = await target(zip, destination.id, chartId(pasted));
		const sourcePart = await target(zip, source.id, chartId(chart));
		expect(pastedPart).not.toBe(sourcePart);
		expect(await zip.file(pastedPart)!.async('string')).toBe(
			await zip.file(sourcePart)!.async('string'),
		);
		expect(pasted.relationshipSourcePath).toBe(destination.id);
	});

	it('moves a chart off a slide that is then deleted, keeping its part', async () => {
		const { handler, data } = await load(CHARTS);
		const { slide, element } = firstOf(data.slides, 'chart');
		const part = element.chartData!.chartPartPath!;
		const destination = data.slides.find((s) => s !== slide)!;
		destination.elements.push({ ...structuredClone(element), id: 'moved' });
		data.slides = data.slides.filter((s) => s !== slide);
		const zip = await JSZip.loadAsync(await handler.save(data.slides));

		const moved = destination.elements.find((e) => e.id === 'moved') as ChartPptxElement;
		expect(await target(zip, destination.id, chartId(moved))).toBe(part);
		expect(await findSharedOwnedParts(zip)).toStrictEqual([]);
	});

	it('gives a duplicated diagram its own parts and points its data part at its own drawing', async () => {
		const { handler, data } = await load(SMART_ART);
		const { slide, element } = firstOf(data.slides, 'smartArt');
		const copy = duplicateElement(element) as SmartArtPptxElement;
		slide.elements.push(copy);
		const zip = await JSZip.loadAsync(await handler.save(data.slides));

		expect(await findSharedOwnedParts(zip)).toStrictEqual([]);
		const ids = (e: SmartArtPptxElement) => e.smartArtData!;
		for (const key of ['dataRelId', 'layoutRelId', 'styleRelId', 'colorsRelId'] as const)
			expect(await target(zip, slide.id, ids(copy)[key]!)).not.toBe(
				await target(zip, slide.id, ids(element)[key]!),
			);
		const dataXml = await zip
			.file(await target(zip, slide.id, ids(copy).dataRelId!))!
			.async('string');
		expect(dataXml).toContain(`relId="${ids(copy).drawingRelId}"`);
		expect(ids(copy).drawingRelId).not.toBe(ids(element).drawingRelId);
	});

	it('keeps copies apart when an undo snapshot taken before a save is saved again', async () => {
		const { handler, data } = await load(CHARTS);
		const { slide, element } = firstOf(data.slides, 'chart');
		slide.elements.push(duplicateElement(element));
		const snapshot = data.slides.map(cloneSlide);
		await handler.save(data.slides);
		const zip = await JSZip.loadAsync(await handler.save(snapshot));
		expect(await findSharedOwnedParts(zip)).toStrictEqual([]);
	});

	it('lets a peer that saves the same edited deck after the owner write its own consistent copy', async () => {
		const owner = await load(CHARTS);
		const peer = await load(CHARTS);
		const { slide, element } = firstOf(owner.data.slides, 'chart');
		const copy = duplicateElement(element) as ChartPptxElement;
		copy.chartData!.series[0]!.values[0] = 27182818;
		slide.elements.push(copy);
		// The room holds the edited model; each client saves it against its own package.
		const shared = owner.data.slides.map(cloneSlide);
		await owner.handler.save(owner.data.slides);
		const saved = await JSZip.loadAsync(await peer.handler.save(shared));
		expect(await findSharedOwnedParts(saved)).toStrictEqual([]);
		const peerSlide = shared.find((s) => s.id === slide.id)!;
		const peerCopy = peerSlide.elements.at(-1) as ChartPptxElement;
		const peerOriginal = peerSlide.elements.find((e) => e.id === element.id) as ChartPptxElement;
		const copyXml = await saved
			.file(await target(saved, slide.id, chartId(peerCopy)))!
			.async('string');
		const originalXml = await saved
			.file(await target(saved, slide.id, chartId(peerOriginal)))!
			.async('string');
		expect(copyXml).toContain('27182818');
		expect(originalXml).not.toContain('27182818');
	});
});
