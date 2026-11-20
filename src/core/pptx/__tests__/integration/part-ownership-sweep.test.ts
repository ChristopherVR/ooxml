/**
 * Save-time invariant: no copy path leaves two owners on one owned part.
 *
 * Every deck below is put through each way the editor, the SDK and a collaborating peer copy
 * content (duplicate an element, paste it on another slide, move it, duplicate a slide, a master
 * or a layout), saved, and the package is checked:
 *
 * - `findSharedOwnedParts` is empty: no chart, diagram, OLE object, embedded workbook, theme or
 *   layout is claimed by two owners (PowerPoint refuses such a file, 0x80070570);
 * - every relationship id a slide cites resolves on that slide;
 * - after a reload each chart shows the data its element held, so a pasted chart shows its
 *   source's series, not the destination slide's own chart;
 * - the charts' ranges, caches and workbooks agree as they did in the deck as loaded.
 *
 * Set `B1_SWEEP_OUT` to a folder to also write every saved deck there (PowerPoint COM checks).
 */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { findSharedOwnedParts } from '../../../opc/part-ownership';
import { parseRelationships, relationshipsPartFor } from '../../../opc/relationships';
import { NS, parseXml, type XmlElement } from '../../../xml';
import { duplicateElement, duplicateSlide } from '../../core/builders/sdk/slide-operations';
import {
	duplicateLayout,
	duplicateSlideMaster,
	insertSlideMaster,
} from '../../core/builders/sdk/master-layout-crud';
import { PptxHandler } from '../../core/PptxHandler';
import type { PptxData, PptxElement, PptxSlide } from '../../core/types';
import { duplicateElementOnSlide } from '../../editor/render/editor-mutations';
import { createBlankSlide } from '../../editor/render/slide-operations';
import { chartPackageIssues } from '../chart-save-sweep';

const FIXTURES = path.resolve(__dirname, '../fixtures');
const DECKS = [
	'e2e/chart-gallery.pptx',
	'e2e/chart-data-fidelity.pptx',
	'e2e/three-d-parity/three-d-charts.pptx',
	'corpus/smartart-chart-table-mix.pptx',
	'smartart-gallery/continuous-cycle--hier5.pptx',
	'corpus/ole-embedded-media.pptx',
	'e2e/ole-editable.pptx',
	'e2e/media-bookmark-trigger-editable.pptx',
	'e2e/audio-embed.pptx',
	'embedded-assets-sample.pptx',
];
const OUT = process.env['B1_SWEEP_OUT'];

interface Deck {
	handler: PptxHandler;
	data: PptxData;
}

const PART_BEARING = new Set<PptxElement['type']>([
	'chart',
	'smartArt',
	'ole',
	'media',
	'picture',
	'image',
	'group',
]);

/** Elements that show package parts: what every copy below copies. */
const partBearing = (slide: PptxSlide): PptxElement[] =>
	slide.elements.filter(
		(element) => PART_BEARING.has(element.type) && element.rawXml !== undefined,
	);

/**
 * What a clipboard round trip gives a binding to paste (`element-clipboard`): a new id, and no
 * native shape id, which the save allocates.
 */
function clipboardCopy(element: PptxElement, id: string): PptxElement {
	const copy = { ...(JSON.parse(JSON.stringify(element)) as PptxElement), id };
	delete copy.shapeId;
	return copy;
}

type Operation = (deck: Deck) => Promise<Deck | void>;

const OPERATIONS: Record<string, Operation> = {
	'duplicate element (sdk)': async ({ data }) => {
		for (const slide of data.slides)
			for (const element of partBearing(slide)) slide.elements.push(duplicateElement(element));
	},
	'duplicate element (editor)': async (deck) => {
		deck.data.slides.forEach((slide, index) => {
			for (const element of partBearing(slide)) {
				const next = duplicateElementOnSlide(deck.data.slides, index, element.id);
				if (next) deck.data.slides = next.slides;
			}
		});
	},
	'shallow element copy': async ({ data }) => {
		for (const slide of data.slides)
			for (const element of partBearing(slide))
				slide.elements.push({ ...element, id: `${element.id}-shallow` });
	},
	'paste on the next slide': async ({ data }) => {
		const copies = data.slides.map((slide) =>
			partBearing(slide).map((element) => clipboardCopy(element, `${element.id}-pasted`)),
		);
		data.slides.forEach((slide, index) => {
			const next = data.slides[(index + 1) % data.slides.length]!;
			next.elements.push(...copies[index]!);
			void slide;
		});
	},
	'paste on a new blank slide': async ({ data }) => {
		const blank = createBlankSlide(data.slides.length + 1);
		blank.layoutPath = data.slides[0]?.layoutPath;
		for (const slide of data.slides)
			for (const element of partBearing(slide))
				blank.elements.push(clipboardCopy(element, `${element.id}-blank`));
		data.slides.push(blank);
	},
	'move to another slide, delete the source': async ({ data }) => {
		const source = data.slides.find((slide) => partBearing(slide).length > 0);
		if (!source || data.slides.length < 2) return;
		const destination = data.slides.find((slide) => slide !== source)!;
		destination.elements.push(...partBearing(source).map((e) => clipboardCopy(e, `${e.id}-m`)));
		data.slides = data.slides.filter((slide) => slide !== source);
	},
	'duplicate slide (sdk)': async ({ data }) => {
		for (const slide of [...data.slides])
			if (partBearing(slide).length > 0)
				data.slides.push(duplicateSlide(slide, data.slides.length + 1));
	},
	'duplicate slide then its elements': async ({ data }) => {
		const source = data.slides.find((slide) => partBearing(slide).length > 0);
		if (!source) return;
		const copy = duplicateSlide(source, data.slides.length + 1);
		copy.elements.push(...partBearing(copy).map(duplicateElement));
		data.slides.push(copy);
	},
	'duplicate master': async ({ handler, data }) => {
		const result = await duplicateSlideMaster(handler, data, data.slideMasters![0]!.path);
		if (result.ok) return result;
	},
	'insert master': async ({ handler, data }) => {
		const result = await insertSlideMaster(handler, data);
		if (result.ok) return result;
	},
	'duplicate every layout': async (deck) => {
		let current = deck;
		for (const layout of [...(deck.data.slideMasters?.[0]?.layouts ?? [])]) {
			const result = await duplicateLayout(current.handler, current.data, layout.path);
			if (result.ok) current = result;
		}
		return current;
	},
};

async function load(bytes: Uint8Array): Promise<Deck> {
	const handler = new PptxHandler();
	const data = await handler.load(bytes.slice().buffer as ArrayBuffer);
	return { handler, data };
}

function relationshipCitations(root: XmlElement, ids: string[]): string[] {
	for (const attribute of Array.from(root.attributes))
		if (attribute.namespaceURI === NS.r && attribute.value) ids.push(attribute.value);
	for (const child of Array.from(root.children)) relationshipCitations(child, ids);
	return ids;
}

/** `slide#rId` for every relationship id a slide cites that its relationships do not declare. */
async function unresolvedCitations(zip: JSZip): Promise<string[]> {
	const missing: string[] = [];
	for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/[^/]+\.xml$/u.test(n))) {
		const declared = parseRelationships(
			await zip.file(relationshipsPartFor(name))?.async('string'),
		);
		const xml = await zip.file(name)!.async('string');
		for (const id of relationshipCitations(parseXml(xml).documentElement, []))
			if (!declared.has(id)) missing.push(`${name}#${id}`);
	}
	return missing;
}

/** Each chart's series values, slide by slide in paint order. */
function chartSeries(slides: readonly PptxSlide[]): string[][] {
	return slides.map((slide) =>
		slide.elements.flatMap((element) =>
			element.type === 'chart' && element.chartData
				? [JSON.stringify(element.chartData.series.map((series) => series.values))]
				: [],
		),
	);
}

const issueKeys = async (bytes: Uint8Array) =>
	new Set((await chartPackageIssues(bytes)).map((issue) => `${issue.kind} ${issue.detail}`));

describe('part ownership sweep', () => {
	for (const deckName of DECKS) {
		const bytes = new Uint8Array(readFileSync(path.join(FIXTURES, deckName)));
		describe(deckName, () => {
			for (const [operationName, operation] of Object.entries(OPERATIONS)) {
				it(
					operationName,
					async () => {
						const loaded = await load(bytes);
						const baseIssues = await issueKeys(await loaded.handler.save(loaded.data.slides));
						const deck = await load(bytes);
						const { handler, data } = (await operation(deck)) ?? deck;
						const saved = await handler.save(data.slides);
						// A second save of the same session must keep the copies apart too.
						const resaved = await handler.save(data.slides);
						for (const [label, output] of [
							['first save', saved],
							['second save', resaved],
						] as const) {
							const zip = await JSZip.loadAsync(output);
							expect(await findSharedOwnedParts(zip), label).toStrictEqual([]);
							expect(await unresolvedCitations(zip), label).toStrictEqual([]);
							const issues = [...(await issueKeys(output))].filter((key) => !baseIssues.has(key));
							expect(issues, label).toStrictEqual([]);
						}
						// Each chart frame shows the data its element holds, a pasted copy its source's.
						const reloaded = await load(saved);
						expect(chartSeries(reloaded.data.slides)).toStrictEqual(chartSeries(data.slides));
						if (OUT) {
							mkdirSync(OUT, { recursive: true });
							const name = `${deckName.replace(/[\/]/gu, '_')}__${operationName.replace(/[^a-z]+/giu, '-')}.pptx`;
							writeFileSync(path.join(OUT, name), saved);
						}
					},
					120_000,
				);
			}
		});
	}
});
