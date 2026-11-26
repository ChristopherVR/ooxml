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
import { PptxHandler } from '../../core/PptxHandler';
import type { PptxSlide } from '../../core/types';
import { OPERATIONS, type Deck } from './part-ownership-operations';
import { chartPackageIssues } from '../chart-save-sweep';

const FIXTURES = path.resolve(__dirname, '../fixtures');
const THREE_D = 'e2e/three-d-parity/three-d-charts.pptx';
const OLE = 'e2e/ole-editable.pptx';
const SMARTART = 'corpus/smartart-chart-table-mix.pptx';
const DECKS = {
	gallery: ['e2e/chart-gallery.pptx'],
	charts: ['e2e/chart-data-fidelity.pptx'],
	'three-d-elements': [THREE_D],
	'three-d-slides': [THREE_D],
	smartart: [SMARTART, 'smartart-gallery/continuous-cycle--hier5.pptx'],
	'smartart-layouts': [SMARTART],
	ole: ['corpus/ole-embedded-media.pptx'],
	'ole-elements': [OLE],
	'ole-slides': [OLE],
	'ole-masters': [OLE],
	'ole-layouts': [OLE],
	media: [
		'e2e/media-bookmark-trigger-editable.pptx',
		'e2e/audio-embed.pptx',
		'embedded-assets-sample.pptx',
	],
} as const;
type DeckGroup = keyof typeof DECKS;
const OUT = process.env['B1_SWEEP_OUT'];

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

const issueKeys = async (zip: JSZip) =>
	new Set((await chartPackageIssues(zip)).map((issue) => `${issue.kind} ${issue.detail}`));

async function baselineIssues(bytes: Uint8Array): Promise<ReadonlySet<string>> {
	const { handler, data } = await load(bytes);
	try {
		return issueKeys(await JSZip.loadAsync(await handler.save(data.slides)));
	} finally {
		handler.dispose();
	}
}

type OperationName = keyof typeof OPERATIONS;
const elements = [
	'duplicate element (sdk)',
	'duplicate element (editor)',
	'shallow element copy',
	'paste on the next slide',
	'paste on a new blank slide',
	'move to another slide, delete the source',
] satisfies OperationName[];
const slides = [
	'duplicate slide (sdk)',
	'duplicate slide then its elements',
] satisfies OperationName[];
const masters = ['duplicate master', 'insert master'] satisfies OperationName[];
const layouts = ['duplicate every layout'] satisfies OperationName[];

function operationGroups(deck: string): Partial<Record<DeckGroup, readonly OperationName[]>> {
	switch (deck) {
		case THREE_D:
			return {
				'three-d-elements': elements,
				'three-d-slides': [...slides, ...masters, ...layouts],
			};
		case SMARTART:
			return { smartart: [...elements, ...slides, ...masters], 'smartart-layouts': layouts };
		case OLE:
			return {
				'ole-elements': elements,
				'ole-slides': slides,
				'ole-masters': masters,
				'ole-layouts': layouts,
			};
		default:
			return {};
	}
}

const originals = new Map<
	string,
	{
		bytes: Uint8Array;
		baseline?: Promise<ReadonlySet<string>>;
	}
>();

function originalDeck(deckName: string) {
	let original = originals.get(deckName);
	if (!original) {
		original = { bytes: new Uint8Array(readFileSync(path.join(FIXTURES, deckName))) };
		originals.set(deckName, original);
	}
	return original;
}

export function describePartOwnershipSweep(group: DeckGroup): void {
	describe('part ownership sweep', () => {
		for (const deckName of DECKS[group]) {
			const source = originalDeck(deckName);
			const selected: readonly string[] | undefined = operationGroups(deckName)[group];
			describe(deckName, () => {
				for (const [operationName, operation] of Object.entries(OPERATIONS)) {
					if (selected && !selected.includes(operationName)) continue;
					it(
						operationName,
						async () => {
							const baseIssues = await (source.baseline ??= baselineIssues(source.bytes));
							const deck = await load(source.bytes);
							let current = deck;
							try {
								current = (await operation(deck)) ?? deck;
								const { handler, data } = current;
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
									const issues = [...(await issueKeys(zip))].filter((key) => !baseIssues.has(key));
									expect(issues, label).toStrictEqual([]);
								}
								// Each chart frame shows the data its element holds, a pasted copy its source's.
								const reloaded = await load(saved);
								try {
									expect(chartSeries(reloaded.data.slides)).toStrictEqual(chartSeries(data.slides));
								} finally {
									reloaded.handler.dispose();
								}
								if (OUT) {
									mkdirSync(OUT, { recursive: true });
									const name = `${deckName.replace(/[\/]/gu, '_')}__${operationName.replace(/[^a-z]+/giu, '-')}.pptx`;
									writeFileSync(path.join(OUT, name), saved);
								}
							} finally {
								if (current !== deck) current.handler.dispose();
								deck.handler.dispose();
							}
						},
						120_000,
					);
				}
			});
		}
	});
}

export function checkOwnershipOperations(): void {
	it('assigns all copy operations once for each split deck', () => {
		for (const deck of [THREE_D, SMARTART, OLE]) {
			expect(Object.values(operationGroups(deck)).flat().sort()).toEqual(
				Object.keys(OPERATIONS).sort(),
			);
		}
	});
}
