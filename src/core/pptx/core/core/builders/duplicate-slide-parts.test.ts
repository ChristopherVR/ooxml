import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { describe, expect, it } from 'vitest';
import { contentTypeForPart, parseContentTypes } from '../../../../opc/content-types';
import { ensureContentTypeOverride } from '../../../../opc/package';
import {
	buildRelationshipsXml,
	parseRelationships,
	relationshipsPartFor,
	resolvePartPath,
} from '../../../../opc/relationships';
import { duplicateSlide as duplicateToolSlide } from '../../../automation/tools/slide-tools';
import { duplicateSlide } from '../../builders/sdk/slide-operations';
import { Presentation } from '../../builders/sdk/Presentation';
import { SlideBuilder } from '../../builders/sdk/SlideBuilder';
import { PptxHandler } from '../../PptxHandler';
import type { ChartPptxElement, PptxSlide } from '../../types';
import { cloneSlide } from '../../utils/clone-utils';

const relationshipBase = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const chart = (slide: PptxSlide) =>
	slide.elements.find((element) => element.type === 'chart') as ChartPptxElement;
const buffer = (bytes: Uint8Array) =>
	bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

async function sourceDeck(withDependencies = false): Promise<Uint8Array> {
	const { handler, data, createSlide } = await PptxHandler.createBlank();
	const slide = createSlide('Blank')
		.setNotes('Original notes')
		.addText('Link')
		.addChart(
			'line',
			{
				categories: ['A', 'B', 'C'],
				series: [{ name: 'Series A', values: [0, 2, 3] }],
			},
			{ x: 60, y: 80, width: 700, height: 400 },
		);
	slide.addImage(
		'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=',
	);
	slide.addMedia('audio', 'data:audio/wav;base64,UklGRgQAAABXQVZF');
	data.slides.push(slide.build());
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const slideXml = await zip.file(data.slides[0].id)!.async('string');
	const withLink = slideXml.replace('</a:rPr>', '<a:hlinkClick r:id="sourceLink42"/></a:rPr>');
	expect(withLink).not.toBe(slideXml);
	zip.file(data.slides[0].id, withLink);
	const relsPath = relationshipsPartFor(data.slides[0].id);
	const rels = parseRelationships(await zip.file(relsPath)!.async('string'));
	rels.set('sourceLink42', {
		id: 'sourceLink42',
		type: `${relationshipBase}hyperlink`,
		target: 'https://example.com/?x=1&y=2',
		mode: 'External',
	});
	zip.file(relsPath, buildRelationshipsXml(rels));
	if (!withDependencies) return zip.generateAsync({ type: 'uint8array' });

	// Add optional chart-owned parts to the native chart fixture. Their bytes
	// are opaque to duplication, including the embedded workbook package.
	const workbook = new JSZip();
	workbook.file(
		'[Content_Types].xml',
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
	);
	zip.file('ppt/embeddings/workbook1.xlsx', await workbook.generateAsync({ type: 'uint8array' }));
	zip.file(
		'ppt/charts/colors1.xml',
		'<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="10" meth="cycle"><a:srgbClr val="FF0000"/></cs:colorStyle>',
	);
	zip.file(
		'ppt/charts/style1.xml',
		'<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="10"/>',
	);
	zip.file(
		'ppt/charts/_rels/chart1.xml.rels',
		buildRelationshipsXml(
			new Map([
				[
					'book',
					{
						type: `${relationshipBase}package`,
						target: '../embeddings/workbook1.xlsx',
						mode: 'Internal',
					},
				],
				[
					'colors',
					{
						type: 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle',
						target: 'colors1.xml',
						mode: 'Internal',
					},
				],
				[
					'style',
					{
						type: 'http://schemas.microsoft.com/office/2011/relationships/chartStyle',
						target: 'style1.xml',
						mode: 'Internal',
					},
				],
			]),
		),
	);
	for (const [path, type] of [
		[
			'ppt/embeddings/workbook1.xlsx',
			'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		],
		['ppt/charts/colors1.xml', 'application/vnd.ms-office.chartcolorstyle+xml'],
		['ppt/charts/style1.xml', 'application/vnd.ms-office.chartstyle+xml'],
	])
		await ensureContentTypeOverride(zip, path, type);
	return zip.generateAsync({ type: 'uint8array' });
}

async function relatedPart(zip: JSZip, part: string, type: string): Promise<string> {
	const relationships = parseRelationships(
		await zip.file(relationshipsPartFor(part))!.async('string'),
	);
	const relationship = [...relationships.values()].find((entry) =>
		entry.type.endsWith(`/${type}`),
	)!;
	expect(relationship, `${part}: ${type}`).toBeDefined();
	return resolvePartPath(part, relationship.target);
}

async function assertResolvedReferences(zip: JSZip, part: string): Promise<void> {
	const relationships = parseRelationships(
		await zip.file(relationshipsPartFor(part))?.async('string'),
	);
	const tree: unknown = new XMLParser({ ignoreAttributes: false }).parse(
		await zip.file(part)!.async('string'),
	);
	const visit = (node: unknown): void => {
		if (!node || typeof node !== 'object') return;
		for (const [key, value] of Object.entries(node)) {
			if (/^@_r:(?:id|embed|link)$/u.test(key))
				expect(relationships.has(String(value)), `${part}: ${value}`).toBe(true);
			else visit(value);
		}
	};
	visit(tree);
	const types = parseContentTypes(await zip.file('[Content_Types].xml')!.async('string'));
	for (const relationship of relationships.values()) {
		if (relationship.mode === 'External') continue;
		const target = resolvePartPath(part, relationship.target);
		expect(zip.file(target), target).not.toBeNull();
		expect(contentTypeForPart(types, target), target).toBeDefined();
	}
}

describe('duplicated slide package ownership', () => {
	it.each(['sdk', 'presentation', 'automation', 'editor-deep', 'editor-shallow'] as const)(
		'%s keeps chart, notes, images, media and hyperlinks resolved and independent',
		async (entry) => {
			const bytes = await sourceDeck();
			const presentation = await Presentation.load(buffer(bytes));
			const { handler, data } = presentation;
			const original = data.slides[0];
			if (entry === 'presentation') presentation.duplicateSlide(0);
			else if (entry === 'automation') duplicateToolSlide({ pptxData: data }, { slideIndex: 0 });
			else if (entry === 'sdk') data.slides.push(duplicateSlide(original, 2));
			else
				data.slides.push({
					...(entry === 'editor-deep' ? cloneSlide(original) : original),
					id: 'editor-copy',
				});
			const clone = data.slides[1];
			const saved = await handler.save(data.slides);
			const zip = await JSZip.loadAsync(saved);
			await assertResolvedReferences(zip, original.id);
			await assertResolvedReferences(zip, clone.id);
			for (const type of ['chart', 'notesSlide', 'image', 'media']) {
				const sourcePart = await relatedPart(zip, original.id, type);
				const clonedPart = await relatedPart(zip, clone.id, type);
				expect(clonedPart).not.toBe(sourcePart);
				if (type === 'notesSlide')
					expect(await relatedPart(zip, clonedPart, 'slide')).toBe(clone.id);
			}
			const cloneRels = parseRelationships(
				await zip.file(relationshipsPartFor(clone.id))!.async('string'),
			);
			expect([...cloneRels.values()].find((rel) => rel.type.endsWith('/hyperlink'))).toMatchObject({
				mode: 'External',
				target: 'https://example.com/?x=1&y=2',
			});
			expect(chart(clone).chartData!.chartPartPath).not.toBe(
				chart(original).chartData!.chartPartPath,
			);
			chart(clone).chartData!.series[0].values[0] = 42;
			clone.notes = 'Clone notes';
			clone.isDirty = true;
			expect(chart(original).chartData!.series[0].values[0]).toBe(0);
			const edited = await handler.save(data.slides);
			const reloaded = await new PptxHandler().load(buffer(edited));
			expect(chart(reloaded.slides[0]).chartData!.series[0].values).toEqual([0, 2, 3]);
			expect(chart(reloaded.slides[1]).chartData!.series[0].values).toEqual([42, 2, 3]);
			expect(reloaded.slides.map((slide) => slide.notes)).toEqual([
				'Original notes',
				'Clone notes',
			]);
			// Saving a copy again must not allocate another chart or notes part.
			expect(
				Object.keys((await JSZip.loadAsync(edited)).files).filter((path) =>
					/^ppt\/charts\/chart\d+\.xml$/u.test(path),
				),
			).toHaveLength(2);
		},
	);

	it('copies workbook, colors and style dependencies with content types and unchanged bytes', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(buffer(await sourceDeck(true)));
		data.slides.push(duplicateSlide(data.slides[0], 2));
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		const sourceChart = await relatedPart(zip, data.slides[0].id, 'chart');
		const clonedChart = await relatedPart(zip, data.slides[1].id, 'chart');
		await assertResolvedReferences(zip, data.slides[1].id);
		await assertResolvedReferences(zip, clonedChart);
		expect(await zip.file(clonedChart)!.async('string')).toBe(
			await zip.file(sourceChart)!.async('string'),
		);
		for (const type of ['package', 'chartColorStyle', 'chartStyle']) {
			const sourcePart = await relatedPart(zip, sourceChart, type);
			const clonedPart = await relatedPart(zip, clonedChart, type);
			expect(clonedPart).not.toBe(sourcePart);
			expect(await zip.file(clonedPart)!.async('uint8array')).toEqual(
				await zip.file(sourcePart)!.async('uint8array'),
			);
		}
	});

	it('uses a saved clone as the source of a later copy, not its ancestor', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(buffer(await sourceDeck()));
		data.slides.push(duplicateSlide(data.slides[0], 2));
		await handler.save(data.slides);
		chart(data.slides[1]).chartData!.series[0].values[0] = 12;
		data.slides[1].isDirty = true;
		await handler.save(data.slides);
		data.slides.push(duplicateSlide(data.slides[1], 3));
		chart(data.slides[2]).chartData!.series[0].values[0] = 24;
		data.slides[2].isDirty = true;
		const bytes = await handler.save(data.slides);
		const loaded = await new PptxHandler().load(buffer(bytes));
		expect(loaded.slides.map((slide) => chart(slide).chartData!.series[0].values[0])).toEqual([
			0, 12, 24,
		]);
		const zip = await JSZip.loadAsync(bytes);
		await assertResolvedReferences(zip, loaded.slides[2].id);
	});

	it('preserves an untouched PowerPoint-authored chart byte for byte', async () => {
		const bytes = readFileSync(
			new URL('../../../__tests__/fixtures/e2e/chart-data-fidelity.pptx', import.meta.url),
		);
		const handler = new PptxHandler();
		const data = await handler.load(buffer(bytes));
		data.slides.push(duplicateSlide(data.slides[0], data.slides.length + 1));
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		const sourcePart = await relatedPart(zip, data.slides[0].id, 'chart');
		const copiedPart = await relatedPart(zip, data.slides.at(-1)!.id, 'chart');
		expect(await zip.file(copiedPart)!.async('uint8array')).toEqual(
			await zip.file(sourcePart)!.async('uint8array'),
		);
		await assertResolvedReferences(zip, data.slides.at(-1)!.id);
		await assertResolvedReferences(zip, copiedPart);
	});

	it('reserves copied media names before embedding new media in the same save', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(buffer(await sourceDeck()));
		data.slides.push(duplicateSlide(data.slides[0], 2));
		data.slides.push(new SlideBuilder(3).addImage('data:image/png;base64,AQID').build());
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		const sourceImage = await relatedPart(zip, data.slides[0].id, 'image');
		const clonedImage = await relatedPart(zip, data.slides[1].id, 'image');
		const newImage = await relatedPart(zip, data.slides[2].id, 'image');
		expect(newImage).not.toBe(clonedImage);
		expect(await zip.file(clonedImage)!.async('uint8array')).toEqual(
			await zip.file(sourceImage)!.async('uint8array'),
		);
		expect(await zip.file(newImage)!.async('uint8array')).toEqual(new Uint8Array([1, 2, 3]));
		await assertResolvedReferences(zip, data.slides[2].id);
	});

	it('does not copy unrelated parts when a foreign blank slide has a colliding rId', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(buffer(await sourceDeck()));
		const blank = await PptxHandler.createBlank({ initialSlideCount: 1 });
		const unrelated = blank.data.slides[0];
		unrelated.id = 'unrelated-blank';
		unrelated.rId = data.slides[0].rId;
		data.slides.push(unrelated);
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		const relationships = parseRelationships(
			await zip.file(relationshipsPartFor(unrelated.id))!.async('string'),
		);
		expect(
			[...relationships.values()].some(
				(relationship) =>
					relationship.type.endsWith('/chart') || relationship.type.endsWith('/notesSlide'),
			),
		).toBe(false);
		expect(
			Object.keys(zip.files).filter((path) => /^ppt\/charts\/chart\d+\.xml$/u.test(path)),
		).toHaveLength(1);
	});

	it('duplicates an authored chart before its first save without sharing data', async () => {
		const { handler, data, createSlide } = await PptxHandler.createBlank();
		data.slides.push(
			createSlide()
				.setNotes('Notes')
				.addChart('line', { categories: ['A'], series: [{ name: 'One', values: [1] }] })
				.build(),
		);
		data.slides.push(duplicateSlide(data.slides[0], 2));
		chart(data.slides[1]).chartData!.series[0].values[0] = 2;
		const bytes = await handler.save(data.slides);
		const loaded = await new PptxHandler().load(buffer(bytes));
		expect(loaded.slides.map((slide) => chart(slide).chartData!.series[0].values[0])).toEqual([
			1, 2,
		]);
		expect(chart(loaded.slides[0]).chartData!.chartPartPath).not.toBe(
			chart(loaded.slides[1]).chartData!.chartPartPath,
		);
	});
});
