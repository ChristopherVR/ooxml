/**
 * The property writers against real `docProps/*.xml` parts from Excel, openpyxl, Word and
 * PowerPoint packages: a write that changes nothing returns the source byte for byte, and a
 * change touches only the element it is about.
 */
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
	formatW3cdtf,
	parseAppProperties,
	parseCoreProperties,
	parseCustomPropertyTexts,
	writeAppProperties,
	writeCoreProperties,
} from './index';

const FIXTURES = {
	excel: '../../xlsx/__fixtures__/excel-features.xlsx',
	openpyxl: '../../xlsx/__fixtures__/openpyxl-features.xlsx',
	word: '../../docx/layout/fixtures/continuous-page-fields/default.docx',
	powerpoint: '../../pptx/__tests__/fixtures/e2e/comment-mentions.pptx',
	prefixedApp: '../../pptx/__tests__/fixtures/e2e/absolute-path-rels.pptx',
} as const;

async function part(fixture: keyof typeof FIXTURES, name: string): Promise<string> {
	const bytes = readFileSync(fileURLToPath(new URL(FIXTURES[fixture], import.meta.url)));
	const file = (await JSZip.loadAsync(bytes)).file(`docProps/${name}`);
	if (!file) throw new Error(`${fixture} has no docProps/${name}`);
	return file.async('string');
}

const names = Object.keys(FIXTURES) as (keyof typeof FIXTURES)[];

describe('property parts from real packages', () => {
	it.each(names)('rewrites %s core.xml and app.xml unchanged, byte for byte', async (fixture) => {
		const core = await part(fixture, 'core.xml');
		const app = await part(fixture, 'app.xml');
		expect(writeCoreProperties(parseCoreProperties(core), core)).toBe(core);
		expect(writeAppProperties(parseAppProperties(app), app)).toBe(app);
	});

	it('changes only the edited core elements, keeping the declaration line break', async () => {
		const core = await part('powerpoint', 'core.xml');
		const props = parseCoreProperties(core);
		const out = writeCoreProperties(
			{ ...props, revision: '2', modified: '2026-10-08T09:30:00Z', title: 'R&D <1>' },
			core,
		);
		expect(out).toBe(
			core
				.replace('<dc:title></dc:title>', '<dc:title>R&amp;D &lt;1&gt;</dc:title>')
				.replace('<cp:revision>1</cp:revision>', '<cp:revision>2</cp:revision>')
				.replace('2026-08-13T11:40:40Z', '2026-10-08T09:30:00Z'),
		);
	});

	it('patches a part whose namespaces are declared per element (openpyxl)', async () => {
		const core = await part('openpyxl', 'core.xml');
		const out = writeCoreProperties(
			{ ...parseCoreProperties(core), modified: '2026-10-08T09:30:00Z', subject: 'Q4' },
			core,
		);
		expect(out).toContain(
			'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">',
		);
		expect(out).toContain('<dc:subject>Q4</dc:subject></cp:coreProperties>');
		expect(parseCoreProperties(out)).toMatchObject({
			title: 'Fixture',
			subject: 'Q4',
			modified: '2026-10-08T09:30:00Z',
		});
	});

	it('reads the document statistics of Word and PowerPoint', async () => {
		expect(parseAppProperties(await part('word', 'app.xml'))).toMatchObject({
			pages: 3,
			words: 147,
			characters: 843,
			lines: 7,
			paragraphs: 1,
			charactersWithSpaces: 989,
		});
		expect(parseAppProperties(await part('powerpoint', 'app.xml'))).toMatchObject({
			presentationFormat: 'Widescreen',
			words: 0,
			paragraphs: 0,
			slides: 1,
			notes: 0,
			hiddenSlides: 0,
			mmClips: 0,
		});
	});

	it('updates PowerPoint statistics in place', async () => {
		const app = await part('powerpoint', 'app.xml');
		const out = writeAppProperties(
			{ ...parseAppProperties(app), slides: 3, notes: 2, words: 12 },
			app,
		);
		expect(out).toBe(
			app
				.replace('<Slides>1</Slides>', '<Slides>3</Slides>')
				.replace('<Notes>0</Notes>', '<Notes>2</Notes>')
				.replace('<Words>0</Words>', '<Words>12</Words>'),
		);
	});

	it('writes into a prefixed extended-properties root with its own prefix', async () => {
		const app = await part('prefixedApp', 'app.xml');
		const props = parseAppProperties(app);
		expect(props).toMatchObject({ slides: 12, notes: 1, words: 554 });
		const out = writeAppProperties({ ...props, slides: 13, lines: 4 }, app);
		expect(out).toBe(
			app
				.replace('<ap:Slides>12</ap:Slides>', '<ap:Slides>13</ap:Slides>')
				.replace('<ap:Paragraphs>', '<ap:Lines>4</ap:Lines><ap:Paragraphs>'),
		);
	});

	it('reads custom properties as text', async () => {
		expect(parseCustomPropertyTexts(await part('prefixedApp', 'custom.xml'))).toEqual([
			{
				name: 'ContentTypeId',
				type: 'lpwstr',
				text: '0x01010079F111ED35F8CC479449609E8A0923A6',
				pid: 2,
			},
		]);
	});

	it('formats W3CDTF dates at second precision', () => {
		expect(formatW3cdtf(new Date('2026-10-08T09:30:00.123Z'))).toBe('2026-10-08T09:30:00Z');
	});
});
