import { describe, expect, it } from 'vitest';
import { scriptLayoutRuns } from './script-fonts';
import { adaptDocumentModel } from './adapter';
import type { DocumentModel } from '../model';
import type { LayoutParagraph } from './input';
import { layoutDocumentModel } from './layout';

describe('Word script font layout', () => {
	it('preserves source lengths, formatting and the final hard break', () => {
		const pieces = scriptLayoutRuns(
			{ text: 'A中عB', sourceLength: 4, bold: true, breakAfter: 'page' },
			{ fontFamilyEastAsia: 'SimSun', fontFamilyComplexScript: 'Arabic Face' },
			undefined,
		);
		expect(
			pieces.map((run) => [run.text, run.fontFamily, run.sourceLength, run.breakAfter]),
		).toEqual([
			['A', 'Calibri', 1, undefined],
			['中', 'SimSun', 1, undefined],
			['ع', 'Arabic Face', 1, undefined],
			['B', 'Calibri', 1, 'page'],
		]);
		expect(pieces.every((run) => run.bold)).toBe(true);
	});
	it('does not split generated or hidden source mappings', () => {
		for (const run of [
			{ text: 'A中', sourceLength: 1 },
			{ text: 'A中', hidden: true },
		]) {
			expect(scriptLayoutRuns(run, { fontFamilyEastAsia: 'SimSun' }, undefined)).toEqual([run]);
		}
	});
	it('uses inherited and theme script fonts in the actual document adapter', () => {
		const model: DocumentModel = {
			blocks: [
				{ type: 'paragraph', id: 'p', runs: [{ text: 'A中عB', fontTheme: { cs: 'major' } }] },
			],
			page: {
				width: 816,
				height: 1056,
				marginTop: 96,
				marginBottom: 96,
				marginLeft: 96,
				marginRight: 96,
			},
			warnings: [],
			characterStyles: { docDefaults: { fontFamilyEastAsia: 'SimSun' }, styles: {}, warnings: [] },
			theme: {
				colors: {},
				colorMapping: {},
				fonts: { major: { complexScript: 'Arabic Face' }, minor: {} },
			},
		};
		const paragraph = adaptDocumentModel(model).sections[0]?.blocks[0] as LayoutParagraph;
		expect(paragraph.runs.map((run) => run.fontFamily)).toEqual([
			'Calibri',
			'SimSun',
			'Arabic Face',
			'Calibri',
		]);
		expect(paragraph.runs.reduce((sum, run) => sum + (run.sourceLength ?? 0), 0)).toBe(4);
		expect(model.blocks[0]).toMatchObject({ runs: [{ text: 'A中عB' }] });
		const measured: Array<[string, string]> = [];
		const result = layoutDocumentModel(model, {
			widthOf(text, font) {
				measured.push([text, font.family]);
				return text.length * 10;
			},
			lineHeightOf() {
				return 16;
			},
		});
		expect(result.pages).toHaveLength(1);
		expect(measured).toContainEqual(['中', 'SimSun']);
		expect(measured).toContainEqual(['ع', 'Arabic Face']);
	});
});
