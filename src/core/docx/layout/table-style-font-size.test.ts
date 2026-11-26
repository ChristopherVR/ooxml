import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import {
	editStyleCatalogs,
	loadDocx,
	resolveRunFormatting,
	resolveTableStyleFormatting,
} from '../index';
import type { DocumentModel } from '../model';
import { at, expectTable, must } from '../test-support/access';
import { adaptDocumentModel } from './adapter';
import { layoutDocumentModel } from './layout';
import { createFakeMeasurer } from './measure';

const load = async () =>
	loadDocx(
		await readFile(
			new URL(
				'./fixtures/word-layout-reports/issue-353-table-style-precedence.docx',
				import.meta.url,
			),
		),
	);

const cases = [
	{
		name: 'absent setting',
		normal: 11,
		override: undefined,
		inside: [
			[16, 18],
			[16, 18],
			[16, 10],
			[9, 18],
		],
	},
	{
		name: '10 pt Normal with the setting off',
		normal: 10,
		override: false,
		inside: [
			[10, 18],
			[10, 18],
			[10, 10],
			[9, 18],
		],
	},
	{
		name: 'enabled setting',
		normal: 11,
		override: true,
		inside: [
			[11, 12],
			[11, 12],
			[11, 10],
			[9, 12],
		],
	},
];

function resolvedSizes(model: DocumentModel) {
	const table = expectTable(at(model.blocks, 1));
	const cell = at(at(table.rows, 0), 0);
	const tableStyleRun = must(
		resolveTableStyleFormatting(table.style, model.tableStyles, table.look, 0, 1, 0, 1).run,
	);
	return cell.paragraphs.map((paragraph) => {
		const effective = resolveRunFormatting(at(paragraph.runs, 0), {
			paragraphCatalog: model.paragraphStyles,
			runCatalog: model.characterStyles,
			paragraphStyleId: paragraph.style,
			tableStyleRun,
			overrideTableStyleFontSizeAndJustification: model.overrideTableStyleFontSizeAndJustification,
		});
		return [effective.fontSize, effective.fontSizeComplexScript];
	});
}

function laidOutSizes(runs: readonly { text: string; fontSizePt?: number }[]) {
	return ['Lorem', 'لوريم'].map((text) => runs.find((run) => run.text.includes(text))?.fontSizePt);
}

it.each(cases)(
	'resolves and lays out table font sizes with $name (#353)',
	async ({ normal, override, inside }) => {
		const loaded = await load();
		let { model } = loaded;
		if (override !== undefined) {
			const edit = must(
				editStyleCatalogs(model, {
					kind: 'paragraph',
					id: 'Normal',
					name: 'Normal',
					paragraphFormatting: must(model.paragraphStyles?.styles.Normal).formatting,
					runFormatting: { fontSize: normal, fontSizeComplexScript: 12 },
				}),
			);
			({ model } = await loadDocx(
				await loaded.save({
					...model,
					...edit.catalogs,
					overrideTableStyleFontSizeAndJustification: override,
				}),
			));
		}
		expect(model.overrideTableStyleFontSizeAndJustification).toBe(override);
		expect.soft(resolvedSizes(model)).toEqual(inside);

		const [body, table] = at(adaptDocumentModel(model).sections, 0).blocks;
		if (body?.kind !== 'paragraph' || table?.kind !== 'table')
			throw new Error('Expected a paragraph and table');
		expect.soft(laidOutSizes(body.runs)).toEqual([normal, 12]);
		const paragraphs = at(at(table.rows, 0).cells, 0).paragraphs;
		expect.soft(paragraphs.map((paragraph) => laidOutSizes(paragraph.runs))).toEqual(inside);

		const { pages } = layoutDocumentModel(model, createFakeMeasurer());
		const placedTable = at(at(pages, 0).columns, 0).blocks.find((block) => block.kind === 'table');
		const placedCell = at(at(placedTable?.rows, 0).cells, 0);
		expect(
			placedCell.map((paragraph) =>
				laidOutSizes(paragraph.lines.flatMap((line) => line.fragments)),
			),
		).toEqual(inside);
	},
);
