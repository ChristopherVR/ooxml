import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
import { centuryGothicMeasurer } from './test-support/century-gothic-measurer';

it.each([
	{ mode: 15, name: 'justify-compat15', pageCount: 2, endings: ['aliqua', 'aute', 'velit'] },
	{ mode: 14, name: 'justify-compat14-control', pageCount: 3, endings: ['', '', ''] },
])(
	'wraps the reported lines in compatibility mode $mode (#346)',
	async ({ mode, name, pageCount, endings }) => {
		const { model } = await loadDocx(
			await readFile(
				new URL(`./fixtures/word-layout-reports/issue-346-${name}.docx`, import.meta.url),
			),
		);
		expect(model.compatibilityMode).toBe(mode);
		const { pages } = layoutDocumentModel(model, centuryGothicMeasurer);
		const lines = pages.flatMap((page) =>
			page.columns.flatMap((column) =>
				column.blocks.flatMap((block) =>
					block.kind === 'paragraph'
						? block.lines.map((line) => line.fragments.map((fragment) => fragment.text).join(''))
						: [],
				),
			),
		);
		const prefixes = [
			'proident enim do aliqua exercitation',
			'pariatur irure adipiscing dolor in labore',
			'magna fugiat enim culpa eiusmod sunt',
		];
		for (const [index, prefix] of prefixes.entries())
			expect(lines.find((line) => line.startsWith(prefix))).toBe(
				`${prefix} ${endings[index] ?? ''}`.trimEnd(),
			);
		expect(pages).toHaveLength(pageCount);
	},
);
