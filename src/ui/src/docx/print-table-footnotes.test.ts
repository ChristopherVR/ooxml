// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import { layoutDocumentModel } from 'ooxml-core/docx/layout';
import { renderPrintLayout } from './print-layout';
it.each([
	'split-late-reference',
	'repeated-header',
	'exact-clipped-reference',
	'mode14-split-long-note-columns',
])('%s renders each note source line once', async (name) => {
	const loaded = await loadDocx(
		new Uint8Array(
			readFileSync(
				new URL(
					'../../../core/docx/__fixtures__/table-footnotes/' + name + '.docx',
					import.meta.url,
				),
			),
		),
	);
	const layout = layoutDocumentModel(loaded.model, {
		widthOf: (text) => text.length * 5,
		lineHeightOf: () => 16,
	});
	const { element } = renderPrintLayout(layout),
		pages = [...element.querySelectorAll('.dve-print-page')];
	expect(pages).toHaveLength(layout.pages.length);
	for (const [index, page] of pages.entries()) {
		const actual = [...page.querySelectorAll('.dve-print-footnote')].flatMap(
			(note) => note.textContent?.match(/N\dL\d+/g) ?? [],
		);
		const expected = (layout.pages[index]!.footnotes ?? []).flatMap((note) =>
			note.paragraphs.flatMap((p) =>
				p.lines.flatMap(
					(line) =>
						line.fragments
							.map((f) => f.text)
							.join('')
							.match(/N\dL\d+/g) ?? [],
				),
			),
		);
		expect(actual).toEqual(expected);
	}
	if (name === 'repeated-header') {
		expect(element.querySelectorAll('[data-repeated="true"]')).toHaveLength(1);
		expect(pages[1]!.querySelectorAll('.dve-print-footnote')).toHaveLength(0);
	}
	expect(element.textContent?.match(/N0L00/g)).toHaveLength(1);
});
