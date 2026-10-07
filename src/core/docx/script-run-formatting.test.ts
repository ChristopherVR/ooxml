import { describe, expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { loadDocx, saveDocx, resolveRunFormatting } from './index';
import { parseXml, WORD_NS } from './xml';
import { runHasUnknownProperties } from './write-run-validation';
import { expectParagraph, paragraphWithoutXmlBases } from './test-support/access';
import { rejectRevision, rejectAllRevisions } from './revision-commands';

const properties =
	'<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Yu Mincho" w:cs="Amiri"/><w:szCs w:val="29"/><w:bCs w:val="0"/><w:iCs/>';
const expected = {
	fontFamily: 'Arial',
	fontFamilyEastAsia: 'Yu Mincho',
	fontFamilyComplexScript: 'Amiri',
	fontSizeComplexScript: 14.5,
	boldComplexScript: false,
	italicComplexScript: true,
};
async function imported(extra = '') {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:r><w:rPr>${properties}${extra}</w:rPr><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

describe('script-specific run properties', () => {
	it.each(['alignment', 'spacing', 'indent', 'multiple'])(
		'exports tracked typing in native %s paragraph formatting',
		async (name) => {
			const fixture = new URL(
				`./__fixtures__/review-paragraph-formatting/${name}-tracked.docx`,
				import.meta.url,
			);
			const loaded = await loadDocx(new Uint8Array(await readFile(fixture)));
			const paragraph = expectParagraph(loaded.model.blocks[0]);
			const source = paragraph.runs[0]!;
			paragraph.runs = [
				{ ...source, text: source.text.slice(0, 4) },
				{
					...source,
					text: '!',
					revision: {
						kind: 'insert',
						id: 'script-font-text',
						author: 'Codex',
						date: '2026-10-07T00:00:00Z',
					},
				},
				{ ...source, text: source.text.slice(4) },
			];
			const reopened = await loadDocx(await loaded.save(loaded.model));
			const exported = expectParagraph(reopened.model.blocks[0]);
			expect(exported.formatRevision).toEqual(paragraph.formatRevision);
			expect(exported.runs[1]?.revision?.kind).toBe('insert');
			for (const run of exported.runs) expect(run.fontFamilyComplexScript).toBe('Arial');
			const rejected = await loadDocx(await reopened.save(rejectAllRevisions(reopened.model)));
			const native = await loadDocx(
				new Uint8Array(
					await readFile(
						new URL(
							`./__fixtures__/review-paragraph-formatting/${name}-rejected.docx`,
							import.meta.url,
						),
					),
				),
			);
			const actual = expectParagraph(rejected.model.blocks[0]);
			const expected = expectParagraph(native.model.blocks[0]);
			expect({ ...paragraphWithoutXmlBases(actual), runs: undefined }).toEqual({
				...paragraphWithoutXmlBases(expected),
				runs: undefined,
			});
			expect(actual.runs.map((run) => run.text).join('')).toBe(
				expected.runs.map((run) => run.text).join(''),
			);
			const { text: _text, ...format } = expected.runs[0]!;
			for (const run of actual.runs) {
				const { text: _piece, ...props } = run;
				expect(props).toEqual(format);
			}
		},
	);
	it('preserves properties through tracked run splitting and standalone export', async () => {
		const loaded = await imported();
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(paragraph.runs[0]).toMatchObject(expected);
		const source = paragraph.runs[0]!;
		paragraph.runs = [
			{ ...source, text: 'Te' },
			{ ...source, text: '!', revision: { id: 'new', kind: 'insert', author: 'Ada' } },
			{ ...source, text: 'xt' },
		];
		for (const bytes of [await loaded.save(loaded.model), await saveDocx(loaded.model)]) {
			const reopened = await loadDocx(bytes);
			const runs = expectParagraph(reopened.model.blocks[0]).runs;
			expect(runs.map((run) => run.text)).toEqual(['Te', '!', 'xt']);
			for (const run of runs) expect(run).toMatchObject(expected);
			expect(runs[1]?.revision).toMatchObject({ kind: 'insert', author: 'Ada' });
		}
	});

	it('changes and removes independent script settings without losing Latin formatting', async () => {
		const loaded = await imported();
		const run = expectParagraph(loaded.model.blocks[0]).runs[0]!;
		run.fontFamilyComplexScript = 'Noto Naskh Arabic';
		run.fontSizeComplexScript = 18;
		run.boldComplexScript = true;
		run.italicComplexScript = false;
		delete run.fontFamilyEastAsia;
		let reopened = await loadDocx(await loaded.save(loaded.model));
		expect(expectParagraph(reopened.model.blocks[0]).runs[0]).toMatchObject({
			fontFamily: 'Arial',
			fontFamilyComplexScript: 'Noto Naskh Arabic',
			fontSizeComplexScript: 18,
			boldComplexScript: true,
			italicComplexScript: false,
		});
		delete run.fontFamilyComplexScript;
		delete run.fontSizeComplexScript;
		delete run.boldComplexScript;
		delete run.italicComplexScript;
		reopened = await loadDocx(await loaded.save(loaded.model));
		expect(expectParagraph(reopened.model.blocks[0]).runs[0]).toEqual({
			text: 'Text',
			fontFamily: 'Arial',
		});
	});

	it('restores all script settings from a prior format snapshot', async () => {
		const loaded = await imported(
			'<w:rPrChange w:id="format" w:author="Ada"><w:rPr><w:rFonts w:cs="Original"/><w:szCs w:val="22"/><w:bCs/></w:rPr></w:rPrChange>',
		);
		const model = rejectRevision(loaded.model, 'format');
		const reopened = await loadDocx(await loaded.save(model));
		expect(expectParagraph(reopened.model.blocks[0]).runs[0]).toEqual({
			text: 'Text',
			fontFamilyComplexScript: 'Original',
			fontSizeComplexScript: 11,
			boldComplexScript: true,
		});
	});

	it('resolves script-specific direct overrides and style toggles separately', () => {
		const context = { runCatalog: { docDefaults: expected, styles: {}, warnings: [] } };
		expect(
			resolveRunFormatting(
				{ text: 'Text', boldComplexScript: true, fontSizeComplexScript: 10 },
				context,
			),
		).toMatchObject({ ...expected, boldComplexScript: true, fontSizeComplexScript: 10 });
	});

	it('allows known properties while retaining the guard for unknown extensions', () => {
		const node = (extra: string) =>
			parseXml(
				`<w:r xmlns:w="${WORD_NS}"><w:rPr>${properties}${extra}</w:rPr><w:t>Text</w:t></w:r>`,
			).documentElement;
		expect(runHasUnknownProperties(node(''))).toBe(false);
		expect(runHasUnknownProperties(node('<w:szCs w:val="invalid"/>'))).toBe(true);
		expect(runHasUnknownProperties(node('<w:customProperty/>'))).toBe(true);
	});
});
