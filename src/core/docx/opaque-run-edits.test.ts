import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index';
import { expectParagraph, must } from './test-support/access';
import { WORD_NS } from './xml';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const GLOW = '<w14:glow w14:rad="63500"><w14:schemeClr w14:val="accent1"/></w14:glow>';
const EFFECT = '<w:effect w:val="shimmer"/>';

async function load(runs: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}" xmlns:w14="${W14}"><w:body><w:p>${runs}</w:p><w:sectPr/></w:body></w:document>`,
	);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	return { loaded, paragraph: expectParagraph(loaded.model.blocks[0]) };
}
const styled = (text: string, props: string) =>
	`<w:r><w:rPr>${props}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;

async function xmlOf(loaded: Awaited<ReturnType<typeof load>>['loaded']) {
	const bytes = await loaded.save(loaded.model);
	const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
	return { xml, reopened: expectParagraph((await loadDocx(bytes)).model.blocks[0]) };
}

describe('unmodelled run properties through run edits', () => {
	it('copies glow and effect to both halves of a split run', async () => {
		const { loaded, paragraph } = await load(styled('Shiny text', `<w:b/>${EFFECT}${GLOW}`));
		const original = must(paragraph.runs[0]);
		paragraph.runs = [
			{ ...original, text: 'Shiny' },
			{ ...original, text: ' text', italic: true },
		];
		const { xml, reopened } = await xmlOf(loaded);
		expect(xml.match(/<w14:glow /g)).toHaveLength(2);
		expect(xml.match(/<w:effect w:val="shimmer"\/>/g)).toHaveLength(2);
		expect(reopened.runs.map((run) => run.text).join('')).toBe('Shiny text');
	});

	it('lets typed text inherit the properties, plain or tracked', async () => {
		const { loaded, paragraph } = await load(styled('Shiny', `${EFFECT}${GLOW}`));
		const original = must(paragraph.runs[0]);
		paragraph.runs = [
			original,
			{
				...original,
				text: '!',
				revision: { id: 'tracked-1', kind: 'insert', author: 'Ada', date: '2026-10-29T00:00:00Z' },
			},
		];
		const { xml, reopened } = await xmlOf(loaded);
		expect(xml.match(/<w14:glow /g)).toHaveLength(2);
		expect(reopened.runs[1]?.revision?.kind).toBe('insert');
		expect(reopened.runs[1]?.sourceRunPropertiesXml).toContain('glow');
	});

	it('merges runs whose opaque properties are equal', async () => {
		const { loaded, paragraph } = await load(
			styled('Sh', `${EFFECT}${GLOW}`) + styled('iny', `${EFFECT}${GLOW}`),
		);
		const [a] = paragraph.runs;
		paragraph.runs = [{ ...must(a), text: 'Shiny' }];
		const { xml, reopened } = await xmlOf(loaded);
		expect(reopened.runs).toHaveLength(1);
		expect(reopened.runs[0]?.text).toBe('Shiny');
		expect(xml.match(/<w14:glow /g)).toHaveLength(1);
	});

	it('refuses to merge runs whose opaque properties differ, with a diagnostic', async () => {
		const { loaded, paragraph } = await load(
			styled('Sh', EFFECT) + styled('iny', '<w:effect w:val="sparkle"/>'),
		);
		const [a] = paragraph.runs;
		paragraph.runs = [{ ...must(a), text: 'Shiny' }];
		await expect(loaded.save(loaded.model)).rejects.toThrow(
			/Cannot edit paragraph .*changing run boundaries could drop unsupported run properties/,
		);
	});
});
