import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { loadDocx } from './parse';
import { headerFooterForPage } from './layout/page-fields';
import type { LayoutPageBox } from './layout/result';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

it.each(['0', 'false', 'off'])(
	'enables imported explicit-off header/footer flags %s on save',
	async (value) => {
		const { bytes } = await fixture(value);
		const loaded = await loadDocx(bytes);
		loaded.model.evenAndOddHeaders = true;
		loaded.model.sections![0]!.titlePage = true;
		const saved = await loaded.save();
		const reopened = await loadDocx(saved);
		expect(reopened.model.evenAndOddHeaders).toBe(true);
		expect(reopened.model.sections![0]!.titlePage).toBe(true);
	},
);

async function fixture(value: string | undefined) {
	const zip = new JSZip();
	const flag = (name: string) =>
		value === undefined ? '' : `<w:${name}${value ? ` w:val="${value}"` : ''}/>`;
	const references: string[] = [];
	const relationships: string[] = [];
	for (const kind of ['header', 'footer'] as const) {
		for (const slot of ['default', 'even', 'first']) {
			const id = `${kind}-${slot}`;
			references.push(`<w:${kind}Reference w:type="${slot}" r:id="${id}"/>`);
			relationships.push(`<Relationship Id="${id}" Type="${r}/${kind}" Target="${id}.xml"/>`);
			const tag = kind === 'header' ? 'hdr' : 'ftr';
			zip.file(
				`word/${id}.xml`,
				`<w:${tag} xmlns:w="${w}"><w:p><w:r><w:t>${id}</w:t></w:r></w:p></w:${tag}>`,
			);
		}
	}
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr>${references.join('')}${flag('titlePg')}</w:sectPr></w:body></w:document>`,
	);
	const settings = `<w:settings xmlns:w="${w}">${flag('evenAndOddHeaders')}</w:settings>`;
	zip.file('word/settings.xml', settings);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join('')}</Relationships>`,
	);
	return { bytes: await zip.generateAsync({ type: 'uint8array' }), settings };
}

it.each([
	{ value: undefined, enabled: false },
	{ value: '0', enabled: false },
	{ value: 'false', enabled: false },
	{ value: 'off', enabled: false },
	{ value: '', enabled: true },
	{ value: '1', enabled: true },
	{ value: 'true', enabled: true },
	{ value: 'on', enabled: true },
])(
	'honors header/footer on-off flags $value through selection and body-edit saving',
	async ({ value, enabled }) => {
		const { bytes, settings } = await fixture(value);
		const loaded = await loadDocx(bytes);
		const assertSelection = (model: typeof loaded.model) => {
			expect(Boolean(model.evenAndOddHeaders)).toBe(enabled);
			expect(Boolean(model.sections![0]!.titlePage)).toBe(enabled);
			for (const kind of ['headers', 'footers'] as const) {
				const prefix = kind === 'headers' ? 'header' : 'footer';
				for (const [index, slot] of [
					[0, enabled ? 'first' : 'default'],
					[1, enabled ? 'even' : 'default'],
				] as const) {
					const page: LayoutPageBox = {
						index,
						sectionIndex: 0,
						pageInSection: index,
						widthPx: 816,
						heightPx: 1056,
						marginTopPx: 96,
						marginRightPx: 96,
						marginBottomPx: 96,
						marginLeftPx: 96,
						columns: [],
					};
					expect(headerFooterForPage(model, page, index + 1, kind)?.partName).toBe(
						`word/${prefix}-${slot}.xml`,
					);
				}
			}
		};
		assertSelection(loaded.model);
		expect(await loaded.save()).toEqual(bytes);
		const paragraph = loaded.model.blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('Expected body paragraph');
		paragraph.runs[0]!.text = 'Edited';
		const saved = await loaded.save();
		const output = await JSZip.loadAsync(saved);
		expect(await output.file('word/settings.xml')!.async('string')).toBe(settings);
		if (value !== undefined)
			expect(await output.file('word/document.xml')!.async('string')).toContain(
				value ? `<w:titlePg w:val="${value}"` : '<w:titlePg',
			);
		assertSelection((await loadDocx(saved)).model);
	},
);
