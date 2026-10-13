import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { buildXml, parseXml } from '../xml/index';
import { attribute, child, children, VISIO_NS } from './sheet';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { visioPrintTileCount } from './print-count';

const corpus = await JSZip.loadAsync(
	readFileSync(new URL('./__fixtures__/print-count-native.zip', import.meta.url)),
);
const original = await corpus.file('initial/letter-portrait.vsdx')!.async('uint8array');
const patch = async (mutate: (page: Element, document: Element) => void) => {
	const zip = await JSZip.loadAsync(original);
	const pages = parseXml(await zip.file('visio/pages/pages.xml')!.async('string'));
	const document = parseXml(await zip.file('visio/document.xml')!.async('string'));
	mutate(child(pages.documentElement, 'Page')!, document.documentElement);
	zip.file('visio/pages/pages.xml', buildXml(pages));
	zip.file('visio/document.xml', buildXml(document));
	return zip.generateAsync({ type: 'uint8array' });
};
const local = (page: Element, name: string, attributes: Record<string, string>) => {
	const sheet = child(page, 'PageSheet')!;
	for (const node of children(sheet, 'Cell'))
		if (attribute(node, 'N') === name) sheet.removeChild(node);
	const node = page.ownerDocument!.createElementNS(VISIO_NS, 'Cell');
	node.setAttribute('N', name);
	for (const [key, value] of Object.entries(attributes)) node.setAttribute(key, value);
	sheet.appendChild(node);
	return node;
};
const unchanged = async (before: Uint8Array, after: Uint8Array, except: readonly string[]) => {
	const a = await JSZip.loadAsync(before),
		b = await JSZip.loadAsync(after);
	const parts = (zip: JSZip) =>
		Object.entries(zip.files)
			.filter(([, entry]) => !entry.dir)
			.map(([name]) => name)
			.sort();
	expect(parts(b)).toEqual(parts(a));
	for (const [name, entry] of Object.entries(a.files))
		if (!entry.dir && !except.includes(name))
			expect(await b.file(name)!.async('uint8array'), name).toEqual(
				await entry.async('uint8array'),
			);
};

describe('unsupported count metadata preservation and edit admission', () => {
	it.each([
		['malformed', { V: 'bad' }, 'invalid-local-setting'],
		['dynamic', { V: '9', F: 'PageWidth' }, 'dynamic-setting'],
		['inherited marker', { V: '9', F: 'Inh' }, 'invalid-local-setting'],
		['error', { V: '9', E: '' }, 'invalid-local-setting'],
	] as const)(
		'preserves explicit %s local cell without falling back',
		async (_, attributes, reason) => {
			const bytes = await patch((page) => {
				local(page, 'PaperKind', attributes);
			});
			const before = await parseVsdx(bytes);
			expect(visioPrintTileCount(before.pages[0]!)).toEqual({ kind: 'unsupported', reason });
			expect((await editVsdx(bytes, [])).bytes).toEqual(bytes);
			const renamed = await editVsdx(bytes, [
				{ type: 'rename-page', pageId: before.pages[0]!.id, name: 'Unrelated' },
			]);
			await unchanged(bytes, renamed.bytes, ['visio/pages/pages.xml', 'docProps/app.xml']);
			const after = await parseVsdx(renamed.bytes);
			expect(after.pages[0]!.shapes).toEqual(before.pages[0]!.shapes);
			expect(after.pages[0]!.pageSetup).toEqual(before.pages[0]!.pageSetup);
			const a = await JSZip.loadAsync(bytes),
				b = await JSZip.loadAsync(renamed.bytes);
			const read = (xml: string) =>
				children(child(child(parseXml(xml).documentElement, 'Page'), 'PageSheet'), 'Cell')
					.filter((node) => attribute(node, 'N') === 'PaperKind')
					.map((node) => buildXml(node));
			expect(read(await b.file('visio/pages/pages.xml')!.async('string'))).toEqual(
				read(await a.file('visio/pages/pages.xml')!.async('string')),
			);
		},
	);
	it.each(['PagesX', 'PagesY'])(
		'preserves invalid explicit %s after refusal and unrelated rename',
		async (name) => {
			const bytes = await patch((page) => {
				local(page, name, { V: '-1.5' });
			});
			const before = await parseVsdx(bytes);
			expect(visioPrintTileCount(before.pages[0]!)).toEqual({
				kind: 'unsupported',
				reason: 'invalid-local-setting',
			});
			expect((await editVsdx(bytes, [])).bytes).toEqual(bytes);
			const saved = await editVsdx(bytes, [
				{ type: 'rename-page', pageId: before.pages[0]!.id, name: 'Invalid target preserved' },
			]);
			await unchanged(bytes, saved.bytes, ['visio/pages/pages.xml', 'docProps/app.xml']);
			const after = await parseVsdx(saved.bytes);
			expect(after.pages[0]!.pageSetup).toEqual(before.pages[0]!.pageSetup);
			expect(visioPrintTileCount(after.pages[0]!)).toEqual({
				kind: 'unsupported',
				reason: 'invalid-local-setting',
			});
			const zip = await JSZip.loadAsync(saved.bytes);
			const cells = children(
				child(
					child(
						parseXml(await zip.file('visio/pages/pages.xml')!.async('string')).documentElement,
						'Page',
					),
					'PageSheet',
				),
				'Cell',
			);
			expect(
				cells.filter((node) => attribute(node, 'N') === name).map((node) => attribute(node, 'V')),
			).toEqual(['-1.5']);
		},
	);
	it('retains duplicate local print cells on no-op and unrelated edit', async () => {
		const bytes = await patch((page) => {
			const node = local(page, 'PaperKind', { V: '1' });
			node.parentNode!.appendChild(node.cloneNode(true));
		});
		const before = await parseVsdx(bytes);
		expect(visioPrintTileCount(before.pages[0]!)).toEqual({
			kind: 'unsupported',
			reason: 'duplicate-setting',
		});
		expect((await editVsdx(bytes, [])).bytes).toEqual(bytes);
		const saved = await editVsdx(bytes, [
			{ type: 'rename-page', pageId: before.pages[0]!.id, name: 'Duplicate preserved' },
		]);
		await unchanged(bytes, saved.bytes, ['visio/pages/pages.xml', 'docProps/app.xml']);
		expect(visioPrintTileCount((await parseVsdx(saved.bytes)).pages[0]!)).toEqual({
			kind: 'unsupported',
			reason: 'duplicate-setting',
		});
	});
	it('does not grant edit admission to a protected local cell', async () => {
		const bytes = await patch((page) => {
			local(page, 'PrintPageOrientation', { V: '1', F: 'GUARD(1)' });
		});
		const parsed = await parseVsdx(bytes);
		expect(visioPrintTileCount(parsed.pages[0]!)).toEqual({
			kind: 'unsupported',
			reason: 'dynamic-setting',
		});
		await expect(
			editVsdx(bytes, [
				{ type: 'set-page-setup', pageId: parsed.pages[0]!.id, printOrientation: 'landscape' },
			]),
		).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
		expect((await editVsdx(bytes, [])).bytes).toEqual(bytes);
	});
	it('an explicit local override never writes inherited document/style defaults', async () => {
		const parsed = await parseVsdx(original);
		expect(parsed.pages[0]!.pageSetup?.printPageOrientation).toBeUndefined();
		expect(parsed.pages[0]!.pageSetup?.printCountSettings?.printPageOrientation).toBe(1);
		const saved = await editVsdx(original, [
			{ type: 'set-page-setup', pageId: parsed.pages[0]!.id, printOrientation: 'landscape' },
		]);
		expect(saved.changedParts).toEqual(['visio/pages/pages.xml']);
		await unchanged(original, saved.bytes, ['visio/pages/pages.xml']);
		expect((await parseVsdx(saved.bytes)).pages[0]!.pageSetup?.printPageOrientation).toBe(2);
	});
});
