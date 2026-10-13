import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import evidence from './__fixtures__/print-count-native.json';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { visioPrintTileCount } from './print-count';
import { buildXml, parseXml } from '../xml/index';
import { child, readSheet } from './sheet';
import { readVisioPageSetup } from './page-setup';
import { visioPageBreaks } from './ui/page-setup';

const sources = await JSZip.loadAsync(
	readFileSync(new URL('./__fixtures__/print-count-native.zip', import.meta.url)),
);
const digest = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');

it.each(evidence.cases)('native count/source preservation: $id', async (test) => {
	const original = await sources.file(test.file)!.async('uint8array');
	expect(digest(original)).toBe(test.sha256);
	const before = await parseVsdx(original);
	const page = before.pages[0]!;
	expect(page.width).toBeCloseTo(test.physicalWidth, 10);
	expect(page.height).toBeCloseTo(test.physicalHeight, 10);
	const sourceZip = await JSZip.loadAsync(original);
	const sourcePages = parseXml(await sourceZip.file('visio/pages/pages.xml')!.async('string'));
	const legacySetup = readVisioPageSetup(
		readSheet(child(child(sourcePages.documentElement, 'Page'), 'PageSheet')).cells,
	);
	const { pageSetup: _ignoredSetup, ...withoutSetup } = page;
	const legacyPage = {
		...withoutSetup,
		...(legacySetup === undefined ? {} : { pageSetup: legacySetup }),
	};
	expect(visioPageBreaks(page)).toEqual(visioPageBreaks(legacyPage));
	const result = visioPrintTileCount(page);
	const precisionRefusals = new Set([
		'extension/letter-asymmetric-x-0.1-0.7',
		'extension/letter-asymmetric-y-0.1-0.7',
		'extension/a4-landscape-x-first-0',
		'extension/a4-landscape-x-second-0',
		'extension/a4-landscape-y-second-0',
	]);
	if (precisionRefusals.has(test.id))
		expect(result).toEqual({ kind: 'unsupported', reason: 'boundary-precision' });
	else if (test.candidateProfile)
		expect(result).toMatchObject({ kind: 'supported', total: test.nativeCount });
	else expect(result).toEqual({ kind: 'unsupported', reason: test.excludedReason });
	const noop = await editVsdx(original, []);
	expect(noop.bytes).toEqual(original);
	expect(noop.changedParts).toEqual([]);
	const renamed = await editVsdx(original, [
		{ type: 'rename-page', pageId: page.id, name: 'Count preservation' },
	]);
	expect(renamed.changedParts).toEqual(['visio/pages/pages.xml', 'docProps/app.xml']);
	const oldZip = await JSZip.loadAsync(original),
		newZip = await JSZip.loadAsync(renamed.bytes);
	expect(Object.keys(newZip.files).sort()).toEqual(Object.keys(oldZip.files).sort());
	for (const [name, entry] of Object.entries(oldZip.files))
		if (!entry.dir && !['visio/pages/pages.xml', 'docProps/app.xml'].includes(name))
			expect(await newZip.file(name)!.async('uint8array'), name).toEqual(
				await entry.async('uint8array'),
			);
	// Compare normalized XML after restoring only the existing page name/title.
	const oldApp = parseXml(await oldZip.file('docProps/app.xml')!.async('string'));
	const newApp = parseXml(await newZip.file('docProps/app.xml')!.async('string'));
	const titles = Array.from(newApp.getElementsByTagNameNS('*', 'TitlesOfParts'));
	expect(titles).toHaveLength(1);
	const changedTitles = Array.from(titles[0]!.getElementsByTagNameNS('*', 'lpstr')).filter(
		(node) => node.textContent === 'Count preservation',
	);
	expect(changedTitles).toHaveLength(1);
	changedTitles[0]!.textContent = page.name;
	expect(buildXml(newApp).replace(/^<\?xml[^?]*\?>\s*/, '')).toBe(
		buildXml(oldApp).replace(/^<\?xml[^?]*\?>\s*/, ''),
	);
	const newPages = parseXml(await newZip.file('visio/pages/pages.xml')!.async('string'));
	const oldPage = child(sourcePages.documentElement, 'Page')!;
	const renamedPage = child(newPages.documentElement, 'Page')!;
	for (const name of ['Name', 'NameU']) {
		const value = oldPage.getAttribute(name);
		if (value === null) renamedPage.removeAttribute(name);
		else renamedPage.setAttribute(name, value);
	}
	expect(buildXml(newPages).replace(/^<\?xml[^?]*\?>\s*/, '')).toBe(
		buildXml(sourcePages).replace(/^<\?xml[^?]*\?>\s*/, ''),
	);
	const after = await parseVsdx(renamed.bytes);
	expect(after.pages[0]!.pageSetup).toEqual(page.pageSetup);
	expect(visioPrintTileCount(after.pages[0]!)).toEqual(result);
	// A core save/reparse is not native reopened acceptance; the evidence declares zero.
});
