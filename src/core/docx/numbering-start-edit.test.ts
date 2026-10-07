import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { setListStartOverride } from './numbering-start-edit';
import { loadDocx } from './parse';
import { computeListLabels } from './numbering-format';
import { numberingStartFixture } from './test-support/numbering-start-fixture';
import { buildXml, children, getW, isElement, parseXml } from './xml';
import { numberingStartChanges, patchNumberingStarts } from './numbering-start-patch';

async function fixture(startOverride: number | null = 7, fullLevel = false) {
	const zip = await JSZip.loadAsync(
		await numberingStartFixture({
			name: 'edit',
			abstractRestart: null,
			abstractStart: 3,
			fullLevel,
			fullStart: fullLevel ? 5 : null,
			startOverride,
		}),
	);
	let xml = await zip.file('word/numbering.xml')!.async('string');
	xml = xml
		.replace('<w:numbering xmlns:w=', '<w:numbering xmlns:x="urn:opaque-numbering" xmlns:w=')
		.replace('<w:num w:numId="1">', '<w:num w:numId="1" x:source="keep">')
		.replace('<w:lvlOverride w:ilvl="2">', '<w:lvlOverride w:ilvl="2" x:level="keep">')
		.replace('<w:startOverride w:val="7"/>', '<w:startOverride w:val="7" x:identity="keep"/>')
		.replace('</w:lvlOverride>', '<x:levelExtension x:value="keep"/></w:lvlOverride>')
		.replace(
			'</w:num>',
			'<x:extension x:value="keep"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>',
		);
	zip.file('word/numbering.xml', xml);
	return { bytes: await zip.generateAsync({ type: 'uint8array' }), xml };
}

it('changes one start immutably and returns the original catalog for a no-op', async () => {
	const { bytes } = await fixture();
	const loaded = await loadDocx(bytes);
	const catalog = loaded.model.numberingCatalog!;
	const snapshot = structuredClone(catalog);
	expect(setListStartOverride(catalog, 1, 2, 7)).toBe(catalog);
	const next = setListStartOverride(catalog, 1, 2, 4);
	expect(catalog).toEqual(snapshot);
	expect(next.abstractNums).toBe(catalog.abstractNums);
	expect(next.nums['2']).toBe(catalog.nums['2']);
	expect(next.nums['1']!.levelOverrides![2]!.startOverride).toBe(4);
	expect(await loaded.save({ ...loaded.model, numberingCatalog: catalog })).toEqual(bytes);
});

it.each([4, 0])(
	'patches only the authored start %s while preserving opaque source and sibling definitions',
	async (start) => {
		const { bytes, xml } = await fixture();
		const loaded = await loadDocx(bytes);
		const model = {
			...loaded.model,
			numberingCatalog: setListStartOverride(loaded.model.numberingCatalog!, 1, 2, start),
		};
		const saved = await loaded.save(model);
		const output = await (await JSZip.loadAsync(saved)).file('word/numbering.xml')!.async('string');
		const before = parseXml(xml);
		const after = parseXml(output);
		expect(buildXml(children(after.documentElement, 'abstractNum')[0]!)).toBe(
			buildXml(children(before.documentElement, 'abstractNum')[0]!),
		);
		expect(buildXml(children(after.documentElement, 'num')[1]!)).toBe(
			buildXml(children(before.documentElement, 'num')[1]!),
		);
		const original = children(before.documentElement, 'num')[0]!;
		const changed = children(after.documentElement, 'num')[0]!;
		const startNode = children(children(original, 'lvlOverride')[0]!, 'startOverride')[0]!;
		startNode.setAttributeNS(
			'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
			'w:val',
			String(start),
		);
		expect(buildXml(changed)).toBe(buildXml(original));
		const reloaded = (await loadDocx(saved)).model;
		const labels = computeListLabels(reloaded);
		expect(reloaded.blocks.map((block) => labels.get(block.id)?.text)).toEqual([
			'1.',
			'1.',
			`${start}.`,
			`${start + 1}.`,
			'2.',
			`${start}.`,
			'2.',
			`${start}.`,
		]);
	},
);

it('adds a schema-ordered start wrapper without replacing existing num metadata', async () => {
	const { bytes } = await fixture(null);
	const loaded = await loadDocx(bytes);
	const saved = await loaded.save({
		...loaded.model,
		numberingCatalog: setListStartOverride(loaded.model.numberingCatalog!, 1, 2, 4),
	});
	const xml = await (await JSZip.loadAsync(saved)).file('word/numbering.xml')!.async('string');
	const num = children(parseXml(xml).documentElement, 'num')[0]!;
	expect(num.getAttributeNS('urn:opaque-numbering', 'source')).toBe('keep');
	expect(
		Array.from(num.childNodes)
			.filter(isElement)
			.map((node) => node.localName),
	).toEqual(['abstractNumId', 'lvlOverride', 'extension']);
	expect(getW(children(children(num, 'lvlOverride')[0]!, 'startOverride')[0], 'val')).toBe('4');
});

it('adds a missing start inside an opaque existing wrapper without replacing it', async () => {
	const source = await fixture(null);
	const zip = await JSZip.loadAsync(source.bytes);
	zip.file(
		'word/numbering.xml',
		source.xml.replace(
			'<x:extension',
			'<w:lvlOverride w:ilvl="2" x:level="keep"><x:levelExtension x:value="keep"/></w:lvlOverride><x:extension',
		),
	);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const saved = await loaded.save({
		...loaded.model,
		numberingCatalog: setListStartOverride(loaded.model.numberingCatalog!, 1, 2, 4),
	});
	const xml = await (await JSZip.loadAsync(saved)).file('word/numbering.xml')!.async('string');
	const wrapper = children(children(parseXml(xml).documentElement, 'num')[0]!, 'lvlOverride')[0]!;
	expect(wrapper.getAttributeNS('urn:opaque-numbering', 'level')).toBe('keep');
	expect(
		Array.from(wrapper.childNodes)
			.filter(isElement)
			.map((node) => node.localName),
	).toEqual(['startOverride', 'levelExtension']);
});

it.each([-1, 1.5, Infinity, NaN, 2147483648, Number.MAX_SAFE_INTEGER + 1])(
	'rejects invalid start %s',
	async (start) => {
		const loaded = await loadDocx((await fixture()).bytes);
		expect(() => setListStartOverride(loaded.model.numberingCatalog!, 1, 2, start)).toThrow(
			'integer from 0 through 2147483647',
		);
	},
);

it('accepts the signed32-bit authoring boundary without narrowing parsed definitions', async () => {
	const loaded = await loadDocx((await fixture()).bytes);
	const next = setListStartOverride(loaded.model.numberingCatalog!, 1, 2, 2147483647);
	const reloaded = await loadDocx(await loaded.save({ ...loaded.model, numberingCatalog: next }));
	expect(reloaded.model.numberingCatalog!.nums['1']!.levelOverrides![2]!.startOverride).toBe(
		2147483647,
	);
});

it('rejects missing instances/levels, invalid levels and full overrides', async () => {
	const loaded = await loadDocx((await fixture()).bytes);
	const catalog = loaded.model.numberingCatalog!;
	expect(() => setListStartOverride(catalog, 99, 2, 4)).toThrow('does not exist');
	expect(() => setListStartOverride(catalog, 1, 8, 4)).toThrow('does not exist');
	expect(() => setListStartOverride(catalog, 1, 9, 4)).toThrow('level');
	expect(() => setListStartOverride(catalog, 0, 2, 4)).toThrow('instance id');
	expect(() => setListStartOverride(catalog, 1.5, 2, 4)).toThrow('instance id');
	const full = (await loadDocx((await fixture(7, true)).bytes)).model.numberingCatalog!;
	expect(() => setListStartOverride(full, 1, 2, 4)).toThrow('full level override');
});

it('rejects direct model full-level and missing-level edits before package patching', async () => {
	const loaded = await loadDocx((await fixture(7, true)).bytes);
	const catalog = loaded.model.numberingCatalog!;
	const full = structuredClone(catalog);
	full.nums['1']!.levelOverrides![2]!.startOverride = 4;
	expect(() => numberingStartChanges(catalog, full)).toThrow('full level override');
	const missing = structuredClone(catalog);
	missing.nums['1']!.levelOverrides![8] = { startOverride: 4 };
	expect(() => numberingStartChanges(catalog, missing)).toThrow('does not exist');
});

it('rejects removal or unrelated definition changes before patching any source XML', async () => {
	const loaded = await loadDocx((await fixture()).bytes);
	const original = loaded.model.numberingCatalog!;
	const removal = structuredClone(original);
	delete removal.nums['1']!.levelOverrides![2]!.startOverride;
	expect(() => numberingStartChanges(original, removal)).toThrow('Removing');
	const unrelated = setListStartOverride(original, 1, 2, 4);
	unrelated.nums['2'] = { ...unrelated.nums['2']!, abstractNumId: 'bad' };
	expect(() => numberingStartChanges(original, unrelated)).toThrow(
		'Cannot edit numbering definition',
	);
	await expect(loaded.save({ ...loaded.model, numberingCatalog: unrelated })).rejects.toThrow();
	const source = parseXml(
		'<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
	);
	const before = buildXml(source);
	expect(() =>
		patchNumberingStarts(source, [
			{ numId: '1', level: 2, start: 4 },
			{ numId: '2', level: 2, start: 5 },
		]),
	).toThrow('missing');
	expect(buildXml(source)).toBe(before);
});

it('rejects ambiguous source wrappers before changing any start value', () => {
	const source = parseXml(
		'<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="2"><w:startOverride w:val="7"/></w:lvlOverride><w:lvlOverride w:ilvl="2"><w:startOverride w:val="8"/></w:lvlOverride></w:num></w:numbering>',
	);
	const before = buildXml(source);
	expect(() => patchNumberingStarts(source, [{ numId: '1', level: 2, start: 4 }])).toThrow(
		'duplicate',
	);
	expect(buildXml(source)).toBe(before);
});

it.each([
	'<w:num w:numId="01"><w:abstractNumId w:val="0"/></w:num>',
	'<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="01"><w:abstractNumId w:val="0"/></w:num>',
	'<w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="02"><w:startOverride w:val="7"/></w:lvlOverride></w:num>',
])('rejects misleading numeric source identities before patching', (body) => {
	const source = parseXml(
		`<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${body}</w:numbering>`,
	);
	const before = buildXml(source);
	expect(() => patchNumberingStarts(source, [{ numId: '1', level: 2, start: 4 }])).toThrow();
	expect(buildXml(source)).toBe(before);
});
