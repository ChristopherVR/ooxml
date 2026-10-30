import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { createVsdx } from './create-document';
import { snapshotPageTheme } from './edit-page-theme-commands';
import { visioFontAssumed } from './style';
import { visioBuiltInTheme } from './theme-builtins';
import { copySnapshotScene } from './ui/snapshot-scene';
import { assertViewableDocument } from './ui/scene-validation';
import { attribute, child, children } from './sheet';

const page = '0';
const THEMED_CELLS =
	'<Cell N="ShdwPattern" V="Themed" F="THEMEVAL()"/><Cell N="GlowSize" V="Themed" F="THEMEVAL()"/><Cell N="SoftEdgesSize" V="Themed" F="THEMEVAL()"/><Cell N="ReflectionSize" V="Themed" F="THEMEVAL()"/><Cell N="QuickStyleEffectsMatrix" V="4"/>';
/** A text box with themed text and themed effects, and a line with a themed weight. */
async function drawing(): Promise<Uint8Array> {
	const drawn = await editVsdx(await createVsdx(), [
		{
			type: 'create-text-box',
			pageId: page,
			shapeId: '1',
			x: 1,
			y: 1,
			width: 2,
			height: 1,
			text: 'Hi',
		},
		{ type: 'create-line', pageId: page, shapeId: '2', beginX: 4, beginY: 1, endX: 6, endY: 1 },
	]);
	const zip = await JSZip.loadAsync(drawn.bytes);
	let xml = await zip.file('visio/pages/page1.xml')!.async('string');
	xml = xml.replace(
		'<Section N="Geometry" IX="0">',
		`${THEMED_CELLS}<Section N="Geometry" IX="0">`,
	);
	xml = xml.replace(
		'<Text>',
		'<Section N="Character"><Row IX="0"><Cell N="Font" V="Themed" F="THEMEVAL()"/></Row></Section><Text>',
	);
	// The line (shape 2) takes its weight and its quick style matrix from the connector scheme.
	xml = xml.replace(/(<Shape ID="2"[^>]*>)/, '$1<Cell N="LineWeight" V="Themed" F="THEMEVAL()"/>');
	xml = xml.replace(/(<Cell N="QuickStyleLineMatrix" V=")1"/, '$14"');
	zip.file('visio/pages/page1.xml', xml);
	return zip.generateAsync({ type: 'uint8array' });
}
async function pageSheet(bytes: Uint8Array): Promise<Map<string, string | undefined>> {
	const root = await (await VisioPackage.open(bytes)).readXml('visio/pages/pages.xml');
	const sheet = child(children(root, 'Page')[0], 'PageSheet');
	return new Map(
		children(sheet, 'Cell').map((cell) => [attribute(cell, 'N')!, attribute(cell, 'V')]),
	);
}
const apply = async (bytes: Uint8Array, edit: Record<string, string>) =>
	(await editVsdx(bytes, [{ type: 'set-page-theme', pageId: page, ...edit } as never])).bytes;
const id = (name: Parameters<typeof visioBuiltInTheme>[0]) =>
	String(visioBuiltInTheme(name).schemeId);
const themeParts = async (bytes: Uint8Array) =>
	(await VisioPackage.open(bytes)).paths().filter((path) => path.startsWith('visio/theme/'));

describe('theme fonts', () => {
	it('writes only FontSchemeIndex and draws themed text in the theme body font', async () => {
		const original = await drawing();
		const before = (await parseVsdx(original)).pages[0]!.shapes[0]!.text;
		expect(visioFontAssumed(before)).toBe(true);
		const saved = await apply(original, { fonts: 'harbor' });
		const cells = await pageSheet(saved);
		expect(cells.get('FontSchemeIndex')).toBe(id('harbor'));
		for (const name of ['ColorSchemeIndex', 'EffectSchemeIndex', 'ConnectorSchemeIndex'])
			expect(cells.has(name) ? cells.get(name) : undefined).not.toBe(id('harbor'));
		const document = await parseVsdx(saved);
		const text = document.pages[0]!.shapes[0]!.text;
		expect(text.fontFamily).toBe('Trebuchet MS');
		expect(visioFontAssumed(text)).toBe(false);
		expect(document.pages[0]!.themeSchemes).toEqual({ fonts: 'harbor' });
		expect(document.pages[0]!.theme).toBeUndefined();
		assertViewableDocument(copySnapshotScene(document));
	});
	it('changes fonts and colours independently and survives another edit', async () => {
		let bytes = await apply(await drawing(), { fonts: 'harbor' });
		bytes = await apply(bytes, { colors: 'ember' });
		bytes = await apply(bytes, { fonts: 'slate' });
		const cells = await pageSheet(bytes);
		expect(cells.get('FontSchemeIndex')).toBe(id('slate'));
		expect(cells.get('ColorSchemeIndex')).toBe(id('ember'));
		const document = await parseVsdx(bytes);
		expect(document.pages[0]!.shapes[0]!.text.fontFamily).toBe('Segoe UI');
		expect(document.pages[0]!.themeSchemes).toEqual({ colors: 'ember', fonts: 'slate' });
		expect(await themeParts(bytes)).toHaveLength(3);
		bytes = await apply(bytes, { fonts: 'harbor' });
		expect(await themeParts(bytes)).toHaveLength(3);
		expect((await parseVsdx(bytes)).pages[0]!.shapes[0]!.text.fontFamily).toBe('Trebuchet MS');
	});
	it('replaces a built-in part written before font schemes existed instead of adding a second', async () => {
		const colored = await apply(await drawing(), { colors: 'harbor' });
		const zip = await JSZip.loadAsync(colored);
		const [path] = await themeParts(colored);
		const old = (await zip.file(path!)!.async('string')).replace(
			/(<\/a:minorFont>)<a:extLst>.*?<\/a:extLst>/s,
			'$1',
		);
		expect(old).not.toBe(await zip.file(path!)!.async('string'));
		zip.file(path!, old);
		const legacy = await zip.generateAsync({ type: 'uint8array' });
		const saved = await apply(legacy, { fonts: 'harbor' });
		expect(await themeParts(saved)).toEqual([path]);
		const document = await parseVsdx(saved);
		expect(document.pages[0]!.shapes[0]!.text.fontFamily).toBe('Trebuchet MS');
		expect(document.pages[0]!.theme).toMatchObject({ builtIn: 'harbor' });
	});
});

describe('theme effects', () => {
	const shape = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
	it('writes only EffectSchemeIndex; each effect scheme draws its own effect', async () => {
		const original = await drawing();
		const flat = (await shape(original)).style;
		expect(flat.shadow ?? flat.glow ?? flat.softEdges ?? flat.reflection).toBeUndefined();
		const slate = await apply(original, { effects: 'slate' });
		const cells = await pageSheet(slate);
		expect(cells.get('EffectSchemeIndex')).toBe(id('slate'));
		expect(cells.get('ColorSchemeIndex')).toBeUndefined();
		const shadow = (await shape(slate)).style;
		expect(shadow.shadow).toMatchObject({ color: '#000000' });
		expect(shadow.shadow!.opacity).toBeCloseTo(0.44, 2);
		expect(shadow.shadow!.offsetX).toBeGreaterThan(0);
		expect(shadow.shadow!.offsetY).toBeLessThan(0);
		expect(shadow.shadow!.blur).toBeGreaterThan(0);
		expect(shadow.glow ?? shadow.softEdges ?? shadow.reflection).toBeUndefined();
		const ember = (await shape(await apply(original, { effects: 'ember' }))).style;
		expect(ember.glow).toMatchObject({ color: '#d9480f' });
		expect(ember.glow!.size).toBeCloseTo((3 * 63500) / 914400, 4);
		expect(ember.shadow).toBeUndefined();
		const grove = (await shape(await apply(original, { effects: 'grove' }))).style;
		expect(grove.softEdges).toBeCloseTo((3 * 25400) / 914400, 4);
		const harbor = (await shape(await apply(original, { effects: 'harbor' }))).style;
		expect(harbor.reflection).toMatchObject({ opacity: 0.4, size: 0.45 });
		const office = (await shape(await apply(original, { effects: 'office' }))).style;
		expect(office.shadow ?? office.glow ?? office.softEdges ?? office.reflection).toBeUndefined();
	});
	it('does not report the rendered theme effects as unsupported', async () => {
		const saved = await apply(await drawing(), { effects: 'slate' });
		const codes = (await parseVsdx(saved)).diagnostics.map((item) => item.code);
		expect(codes).not.toContain('unsupported-theme-shadow');
		expect(codes).not.toContain('unsupported-theme-effects');
	});
});

describe('theme connectors', () => {
	const line = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!.shapes[1]!;
	it('writes only ConnectorSchemeIndex; the line follows the theme connector style', async () => {
		let bytes = await apply(await drawing(), { connectors: 'harbor' });
		const cells = await pageSheet(bytes);
		expect(cells.get('ConnectorSchemeIndex')).toBe(id('harbor'));
		expect(cells.get('EffectSchemeIndex')).toBeUndefined();
		expect((await line(bytes)).style.lineWidth).toBeCloseTo(2.25 / 72, 5);
		bytes = await apply(bytes, { connectors: 'office' });
		expect((await line(bytes)).style.lineWidth).toBeCloseTo(1 / 72, 5);
		expect((await parseVsdx(bytes)).pages[0]!.themeSchemes).toEqual({ connectors: 'office' });
	});
});

describe('page theme scheme commands', () => {
	it('takes one scheme part at a time', () => {
		for (const part of ['fonts', 'effects', 'connectors', 'colors'])
			expect(
				snapshotPageTheme({ type: 'set-page-theme', pageId: page, [part]: 'grove' } as never),
			).toMatchObject({ [part]: 'grove' });
		for (const edit of [
			{ fonts: 'x' },
			{ effects: 'grove', fonts: 'grove' },
			{ theme: 'office', connectors: 'grove' },
		])
			expect(() =>
				snapshotPageTheme({ type: 'set-page-theme', pageId: page, ...edit } as never),
			).toThrow();
	});
});
