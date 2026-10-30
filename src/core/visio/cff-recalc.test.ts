import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { recalculateMasterInstances, type CffMasters } from './cff-recalc';
import { attribute, children } from './sheet';

/** The diagram Visio 16 built from its Cross-Functional Flowchart template (see the .md). */
async function load() {
	const zip = await JSZip.loadAsync(
		readFileSync(new URL('./__fixtures__/cross-functional-native.vsdx', import.meta.url)),
	);
	const text = async (path: string) => await zip.file(path)!.async('string');
	const masterList = parseXml(await text('visio/masters/masters.xml')).documentElement!;
	const rels = await text('visio/masters/_rels/masters.xml.rels');
	const shapes = new Map<string, Element>();
	for (const master of children(masterList, 'Master')) {
		const rid = attribute(children(master, 'Rel')[0], 'r:id')!;
		const target = new RegExp(
			`Id="${rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${rid}"`,
		).exec(rels)!;
		const part = parseXml(await text(`visio/masters/${target[1] ?? target[2]}`)).documentElement!;
		shapes.set(attribute(master, 'ID')!, children(children(part, 'Shapes')[0], 'Shape')[0]!);
	}
	const page = parseXml(await text('visio/pages/page1.xml')).documentElement!;
	const masters: CffMasters = { get: (id) => shapes.get(id) };
	return { page, masters };
}
const snapshot = (root: Element) =>
	[...root.getElementsByTagName('Cell')].map(
		(cell) =>
			`${attribute(cell.parentNode as Element, 'ID') ?? attribute((cell.parentNode!.parentNode as Element) ?? undefined, 'ID')}|${attribute(cell, 'N')}|${attribute(cell, 'V')}`,
	);

const shapeById = (root: Element, id: string) =>
	[...root.getElementsByTagName('Shape')].find((shape) => attribute(shape, 'ID') === id)!;
const setLocal = (shape: Element, name: string, value: number) =>
	children(shape, 'Cell')
		.find((cell) => attribute(cell, 'N') === name)!
		.setAttribute('V', String(value));

describe('recalculating the instances Visio saved', () => {
	it('reproduces the caches Visio wrote', async () => {
		const { page, masters } = await load();
		const before = snapshot(page);
		const written = recalculateMasterInstances(page, masters);
		const after = snapshot(page);
		const changed = before.map((value, index) => [value, after[index]]).filter(([a, b]) => a !== b);
		expect(written).toBe(0);
	});

	it('follows a resize through the master formulas', async () => {
		const { page, masters } = await load();
		const before = snapshot(page);
		setLocal(shapeById(page, '6'), 'Width', 12);
		setLocal(shapeById(page, '1'), 'Height', 4);
		const written = recalculateMasterInstances(page, masters);
		const after = snapshot(page);
		const changed = before.map((value, index) => [value, after[index]]).filter(([a, b]) => a !== b);
		expect(written).toBeGreaterThan(10);
		const now = new Map(
			changed.map(([, b]) => [b!.split('|').slice(0, 2).join('|'), b!.split('|')[2]]),
		);
		expect(now.get('7|Width')).toBe('12');
		expect(Number(now.get('3|PinY'))).toBeCloseTo(4 - 0.19685, 4);
	});
});
