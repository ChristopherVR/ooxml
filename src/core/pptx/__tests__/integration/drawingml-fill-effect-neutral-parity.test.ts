import JSZip from 'jszip';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';

import { resolveDrawingColor } from '../../../drawingml/drawing-color';
import { readDrawingGradient, type DrawingGradientRead } from '../../../drawingml/gradient-read';
import { DOM_DRAWING_VIEW } from '../../../drawingml/node-view';
import type { DrawingColor } from '../../../drawingml/types';
import { NS, elements, parseXml, type XmlElement } from '../../../xml/index';
import { OBJECT_DRAWING_VIEW } from '../../core/color/object-drawing-view';
import { PptxRuntimeDependencyFactory } from '../../core/core/factories/PptxRuntimeDependencyFactory';
import type { XmlObject } from '../../core/types';

/**
 * DrawingML fill parity (docs/agnostic-core-plan.md step 3): the pptx gradient codec
 * (`PptxGradientStyleCodec`) reads `a:gradFill` through the shared reader
 * (`drawingml/gradient-read.ts`) over the object-tree view, and the DOM fill reader runs the same
 * reader over the DOM view. This harness reads every `a:gradFill` of every committed deck through
 * both views, paired by element path, and requires the same read (stops with their resolved
 * colours, angle, scaling, path, focus and tile rectangles, flip and rotation).
 *
 * One view difference is by design: the object view does not return an empty element as a child
 * (fast-xml-parser gives `<a:tileRect/>` as `''`; the codec always read it as absent and writes it
 * back verbatim), so an attribute-less `a:tileRect` or `a:path` is compared as absent.
 *
 * Before the switch, the previous codec (its 13 readers and the standalone tile reader) and the
 * previous DOM gradient reader were compared with the new ones over the 1,542 `a:gradFill`
 * elements of the corpus: identical.
 */

const CORE = path.resolve(__dirname, '../../..');
const parser = new PptxRuntimeDependencyFactory().createParser();
const theme = { scheme: (name: string) => (name ? '#4472C4' : undefined) };
const hex = (color: DrawingColor | undefined) =>
	color ? resolveDrawingColor(color, theme)?.hex : undefined;

function decks(): string[] {
	const found: string[] = [];
	const walk = (dir: string) => {
		for (const name of readdirSync(dir).sort()) {
			const full = path.join(dir, name);
			if (statSync(full).isDirectory()) walk(full);
			else if (/\.(pptx|pptm|potx|ppsx)$/i.test(name)) found.push(full);
		}
	};
	walk(path.join(CORE, 'pptx/__tests__/fixtures'));
	return found;
}

function objectChild(node: unknown, name: string, index: number): unknown {
	if (!node || typeof node !== 'object') return undefined;
	const value = (node as XmlObject)[name];
	return Array.isArray(value) ? value[index] : index === 0 ? value : undefined;
}

const emptyElement = (element: XmlElement | null | undefined) =>
	Boolean(element) && element!.attributes.length === 0 && elements(element!).length === 0;

function gradientSummary<N>(read: DrawingGradientRead<N>, dropTile: boolean, dropPath: boolean) {
	const { stops: _stops, ...fill } = read.fill;
	return {
		fill: {
			...fill,
			...(dropTile ? { tileRect: undefined } : {}),
			...(dropPath ? { path: undefined, fillToRect: undefined } : {}),
		},
		stops: read.stops.map((stop) => ({ position: stop.position, color: hex(stop.color) })),
		neutralStops: read.fill.stops.map((stop) => [stop.position, hex(stop.color)]),
		hasPath: dropPath ? false : read.hasPath,
		flip: read.flip,
		rotateWithShape: read.rotateWithShape,
		fillToRectSides: dropPath ? undefined : read.fillToRectSides,
		tileRectSides: dropTile ? undefined : read.tileRectSides,
	};
}

interface Difference {
	label: string;
	objectTree: string;
	dom: string;
}

const differences: Difference[] = [];
let gradientsCompared = 0;
let partsCompared = 0;

function compare(label: string, xml: string): void {
	let root: XmlElement;
	try {
		root = parseXml(xml, { label }).documentElement;
	} catch {
		return;
	}
	const tree = parser.parse(xml) as XmlObject;
	partsCompared++;
	const visit = (element: XmlElement, node: unknown, trail: string) => {
		if (node === undefined) return;
		if (element.namespaceURI === NS.a && element.localName === 'gradFill') {
			if (!node || typeof node !== 'object') return;
			gradientsCompared++;
			const dropTile = emptyElement(DOM_DRAWING_VIEW.child(element, 'tileRect'));
			const dropPath = emptyElement(DOM_DRAWING_VIEW.child(element, 'path'));
			const objectTree = JSON.stringify(
				gradientSummary(
					readDrawingGradient(OBJECT_DRAWING_VIEW, node as XmlObject),
					dropTile,
					dropPath,
				),
			);
			const dom = JSON.stringify(
				gradientSummary(readDrawingGradient(DOM_DRAWING_VIEW, element), dropTile, dropPath),
			);
			if (objectTree !== dom) differences.push({ label: `${label} ${trail}`, objectTree, dom });
			return;
		}
		const seen = new Map<string, number>();
		for (const child of elements(element)) {
			const index = seen.get(child.tagName) ?? 0;
			seen.set(child.tagName, index + 1);
			visit(child, objectChild(node, child.tagName, index), `${trail}/${child.tagName}#${index}`);
		}
	};
	visit(root, tree[root.tagName], root.tagName);
}

beforeAll(async () => {
	for (const deck of decks()) {
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(readFileSync(deck));
		} catch {
			continue; // The password-protected fixture is a CFB container, not a zip.
		}
		for (const name of Object.keys(zip.files).sort()) {
			if (zip.files[name]!.dir || !name.endsWith('.xml')) continue;
			const xml = await zip.file(name)!.async('string');
			if (!xml.includes('gradFill')) continue;
			compare(`${path.relative(CORE, deck)}!${name}`, xml);
		}
	}
}, 600_000);

describe('DrawingML fills: pptx object tree vs shared DOM reader', () => {
	it('compares a real corpus', () => {
		expect(partsCompared).toBeGreaterThan(300);
		expect(gradientsCompared).toBeGreaterThan(1_400);
	});

	it('reads every gradient the same way', () => {
		expect(differences.slice(0, 20)).toEqual([]);
	});
});
