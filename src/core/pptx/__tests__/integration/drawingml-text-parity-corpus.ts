/**
 * Shared corpus walk of the DrawingML text parity harnesses
 * (`drawingml-paragraph-neutral-parity`, `drawingml-run-neutral-parity`,
 * `drawingml-body-neutral-parity`): every XML part of every committed deck,
 * with each DOM element of interest paired with its object-tree node by path
 * (same-name sibling index).
 */
import JSZip from 'jszip';
import { readFileSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';

import { NS, elements, parseXml, type XmlElement } from '../../../xml/index';
import { PptxRuntimeDependencyFactory } from '../../core/core/factories/PptxRuntimeDependencyFactory';
import type { XmlObject } from '../../core/types';

const CORE = path.resolve(__dirname, '../../..');
const parser = new PptxRuntimeDependencyFactory().createParser();

/** Every committed deck under `pptx/__tests__/fixtures`, including the `e2e` snapshot. */
export function parityDecks(): string[] {
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

/** The object-tree child for a DOM child: same-name siblings are an array, a bare element `''`. */
function objectChild(node: unknown, name: string, index: number): unknown {
	if (!node || typeof node !== 'object') return undefined;
	const value = (node as XmlObject)[name];
	return Array.isArray(value) ? value[index] : index === 0 ? value : undefined;
}

/** JSON with object keys sorted, for order-insensitive comparison. */
export function canonicalJson(value: unknown): string {
	return JSON.stringify(value ?? null, (_key, entry: unknown) =>
		entry && typeof entry === 'object' && !Array.isArray(entry)
			? Object.fromEntries(
					Object.entries(entry as Record<string, unknown>).sort(([a], [b]) => a.localeCompare(b)),
				)
			: entry,
	);
}

export interface ParityVisit {
	label: string;
	element: XmlElement;
	node: XmlObject;
}

/**
 * Walks every part whose text matches `filter` and calls `visit` for each `a:` element whose local
 * name is in `locals`, with its object-tree node. Returns the number of parts compared.
 */
export async function walkParityCorpus(
	filter: RegExp,
	locals: ReadonlySet<string>,
	visit: (entry: ParityVisit) => void,
): Promise<number> {
	let parts = 0;
	for (const deck of parityDecks()) {
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(readFileSync(deck));
		} catch {
			continue; // The password-protected fixture is a CFB container, not a zip.
		}
		for (const name of Object.keys(zip.files).sort()) {
			if (zip.files[name]!.dir || !name.endsWith('.xml')) continue;
			const xml = await zip.file(name)!.async('string');
			if (!filter.test(xml)) continue;
			const label = `${path.relative(CORE, deck)}!${name}`;
			let root: XmlElement;
			try {
				root = parseXml(xml, { label }).documentElement;
			} catch {
				continue;
			}
			const tree = parser.parse(xml) as XmlObject;
			parts++;
			const walk = (element: XmlElement, node: unknown, trail: string) => {
				if (node === undefined) return;
				if (element.namespaceURI === NS.a && locals.has(element.localName)) {
					const objectNode = (typeof node === 'object' ? node : {}) as XmlObject;
					visit({ label: `${label} ${trail}`, element, node: objectNode });
					return;
				}
				const seen = new Map<string, number>();
				for (const child of elements(element)) {
					const index = seen.get(child.tagName) ?? 0;
					seen.set(child.tagName, index + 1);
					walk(
						child,
						objectChild(node, child.tagName, index),
						`${trail}/${child.tagName}#${index}`,
					);
				}
			};
			walk(root, tree[root.tagName], root.tagName);
		}
	}
	return parts;
}
