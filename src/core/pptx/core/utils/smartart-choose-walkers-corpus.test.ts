/**
 * Corpus parity for the `dgm:choose` walkers: every layout definition in the
 * committed decks (`__tests__/fixtures`, including `smartart-gallery` and the
 * corpus) and every built-in layout, every layout node, a range of node
 * counts and contexts, through every walker that reads a branch's raw XML.
 * The digest per layout was recorded with the `XmlObject` walkers before they
 * were ported to the `diagram` ordered-XML walkers; it must never change.
 */
import { createHash } from 'node:crypto';
import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import type { PptxSmartArtLayoutDefinition, PptxSmartArtLayoutNode } from '../types';
import {
	listBuiltinSmartArtLayouts,
	loadBuiltinSmartArtLayoutXml,
	parseBuiltinLayoutDefinition,
} from './smartart-builtin-layouts';
import { tailedHierarchyDeclaresChAlign } from './smartart-hierarchy-tailed-transpose';
import { resolveItemTxAnchor } from './smartart-layout-item-tx-anchor';
import {
	chooseAlgorithm,
	chooseAlgorithmOfType,
	chooseAlgType,
} from './smartart-layout-interpreter-choose-algorithm';
import {
	structuralChooseAlgDepth,
	tunnelsPastOwnCompositeSlot,
} from './smartart-layout-interpreter-choose-depth';
import { arrangerRepeatsChildTemplate } from './smartart-layout-interpreter-hub-detect';
import { isColumnWrapper } from './smartart-layout-interpreter-item-role-orientation';
import { discoverArrangement } from './smartart-layout-interpreter-model';
import type { WhenContext } from '../../../diagram/layout/smartart-layout-interpreter-when';

const FIXTURES = join(__dirname, '../../__tests__/fixtures');
const COUNTS = [0, 1, 2, 3, 4, 5, 7, 9];
const TX = new Set(['tx']);
const SP = new Set(['sp', 'composite']);

function decks(dir: string): string[] {
	const out: string[] = [];
	for (const entry of readdirSync(dir)) {
		const path = join(dir, entry);
		if (statSync(path).isDirectory()) {
			out.push(...decks(path));
		} else if (entry.endsWith('.pptx')) {
			out.push(path);
		}
	}
	return out.sort();
}

async function layoutParts(): Promise<Map<string, string>> {
	const parts = new Map<string, string>();
	for (const entry of listBuiltinSmartArtLayouts()) {
		parts.set(`builtin:${entry.id}`, loadBuiltinSmartArtLayoutXml(entry.id));
	}
	for (const deck of decks(FIXTURES)) {
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(readFileSync(deck));
		} catch {
			continue;
		}
		for (const name of Object.keys(zip.files).sort()) {
			if (/^ppt\/diagrams\/layout[^/]*\.xml$/u.test(name)) {
				const text = await zip.file(name)!.async('string');
				parts.set(createHash('sha1').update(text).digest('hex'), text);
			}
		}
	}
	return parts;
}

function walk(node: PptxSmartArtLayoutNode, out: PptxSmartArtLayoutNode[]): void {
	out.push(node);
	node.children?.forEach((child) => walk(child, out));
}

function contexts(): WhenContext[] {
	return [
		{},
		{ presLayoutVars: {} },
		{ presLayoutVars: { direction: 'rev' } as never },
		{ position: 1, total: 3, depth: 1, maxDepth: 2 },
		{ position: 2, total: 2, depth: 2, maxDepth: 3 },
	];
}

function results(definition: PptxSmartArtLayoutDefinition): unknown[] {
	const nodes: PptxSmartArtLayoutNode[] = [];
	walk(definition.rootNode, nodes);
	const out: unknown[] = [];
	for (const count of COUNTS) {
		out.push(['arr', count, discoverArrangement(definition, count)?.kind]);
		out.push(['arrVars', count, discoverArrangement(definition, count, {} as never)?.node.name]);
		nodes.forEach((node, index) => {
			out.push(['tx', index, count, resolveItemTxAnchor(node, count, count % 2 === 0)]);
			out.push(['col', index, count, isColumnWrapper(node, count, undefined)]);
			out.push(['colVars', index, count, isColumnWrapper(node, count, {} as never)]);
			if (!node.choose?.length) {
				return;
			}
			for (const context of contexts()) {
				out.push([
					index,
					count,
					chooseAlgType(node, count, context),
					chooseAlgorithm(node, count, context),
					chooseAlgorithmOfType(node, count, TX, context),
					chooseAlgorithmOfType(node, count, SP, context),
					structuralChooseAlgDepth(node, count, context),
					tunnelsPastOwnCompositeSlot(node, count, context, new Set()),
				]);
			}
		});
	}
	nodes.forEach((node, index) => {
		out.push(['chAlign', index, tailedHierarchyDeclaresChAlign(node)]);
		out.push(['hub', index, arrangerRepeatsChildTemplate(node)]);
	});
	return out;
}

describe('dgm:choose walkers over the committed layout definitions', () => {
	it('match the recorded digests', { timeout: 600_000 }, async () => {
		const digests: Record<string, string> = {};
		const full: Record<string, unknown> = {};
		for (const [key, xml] of await layoutParts()) {
			const definition = parseBuiltinLayoutDefinition(xml);
			if (!definition) {
				continue;
			}
			const value = results(definition);
			const json = JSON.stringify(value);
			digests[key] = createHash('sha1').update(json).digest('hex');
			full[key] = value;
		}
		if (process.env.SMARTART_WALKER_DUMP) {
			writeFileSync(process.env.SMARTART_WALKER_DUMP, JSON.stringify(full, null, 1));
		}
		expect(Object.keys(digests).length).toBeGreaterThan(100);
		await expect(JSON.stringify(digests, null, '\t')).toMatchFileSnapshot(
			'./__snapshots__/smartart-choose-walkers-corpus.json',
		);
	});
});
