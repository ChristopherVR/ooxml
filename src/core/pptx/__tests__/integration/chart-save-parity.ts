import JSZip from 'jszip';
import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';

import { parseChartSpace, type ChartSpace } from '../../../chart/index';
import type { ChartFixture } from './chart-save-fixture';

export { xmlDifferences } from './chart-save-parity-canon';

/**
 * Support for the chart save harness (`chart-save-parity.test.ts`): load every committed deck that
 * holds a classic chart part, make one edit on every chart it applies to, and save it two ways:
 * through PowerPoint's save (`PptxHandler.save`, the object-tree writers plus whatever already goes
 * through `writeChartSpace`) and through the neutral reference (`parseChartSpace` of the part, the
 * equivalent `chart/edit` operation, `writeChartSpace`). Both results are compared as canonical
 * XML and as the `PptxChartData` they load back as.
 */

const ROOT = path.resolve(__dirname, '../fixtures');
const CHART_PART = /^ppt\/charts\/chart[^/]*\.xml$/u;

function walk(dir: string): string[] {
	return readdirSync(dir, { withFileTypes: true })
		.flatMap((entry) => {
			const full = path.join(dir, entry.name);
			if (entry.isDirectory()) return walk(full);
			return entry.name.endsWith('.pptx') ? [full] : [];
		})
		.sort();
}

export type SaveDeck = ChartFixture<ChartSpace>;

let classicDecks: Promise<readonly SaveDeck[]> | undefined;

/** Every committed deck with at least one classic (`c:chartSpace`) chart part. */
export function classicChartDecks(): Promise<readonly SaveDeck[]> {
	return (classicDecks ??= readClassicChartDecks());
}

async function readClassicChartDecks(): Promise<SaveDeck[]> {
	const decks: SaveDeck[] = [];
	for (const file of walk(ROOT)) {
		const bytes = readFileSync(file);
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(bytes);
		} catch {
			continue;
		}
		const parts: { part: string; space: ChartSpace }[] = [];
		for (const part of Object.keys(zip.files).filter((name) => CHART_PART.test(name))) {
			const xml = await zip.file(part)!.async('string');
			if (!/<c:chartSpace[\s>]/u.test(xml)) continue;
			const { chartSpace } = parseChartSpace(xml);
			parts.push({ part, space: chartSpace });
		}
		parts.sort((a, b) => a.part.localeCompare(b.part));
		if (parts.length > 0)
			decks.push({
				name: path.relative(ROOT, file).replace(/\\/gu, '/'),
				bytes,
				parts,
				archive: zip,
			});
	}
	return decks;
}

/**
 * What depends on the embedded workbook: PowerPoint's save writes edited series data into it, the
 * neutral reference writes the chart part only, so the two are not compared. That covers the
 * workbook contents and whether data labels may be edited through it (`dataLabelTextEditable`
 * and its `dataLabelSourceShape`, set only while every referenced cell agrees with the caches).
 */
const WORKBOOK_PATHS = new Set([
	'.embeddedWorkbookData',
	'.externalData.embeddedWorkbookData',
	'.dataLabelTextEditable',
	'.dataLabelSourceShape',
]);

/** Paths where two JSON values differ, with both values. */
export function modelDifferences(a: unknown, b: unknown, at = ''): string[] {
	if (WORKBOOK_PATHS.has(at) || JSON.stringify(a) === JSON.stringify(b)) return [];
	if (a && b && typeof a === 'object' && typeof b === 'object') {
		const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
		return [...keys].flatMap((key) =>
			modelDifferences(
				(a as Record<string, unknown>)[key],
				(b as Record<string, unknown>)[key],
				`${at}.${key}`,
			),
		);
	}
	return [`${at}: pptx ${JSON.stringify(a)}, neutral ${JSON.stringify(b)}`];
}
