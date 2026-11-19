import { XMLParser } from 'fast-xml-parser';
import type JSZip from 'jszip';

import { deriveForLoad, readPartXml } from '../core/runtime/load-part-reader';
import type { IPptxXmlLookupService } from '../services';
import type { PptxExternalData, PptxChartSeries, XmlObject } from '../types';
import { collectChartWorkbookWrite, writeChartWorkbookUpdates } from './chart-xlsx-writer';
import { xmlAttr, xmlChild, xmlChildren } from './xml-access';
import { parseChartFormulaRange, expandRangeAddresses } from './chart-xlsx-cellref';
import { at } from '../../../geometry/indexed';

// Shared so that, during a load, each relationship part is parsed once for every chart.
const relationshipParser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });

function formulaCount(value: unknown): number {
	if (!value || typeof value !== 'object') return 0;
	if (Array.isArray(value)) return value.reduce((total, item) => total + formulaCount(item), 0);
	return Object.entries(value).reduce(
		(total, [key, child]) => total + (key === 'f' || key.endsWith(':f') ? 1 : formulaCount(child)),
		0,
	);
}

/** How many internal relationships, in any part, target each part of the package. */
async function relationshipTargetCounts(
	zip: JSZip,
	resolve: (base: string, target: string) => string,
): Promise<Map<string, number>> {
	const counts = new Map<string, number>();
	for (const path of Object.keys(zip.files)) {
		if (!path.endsWith('.rels') || !path.includes('/_rels/')) continue;
		const source = path.replace('/_rels/', '/').slice(0, -5);
		const tree = await readPartXml(zip, relationshipParser, path);
		for (const rel of xmlChildren(xmlChild(tree, 'Relationships'), 'Relationship')) {
			const relationshipTarget = xmlAttr(rel, 'Target');
			if (relationshipTarget && xmlAttr(rel, 'TargetMode') !== 'External') {
				const target = resolve(source, relationshipTarget);
				counts.set(target, (counts.get(target) ?? 0) + 1);
			}
		}
	}
	return counts;
}

/** Conservative load-time preflight for editing classic chart labels and their source cells. */
export async function chartLabelsHaveEditableSources(
	xml: IPptxXmlLookupService,
	zip: JSZip,
	chartSpace: XmlObject | undefined,
	seriesNodes: readonly XmlObject[],
	series: readonly PptxChartSeries[],
	categories: readonly string[],
	external: PptxExternalData | undefined,
	chartPartPath: string,
	resolve: (base: string, target: string) => string,
): Promise<boolean> {
	if (
		!chartSpace ||
		seriesNodes.length !== series.length ||
		!categories.length ||
		external?.isExternal
	)
		return false;
	// Other formula consumers (title, error bars, extension caches) need independent
	// reconciliation. Do not modify their source workbook through this workflow.
	if (formulaCount(chartSpace) !== seriesNodes.reduce((sum, node) => sum + formulaCount(node), 0))
		return false;
	const writes = [];
	const labelOwners = new Map<string, string>();
	for (const [index, node] of seriesNodes.entries()) {
		// `seriesNodes` and `series` have the same length (checked above).
		const entry = at(series, index);
		for (const [name, numeric, values] of [
			['tx', false, [entry.name]],
			['cat', false, [...categories]],
			['val', true, entry.values.map(String)],
		] as const) {
			const container = xml.getChildByLocalName(node, name);
			if (!container) return false;
			const write = collectChartWorkbookWrite(xml, container, numeric, [...values]);
			if (write) {
				writes.push(write);
				if (!numeric) {
					const range = parseChartFormulaRange(write.formula);
					const addresses = range ? expandRangeAddresses(range, values.length) : undefined;
					if (!range || !addresses) return false;
					for (const [point, address] of addresses.entries()) {
						const key = `${range.sheetName.toLowerCase()}!${address}`;
						const owner = name === 'tx' ? `series:${index}` : `category:${point}`;
						if (labelOwners.has(key) && labelOwners.get(key) !== owner) return false;
						labelOwners.set(key, owner);
					}
				}
			} else if (
				external ||
				(!xml.getChildByLocalName(container, numeric ? 'numLit' : 'strLit') &&
					!(name === 'tx' && xml.getScalarChildByLocalName(container, 'v') !== undefined))
			)
				return false;
		}
	}
	if (formulaCount(chartSpace) !== writes.length) return false;
	if (!external) return writes.length === 0;
	if (!external.targetPath || !external.embeddedWorkbookData) return false;
	const target = resolve(chartPartPath, external.targetPath);
	try {
		// The workbook must belong to this chart alone. Every chart of a load shares one count.
		const counts = await deriveForLoad(zip, 'chart-label-source:relationship-targets', () =>
			relationshipTargetCounts(zip, resolve),
		);
		if (counts.get(target) !== 1) return false;
		await writeChartWorkbookUpdates(external.embeddedWorkbookData, writes, {
			atomic: true,
			requireCurrentValues: true,
			dryRun: true,
		});
		return true;
	} catch {
		return false;
	}
}
