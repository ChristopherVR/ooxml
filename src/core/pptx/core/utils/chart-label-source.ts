import { XMLParser } from 'fast-xml-parser';
import type JSZip from 'jszip';

import type { IPptxXmlLookupService } from '../services';
import type { PptxExternalData, PptxChartSeries, XmlObject } from '../types';
import { collectChartWorkbookWrite, writeChartWorkbookUpdates } from './chart-xlsx-writer';
import { xmlAttr, xmlChild, xmlChildren } from './xml-access';
import { parseChartFormulaRange, expandRangeAddresses } from './chart-xlsx-cellref';

function formulaCount(value: unknown): number {
	if (!value || typeof value !== 'object') return 0;
	if (Array.isArray(value)) return value.reduce((total, item) => total + formulaCount(item), 0);
	return Object.entries(value).reduce(
		(total, [key, child]) => total + (key === 'f' || key.endsWith(':f') ? 1 : formulaCount(child)),
		0,
	);
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
		const entry = series[index];
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
	let references = 0;
	const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
	try {
		for (const path of Object.keys(zip.files)) {
			if (!path.endsWith('.rels') || !path.includes('/_rels/')) continue;
			const source = path.replace('/_rels/', '/').slice(0, -5);
			const tree = parser.parse(await zip.file(path)!.async('string')) as XmlObject;
			for (const rel of xmlChildren(xmlChild(tree, 'Relationships'), 'Relationship')) {
				const relationshipTarget = xmlAttr(rel, 'Target');
				if (
					relationshipTarget &&
					xmlAttr(rel, 'TargetMode') !== 'External' &&
					resolve(source, relationshipTarget) === target
				)
					references++;
			}
		}
		if (references !== 1) return false;
		await writeChartWorkbookUpdates(external.embeddedWorkbookData, writes, {
			atomic: true,
			requireCurrentValues: true,
		});
		return true;
	} catch {
		return false;
	}
}
