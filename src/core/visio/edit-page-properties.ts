import { parseAppProperties, writeAppProperties } from '../opc/properties/index.js';
import { NS, parseXml } from '../xml/index.js';
import { attribute, children } from './sheet.js';
import { fail, type VisioPackageLimits } from './package-common.js';
import { serializeEditedXml } from './edit-text.js';
import type { VisioPackage } from './package.js';

/** Refresh the existing Pages category without replacing other property groups/extensions. */
export async function updatePageAppProperties(
	pkg: VisioPackage,
	pages: Element,
	priorCount: number,
	dirty: Map<string, Element>,
	limits: VisioPackageLimits,
	check: () => void,
): Promise<void> {
	const relationships = [...(await pkg.relationships('')).values()].filter(
		(rel) =>
			rel.type ===
			'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
	);
	if (!relationships.length) return;
	if (relationships.length !== 1 || relationships[0]!.mode !== 'Internal')
		fail(
			'EDIT_UNSUPPORTED_APP_PROPERTIES',
			'Extended properties require one internal relationship.',
		);
	const path = relationships[0]!.target;
	const root = await pkg.readXml(path, 'Properties');
	if (root.namespaceURI !== NS.extendedProperties)
		fail('EDIT_UNSUPPORTED_APP_PROPERTIES', 'Invalid extended property namespace.');
	const source = new TextDecoder().decode(serializeEditedXml(root, limits, check).bytes);
	const props = parseAppProperties(source);
	const pairs = props.headingPairs ?? [],
		titles = props.titlesOfParts ?? [];
	const pagePairs = pairs.filter((pair) => pair.name === 'Pages');
	if (!pagePairs.length) return;
	if (
		pagePairs.length !== 1 ||
		pagePairs[0]!.count !== priorCount ||
		pairs.some((pair) => !Number.isSafeInteger(pair.count) || pair.count < 0) ||
		pairs.reduce((total, pair) => total + pair.count, 0) !== titles.length
	)
		fail(
			'EDIT_UNSUPPORTED_APP_PROPERTIES',
			'Page property vectors have ambiguous category counts.',
		);
	const names = children(pages, 'Page').map(
		(page) =>
			attribute(page, 'Name') ?? attribute(page, 'NameU') ?? `Page ${attribute(page, 'ID')}`,
	);
	let offset = 0;
	const newTitles: string[] = [];
	for (const pair of pairs) {
		newTitles.push(...(pair.name === 'Pages' ? names : titles.slice(offset, offset + pair.count)));
		offset += pair.count;
	}
	dirty.set(
		path,
		parseXml(
			writeAppProperties(
				{
					...props,
					headingPairs: pairs.map((pair) =>
						pair.name === 'Pages' ? { ...pair, count: names.length } : pair,
					),
					titlesOfParts: newTitles,
				},
				source,
			),
		).documentElement,
	);
}
