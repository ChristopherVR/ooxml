import { buildXml } from '../xml/index';
import { attribute, children, readSheet } from './sheet';
import { fail } from './package-common';
import type { VisioPackage } from './package';
import { related, visioXml } from './parts';
import { visioPageGeometryScale } from './page-scale';
import type { VisioClipboardSnapshot } from './clipboard-types';
import { executableCellFormula } from './cell-formula';
import { parseVisioFormula } from './formula';

/** Include resource definitions rather than trusting a host-provided fingerprint. */
export async function clipboardResources(
	pkg: VisioPackage,
): Promise<VisioClipboardSnapshot['resources']> {
	const path = await related(pkg, '', 'document');
	if (path !== 'visio/document.xml')
		fail('UNSUPPORTED_CLIPBOARD', 'Clipboard requires the conventional document part.');
	const document = (await visioXml(pkg, path, 'VisioDocument')).cloneNode(true) as Element;
	for (const node of Array.from(document.childNodes))
		if (
			node.nodeType !== 1 ||
			!['Colors', 'FaceNames', 'StyleSheets', 'DocumentSheet'].includes((node as Element).localName)
		)
			document.removeChild(node);
	const resources: { path: string; xml: string }[] = [{ path, xml: buildXml(document) }];
	const relationships = (await pkg.readXml('visio/_rels/document.xml.rels')).cloneNode(
		true,
	) as Element;
	for (const node of Array.from(relationships.childNodes))
		if (node.nodeType !== 1 || !(attribute(node as Element, 'Type') ?? '').endsWith('/theme'))
			relationships.removeChild(node);
	resources.push({ path: 'visio/_rels/document.xml.rels', xml: buildXml(relationships) });
	for (const name of pkg.paths().sort())
		if (/^visio\/theme\/(?:[^/]+\.xml|_rels\/[^/]+\.rels)$/.test(name))
			resources.push({ path: name, xml: buildXml(await pkg.readXml(name)) });
	return resources.sort((a, b) => a.path.localeCompare(b.path));
}
export async function clipboardPageContext(
	pkg: VisioPackage,
	pageId: string,
): Promise<{ xml: string; scale: number }> {
	const document = await related(pkg, '', 'document');
	const pagesPath = await related(pkg, document!, 'pages');
	const pages = children(await visioXml(pkg, pagesPath!, 'Pages'), 'Page');
	const matches = pages.filter((page) => attribute(page, 'ID') === pageId);
	if (matches.length !== 1) fail('EDIT_TARGET_NOT_FOUND', 'Clipboard page must exist uniquely.');
	const page = matches[0]!,
		sheets = children(page, 'PageSheet');
	if (sheets.length !== 1) fail('UNSUPPORTED_CLIPBOARD', 'One explicit PageSheet is required.');
	const scale = visioPageGeometryScale(readSheet(sheets[0]).cells, () =>
		fail('UNSUPPORTED_CLIPBOARD', 'Page drawing ratio is unresolved.'),
	);
	const context = page.ownerDocument!.createElementNS(page.namespaceURI, 'PageContext');
	context.setAttribute('PageNumber', String(pages.indexOf(page) + 1));
	context.setAttribute('PageCount', String(pages.length));
	for (const name of ['Background', 'BackPage']) {
		const value = attribute(page, name);
		if (value !== undefined) context.setAttribute(name, value);
	}
	context.appendChild(sheets[0]!.cloneNode(true));
	return { xml: buildXml(context), scale };
}
/** Every explicit font/style reference must resolve in actual target definitions. */
export function assertClipboardResourceReferences(root: Element, document: Element): void {
	const styles = new Map(
		children(children(document, 'StyleSheets')[0], 'StyleSheet').map((style) => [
			attribute(style, 'ID'),
			style,
		]),
	);
	const fonts = children(children(document, 'FaceNames')[0], 'FaceName');
	const fontIds = new Set(fonts.map((font, index) => attribute(font, 'ID') ?? String(index)));
	const fontNames = new Set(
		fonts.map((font) => attribute(font, 'Name') ?? attribute(font, 'NameU')),
	);
	const fontById = new Map(
		fonts.map((font, index) => [
			attribute(font, 'ID') ?? String(index),
			attribute(font, 'Name') ?? attribute(font, 'NameU'),
		]),
	);
	if (
		fontIds.size !== fonts.length ||
		fontNames.size !== fonts.length ||
		new Set([...fontNames].map((name) => name?.toLowerCase())).size !== fonts.length
	)
		fail(
			'UNSUPPORTED_CLIPBOARD',
			'Clipboard font definitions must have unique IDs and family names.',
		);
	const pool = [root, ...Array.from(root.getElementsByTagName('*'))],
		included = new Set<Element>();
	const validateStyle = (id: string, category: string, seen = new Set<string>()) => {
		if (seen.has(id) || seen.size >= 64)
			fail('UNSUPPORTED_CLIPBOARD', 'Cyclic clipboard style ancestry.');
		const style = styles.get(id);
		if (!style)
			fail('UNSUPPORTED_CLIPBOARD', 'Clipboard style reference does not exist in target.');
		if (!included.has(style)) {
			included.add(style);
			pool.push(style, ...Array.from(style.getElementsByTagName('*')));
		}
		seen.add(id);
		const parent = attribute(style, category);
		if (parent !== undefined && parent !== id) validateStyle(parent, category, seen);
	};
	for (const node of pool) {
		for (const attr of Array.from(node.attributes))
			if (
				attr.namespaceURI === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
			)
				fail('UNSUPPORTED_CLIPBOARD', 'Relationship-bearing clipboard fragments are unsupported.');
		for (const category of ['FillStyle', 'LineStyle', 'TextStyle']) {
			const id = attribute(node, category);
			if (id !== undefined) validateStyle(id, category);
		}
		if (node.localName === 'Cell' && ['Font', 'BulletFont'].includes(attribute(node, 'N') ?? '')) {
			const value = attribute(node, 'V');
			if (
				value !== undefined &&
				value !== 'Themed' &&
				attribute(node, 'F') !== 'Inh' &&
				!fontIds.has(value) &&
				!fontNames.has(value)
			)
				fail('UNSUPPORTED_CLIPBOARD', 'Clipboard font does not exist in target FaceNames.');
			const formula = executableCellFormula(attribute(node, 'F'));
			const ast = formula ? parseVisioFormula(formula) : undefined;
			if (ast?.kind === 'call' && ast.name === 'FONT') {
				const arg = ast.args[0];
				if (
					ast.args.length !== 1 ||
					arg?.kind !== 'string' ||
					!fontNames.has(arg.value) ||
					(value !== undefined && (fontById.get(value) ?? value) !== arg.value)
				)
					fail('UNSUPPORTED_CLIPBOARD', 'Clipboard font formula and existing font cache disagree.');
			}
		}
	}
}
