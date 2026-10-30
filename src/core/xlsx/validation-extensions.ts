import { NS, buildXml, children, elements, first, parseXml } from '../xml/index';
import type { Worksheet } from './model';
import { readDataValidation } from './read/data-validations';
import { dataValidationXml, usesValidationExtension } from './write/data-validations';
import { inlineFragment } from './write/xml-out';

const URI = '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}';

/** Move editable x14 rules into the shared validation model, retaining unrelated extensions. */
export function readValidationExtensions(sheet: Worksheet): void {
	const kept: string[] = [];
	for (const xml of sheet.preserved.get('extLst') ?? []) {
		const doc = parseXml(xml);
		let changed = false;
		for (const ext of children(doc.documentElement, 'ext', NS.x)) {
			if (ext.getAttribute('uri')?.toUpperCase() !== URI.toUpperCase()) continue;
			const collection = first(ext, 'dataValidations', NS.x14);
			if (!collection) continue;
			for (const node of children(collection, 'dataValidation', NS.x14)) {
				const rule = readDataValidation(node);
				if (!rule) continue;
				sheet.dataValidations.push(rule);
				collection.removeChild(node);
				changed = true;
			}
			// Keep collection-level settings (disablePrompts, xWindow, yWindow) for regeneration.
			collection.setAttribute(
				'count',
				String(children(collection, 'dataValidation', NS.x14).length),
			);
		}
		if (!changed) kept.push(xml);
		else if (elements(doc.documentElement).length) kept.push(buildXml(doc));
	}
	if (kept.length) sheet.preserved.set('extLst', kept);
	else sheet.preserved.delete('extLst');
}

/** Add current x14 rules to the other worksheet extensions, never duplicating a base rule. */
export function validationExtensionsXml(
	sheet: Worksheet,
	extensions: string,
	ids = new Set<string>(),
): string {
	const rules = sheet.dataValidations.filter(
		(dv) => dv.ranges.length && usesValidationExtension(dv),
	);
	const doc = parseXml(extensions || `<extLst xmlns="${NS.x}"/>`);
	let ext = elements(doc.documentElement).find(
		(node) => node.getAttribute('uri')?.toUpperCase() === URI.toUpperCase(),
	);
	if (!rules.length) {
		const collection = first(ext, 'dataValidations', NS.x14);
		if (ext && collection && !elements(collection).length) {
			ext.removeChild(collection);
			if (!elements(ext).length) doc.documentElement.removeChild(ext);
		}
		return elements(doc.documentElement).length ? inlineFragment(buildXml(doc)) : '';
	}
	if (!ext) {
		ext = doc.createElementNS(NS.x, 'ext');
		ext.setAttribute('uri', URI);
		doc.documentElement.appendChild(ext);
	}
	let collection = first(ext, 'dataValidations', NS.x14);
	if (!collection) {
		collection = doc.createElementNS(NS.x14, 'x14:dataValidations');
		ext.appendChild(collection);
	}
	for (const rule of rules) {
		const node = doc.importNode(parseXml(dataValidationXml(rule, true)).documentElement, true);
		const id = node.getAttributeNS(NS.xr, 'uid');
		if (id && ids.has(id)) node.removeAttributeNS(NS.xr, 'uid');
		else if (id) ids.add(id);
		collection.appendChild(node);
	}
	collection.setAttribute('count', String(children(collection, 'dataValidation', NS.x14).length));
	return inlineFragment(buildXml(doc));
}
