import { children, elements, type XmlElement } from '../../xml/xml';
import type { AddInRequirement, OfficeAddInHost } from './model';
import {
	identifier,
	invalid,
	localeCandidates,
	localeTag,
	requirement,
	text,
	unique,
} from './validation';

const XSI = 'http://www.w3.org/2001/XMLSchema-instance';
export const BASE = 'http://schemas.microsoft.com/office/appforoffice/1.1';
export const OVERRIDE = 'http://schemas.microsoft.com/office/taskpaneappversionoverrides';
export const TYPES = 'http://schemas.microsoft.com/office/officeappbasictypes/1.0';

export const XML_HOSTS = new Map<string, OfficeAddInHost | null>([
	['Document', 'word'],
	['Workbook', 'excel'],
	['Presentation', 'powerpoint'],
	['Notebook', null],
]);

export const XML_BASE_HOSTS = new Map<string, OfficeAddInHost | null>([
	...XML_HOSTS,
	['Project', null],
	['Database', null],
	['Mailbox', null],
]);

export function all(
	parent: XmlElement | undefined,
	local: string,
	ns = parent?.namespaceURI ?? BASE,
): XmlElement[] {
	return parent ? children(parent, local, ns, false) : [];
}

export function one(
	parent: XmlElement | undefined,
	local: string,
	ns?: string,
): XmlElement | undefined {
	const matches = all(parent, local, ns);
	if (matches.length > 1) invalid(local);
	return matches[0];
}

export function typeName(element: XmlElement): string {
	const type = text(element.getAttributeNS(XSI, 'type'), 'xsi:type');
	const parts = type.split(':');
	if (parts.length === 1) return type;
	const [prefix, name] = parts;
	if (
		parts.length !== 2 ||
		!prefix ||
		!name ||
		element.lookupNamespaceURI(prefix) !== element.namespaceURI
	)
		return invalid('xsi:type namespace');
	return name;
}

export function value(parent: XmlElement | undefined, local: string): string | undefined {
	const node = one(parent, local);
	if (node && elements(node).length) invalid(local);
	return node?.textContent ?? undefined;
}

export function variants(
	element: XmlElement,
	convert: (value: unknown) => string = (v) => text(v, element.localName),
) {
	const defaultValue = convert(element.getAttribute('DefaultValue'));
	const byLocale = new Map<string, string>();
	for (const override of all(element, 'Override')) {
		const locale = localeTag(override.getAttribute('Locale'));
		if (byLocale.has(locale)) invalid('duplicate locale override');
		byLocale.set(locale, convert(override.getAttribute('Value')));
	}
	return (locale: string | undefined, defaultLocale: string): string => {
		for (const candidate of new Set([
			...localeCandidates(locale),
			...localeCandidates(defaultLocale),
		])) {
			const localized = byLocale.get(candidate);
			if (localized !== undefined) return localized;
		}
		return defaultValue;
	};
}

export function xmlRequirements(node: XmlElement | undefined): AddInRequirement[] {
	if (!node) return [];
	const ns = node.namespaceURI === OVERRIDE ? TYPES : BASE;
	const sets = one(node, 'Sets', ns);
	const methods = one(node, 'Methods', ns);
	const defaultMin = sets?.getAttribute('DefaultMinVersion') ?? undefined;
	const result: AddInRequirement[] = all(sets, 'Set').map((set) =>
		requirement(set.getAttribute('Name'), set.getAttribute('MinVersion') ?? defaultMin),
	);
	for (const method of all(methods, 'Method')) {
		result.push({
			kind: 'method',
			name: identifier(method.getAttribute('Name'), 'method'),
			minVersion: '1.1',
		});
	}
	unique(result, (item) => `${item.kind}:${item.name}`, 'duplicate requirement');
	return result;
}
