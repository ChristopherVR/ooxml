import { NS, buildXml, children, elements, parseXml } from '../../xml/index';
import { CUSTOM_PROPERTIES_FMTID, type CustomProperty } from './types';

const escape = (text: string, attribute = false): string => {
	const out = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
	return attribute ? out.replace(/"/g, '&quot;') : out;
};

const isVt = (element: Element) => element.namespaceURI === NS.vt || element.prefix === 'vt';

function readValue(variant: Element): CustomProperty | undefined {
	const text = variant.textContent ?? '';
	switch (isVt(variant) ? variant.localName : '') {
		case 'lpwstr':
			return { name: '', type: 'lpwstr', value: text };
		case 'i4':
		case 'int': {
			const value = Number.parseInt(text, 10);
			return Number.isFinite(value) ? { name: '', type: 'i4', value } : undefined;
		}
		case 'r8': {
			const value = Number(text);
			return text.trim() !== '' && Number.isFinite(value)
				? { name: '', type: 'r8', value }
				: undefined;
		}
		case 'bool':
			return { name: '', type: 'bool', value: text.trim() === 'true' || text.trim() === '1' };
		case 'filetime':
			return { name: '', type: 'filetime', value: text };
		default:
			return undefined;
	}
}

/** Reads `docProps/custom.xml`; a property of an unmodelled variant type is kept as `raw`. */
export function parseCustomProperties(xml: string | undefined): CustomProperty[] {
	if (!xml) return [];
	const root = parseXml(xml, { label: 'custom properties' }).documentElement;
	const out: CustomProperty[] = [];
	for (const property of children(root, 'property', NS.customProperties)) {
		const name = property.getAttribute('name') ?? '';
		const variant = elements(property)[0];
		if (!name || !variant) continue;
		const value = readValue(variant) ?? {
			name,
			type: 'raw' as const,
			xml: buildXml(variant as unknown as Document),
		};
		const pid = Number.parseInt(property.getAttribute('pid') ?? '', 10);
		const linkTarget = property.getAttribute('linkTarget');
		out.push({
			...value,
			name,
			...(Number.isFinite(pid) ? { pid } : {}),
			...(linkTarget ? { linkTarget } : {}),
		});
	}
	return out;
}

/**
 * Property ids for `properties`: a valid, unique `pid` (2 and up) is kept, every other property
 * gets the next free id after the largest one in use.
 */
export function allocateCustomPropertyIds(properties: readonly CustomProperty[]): number[] {
	const used = new Set<number>();
	const kept = properties.map((property) => {
		const pid = property.pid;
		if (pid === undefined || !Number.isInteger(pid) || pid < 2 || used.has(pid)) return undefined;
		used.add(pid);
		return pid;
	});
	let next = Math.max(1, ...used) + 1;
	return kept.map((pid) => pid ?? next++);
}

function variantXml(property: CustomProperty): string {
	switch (property.type) {
		case 'raw':
			return property.xml;
		case 'lpwstr':
			return `<vt:lpwstr>${escape(property.value)}</vt:lpwstr>`;
		case 'i4':
			return `<vt:i4>${Math.trunc(property.value)}</vt:i4>`;
		case 'r8':
			return `<vt:r8>${String(property.value)}</vt:r8>`;
		case 'bool':
			return `<vt:bool>${property.value ? 'true' : 'false'}</vt:bool>`;
		case 'filetime':
			return `<vt:filetime>${escape(property.value)}</vt:filetime>`;
	}
}

/**
 * `docProps/custom.xml` for `properties`, or `undefined` when there are none (the part and its
 * relationship are then omitted). Names must be unique (case-insensitive, as Office requires):
 * a later duplicate is dropped.
 */
export function writeCustomProperties(properties: readonly CustomProperty[]): string | undefined {
	const seen = new Set<string>();
	const unique = properties.filter((property) => {
		const key = property.name.toLowerCase();
		if (!property.name || seen.has(key)) return false;
		seen.add(key);
		return true;
	});
	if (!unique.length) return undefined;
	const pids = allocateCustomPropertyIds(unique);
	const body = unique
		.map((property, index) => {
			const link = property.linkTarget ? ` linkTarget="${escape(property.linkTarget, true)}"` : '';
			return `<property fmtid="${CUSTOM_PROPERTIES_FMTID}" pid="${pids[index] ?? index + 2}" name="${escape(property.name, true)}"${link}>${variantXml(property)}</property>`;
		})
		.join('');
	return (
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
		`<Properties xmlns="${NS.customProperties}" xmlns:vt="${NS.vt}">${body}</Properties>`
	);
}

/**
 * A custom property seen as text, for editors that show every variant type as a string: the
 * variant element's local name (`lpwstr`, `i4`, `ui4`, `lpstr`...) and its text content.
 */
export interface CustomPropertyText {
	name: string;
	type: string;
	text: string;
	pid?: number;
}

/** Reads `docProps/custom.xml` as text; `type` is empty for a value outside the `vt` namespace. */
export function parseCustomPropertyTexts(xml: string | undefined): CustomPropertyText[] {
	if (!xml) return [];
	const root = parseXml(xml, { label: 'custom properties' }).documentElement;
	const out: CustomPropertyText[] = [];
	for (const property of children(root, 'property', NS.customProperties)) {
		const name = property.getAttribute('name') ?? '';
		if (!name) continue;
		const variant = elements(property)[0];
		const pid = Number.parseInt(property.getAttribute('pid') ?? '', 10);
		out.push({
			name,
			type: variant && isVt(variant) ? variant.localName : '',
			text: variant?.textContent ?? '',
			...(Number.isFinite(pid) ? { pid } : {}),
		});
	}
	return out;
}

const INT32 = /^-?\d+$/;

/**
 * The property `name` holding `text` as a `vt:<type>` value. It is typed when the text survives
 * the typed value unchanged (`i4` `42`, `r8` `2.5`, `bool` `true`), and raw otherwise (`i4`
 * `042`, `ui4`, `lpstr`...), so the written text is always exactly `text`. An unusable type
 * name falls back to `lpwstr`.
 */
export function customPropertyFromText(name: string, type: string, text: string): CustomProperty {
	switch (type) {
		case 'lpwstr':
		case 'filetime':
			return { name, type, value: text };
		case 'i4': {
			const value = Number.parseInt(text, 10);
			if (INT32.test(text) && String(value) === text && value === (value | 0))
				return { name, type, value };
			break;
		}
		case 'r8': {
			const value = Number(text);
			if (text.trim() !== '' && String(value) === text) return { name, type, value };
			break;
		}
		case 'bool':
			if (text === 'true' || text === 'false') return { name, type, value: text === 'true' };
			break;
		default:
			if (!/^[A-Za-z][A-Za-z0-9]*$/.test(type)) return { name, type: 'lpwstr', value: text };
	}
	return { name, type: 'raw', xml: `<vt:${type}>${escape(text)}</vt:${type}>` };
}
