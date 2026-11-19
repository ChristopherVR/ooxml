import { NS, elements, parseXml } from '../xml/index';
import { asciiLowerCase, partNameKey } from './part-names';

export interface ContentTypes {
	/** Extension (ASCII lower case, no dot) to content type. */
	defaults: Map<string, string>;
	/** Part name (leading slash) to content type. */
	overrides: Map<string, string>;
}

const escapeAttr = (value: string) =>
	value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

/** Parses `[Content_Types].xml`; an absent part gives empty maps. */
export function parseContentTypes(xml: string | undefined): ContentTypes {
	const defaults = new Map<string, string>();
	const overrides = new Map<string, string>();
	if (!xml) return { defaults, overrides };
	for (const node of elements(parseXml(xml).documentElement)) {
		const contentType = node.getAttribute('ContentType');
		if (!contentType) continue;
		if (node.localName === 'Default') {
			const extension = node.getAttribute('Extension');
			if (extension) defaults.set(asciiLowerCase(extension), contentType);
		} else if (node.localName === 'Override') {
			const partName = node.getAttribute('PartName');
			if (partName) overrides.set(partName, contentType);
		}
	}
	return { defaults, overrides };
}

/** The `Override` for a part. Part names match ASCII case-insensitively (ECMA-376 Part 2, 7.2.3.5). */
function overrideFor(types: ContentTypes, partName: string): string | undefined {
	const exact = types.overrides.get(partName.startsWith('/') ? partName : `/${partName}`);
	if (exact) return exact;
	const key = partNameKey(partName);
	for (const [name, contentType] of types.overrides) {
		if (partNameKey(name) === key) return contentType;
	}
	return undefined;
}

/** The content type of a part: its override if declared, else the default for its extension. */
export function contentTypeForPart(types: ContentTypes, partName: string): string | undefined {
	const override = overrideFor(types, partName);
	if (override) return override;
	const lastSegment = partName.slice(partName.lastIndexOf('/') + 1);
	const dot = lastSegment.lastIndexOf('.');
	return dot < 0 ? undefined : types.defaults.get(asciiLowerCase(lastSegment.slice(dot + 1)));
}

export function buildContentTypesXml(types: ContentTypes): string {
	const defaults = Array.from(types.defaults)
		.map(
			([extension, type]) =>
				`<Default Extension="${extension}" ContentType="${escapeAttr(type)}"/>`,
		)
		.join('');
	const overrides = Array.from(types.overrides)
		.map(
			([part, type]) =>
				`<Override PartName="${escapeAttr(part)}" ContentType="${escapeAttr(type)}"/>`,
		)
		.join('');
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${NS.contentTypes}">${defaults}${overrides}</Types>`;
}

/** Adds a `Default` extension mapping when absent, returning the (possibly unchanged) XML. */
export function ensureContentTypeDefault(
	xml: string | undefined,
	extension: string,
	contentType: string,
): string {
	const types = parseContentTypes(xml);
	const key = asciiLowerCase(extension);
	if (types.defaults.has(key)) return xml ?? buildContentTypesXml(types);
	types.defaults.set(key, contentType);
	return buildContentTypesXml(types);
}
