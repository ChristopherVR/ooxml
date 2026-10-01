import { NS, elements, parseXml } from '../xml/index.js';

export interface ContentTypes {
	/** Extension (lower case, no dot) to content type. */
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
			if (extension) defaults.set(extension.toLowerCase(), contentType);
		} else if (node.localName === 'Override') {
			const partName = node.getAttribute('PartName');
			if (partName) overrides.set(partName, contentType);
		}
	}
	return { defaults, overrides };
}

/** The content type of a part: its override if declared, else the default for its extension. */
export function contentTypeForPart(types: ContentTypes, partName: string): string | undefined {
	const overridePath = partName.startsWith('/') ? partName : `/${partName}`;
	const override = types.overrides.get(overridePath);
	if (override) return override;
	const extension = partName.split('.').pop()?.toLowerCase();
	return extension ? types.defaults.get(extension) : undefined;
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
	if (types.defaults.has(extension.toLowerCase())) return xml ?? buildContentTypesXml(types);
	types.defaults.set(extension.toLowerCase(), contentType);
	return buildContentTypesXml(types);
}
