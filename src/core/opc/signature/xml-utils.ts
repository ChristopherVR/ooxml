// Moved from src/core/pptx/core/utils/signature-xml-utils.ts (see PROVENANCE.md).
/**
 * Pure regex-based XML extraction utilities for digital signature processing.
 *
 * These functions operate on raw XML strings without requiring a DOM parser,
 * making them platform-agnostic (browser + Node).
 */

/**
 * Escape special characters in an XML attribute value.
 * Handles `&`, `<`, `>`, `"`, and `'` so the result is safe inside both
 * double- and single-quoted attributes.
 */
export function escapeXmlAttr(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

/**
 * Escape special characters in XML text content.
 * Only `&`, `<`, and `>` need escaping outside of attribute values.
 */
export function escapeXmlText(value: string): string {
	return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Validate that `value` only contains characters from the standard base64 alphabet. */
export function isValidBase64(value: string): boolean {
	return typeof value === 'string' && value.length > 0 && /^[A-Za-z0-9+/=\s]+$/.test(value);
}

/**
 * Extract an attribute value from the first matching XML tag via regex.
 * Namespace prefixes on the tag name are supported via the pattern.
 */
export function extractTagAttribute(
	xml: string,
	tagName: string,
	attributeName: string,
): string | undefined {
	const pattern = new RegExp(
		`<${tagName}\\b[^>]*\\b${attributeName}="(?<attributeValue>[^"]+)"`,
		'i',
	);
	const match = xml.match(pattern);
	return match?.groups?.['attributeValue']?.trim();
}

const isNameChar = (c: string | undefined): boolean => c !== undefined && /[\w.-]/.test(c);
const isWordChar = (c: string | undefined): boolean => c !== undefined && /\w/.test(c);

/** Index just after `[prefix:]localName` starting at `from` (case-insensitive), or -1. */
function matchName(xml: string, lowerXml: string, from: number, lname: string): number {
	let pos = from;
	if (!lowerXml.startsWith(lname, pos)) {
		let end = pos;
		while (isNameChar(xml[end])) end++;
		if (end === pos || xml[end] !== ':') return -1;
		pos = end + 1;
		if (!lowerXml.startsWith(lname, pos)) return -1;
	}
	return pos + lname.length;
}

interface TagMatch {
	text: string;
	end: number;
}

/**
 * Linear scan equivalent of `<(p:)?name[^>]*>(...)</(p:)?name>` (lazy,
 * case-insensitive). Once a start tag has no closing tag after it, no later
 * start tag can have one either, so the scan stops (no quadratic backtracking).
 */
function findTagText(xml: string, lowerXml: string, lname: string, from: number): TagMatch | null {
	let lt = xml.indexOf('<', from);
	while (lt !== -1) {
		const afterName = matchName(xml, lowerXml, lt + 1, lname);
		if (afterName !== -1 && !isWordChar(xml[afterName])) {
			const gt = xml.indexOf('>', afterName);
			if (gt === -1) return null;
			let close = xml.indexOf('</', gt + 1);
			while (close !== -1) {
				const closeEnd = matchName(xml, lowerXml, close + 2, lname);
				if (closeEnd !== -1 && xml[closeEnd] === '>') {
					return { text: xml.slice(gt + 1, close), end: closeEnd + 1 };
				}
				close = xml.indexOf('</', close + 2);
			}
			return null;
		}
		lt = xml.indexOf('<', lt + 1);
	}
	return null;
}

/**
 * Extract the text content of the first matching tag, ignoring namespace prefixes.
 * Whitespace within the content is collapsed.
 */
export function extractFirstTagText(xml: string, localName: string): string | undefined {
	const found = findTagText(xml, xml.toLowerCase(), localName.toLowerCase(), 0);
	return found?.text.replace(/\s+/g, '').trim() || undefined;
}

/**
 * Extract the text content of all matching tags, ignoring namespace prefixes.
 * Whitespace within each match is collapsed.
 */
export function extractAllTagText(xml: string, localName: string): string[] {
	const result: string[] = [];
	const lowerXml = xml.toLowerCase();
	const lname = localName.toLowerCase();
	let found = findTagText(xml, lowerXml, lname, 0);
	while (found) {
		const value = found.text.replace(/\s+/g, '').trim();
		if (value) {
			result.push(value);
		}
		found = findTagText(xml, lowerXml, lname, found.end);
	}
	return result;
}
