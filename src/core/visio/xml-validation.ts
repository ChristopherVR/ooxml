import { fail, type VisioPackageLimits } from './package-common.js';

const xmlName = '[\\p{L}_:][\\p{L}\\p{N}\\p{M}_.:\\-\\u00b7]*';
const tagName = new RegExp(`^(${xmlName})`, 'u');
const attribute = new RegExp(
	`[ \\t\\r\\n]+(${xmlName})[ \\t\\r\\n]*=[ \\t\\r\\n]*(?:"([^"<]*)"|'([^'<]*)')`,
	'uy',
);
const xmlWhitespace = (value: string) => /^[ \t\r\n]*$/.test(value);
/** Bound DOM construction first and reject the malformed input that tolerant XML parsers repair. */
export function inspectXml(
	xml: string,
	limits: VisioPackageLimits,
	check: () => void,
	encoding = 'utf-8',
): number {
	if (xml.length > limits.maxXmlChars) fail('LIMIT_XML', 'XML character limit exceeded');
	if (
		/<!DOCTYPE|<!ENTITY/i.test(xml) ||
		/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/.test(xml)
	)
		fail('INVALID_XML', 'DTD, entities, or invalid XML characters');
	const stack: { name: string; namespaces: Record<string, string> }[] = [];
	let nodes = 0,
		roots = 0,
		at = 0;
	const node = () => {
		if (++nodes > limits.maxXmlNodes) fail('LIMIT_XML', 'XML node limit exceeded');
	};
	const text = (value: string) => {
		if (/&(?!(?:amp|lt|gt|apos|quot|#\d+|#x[\da-fA-F]+);)/.test(value))
			fail('INVALID_XML', 'Invalid XML entity reference');
		for (const match of value.matchAll(/&#(x[\da-fA-F]+|\d+);/g)) {
			const digits = match[1] ?? '',
				code = digits.startsWith('x') ? parseInt(digits.slice(1), 16) : Number(digits);
			if (
				!(
					code === 9 ||
					code === 10 ||
					code === 13 ||
					(code >= 32 && code <= 0xd7ff) ||
					(code >= 0xe000 && code <= 0xfffd) ||
					(code >= 0x10000 && code <= 0x10ffff)
				)
			)
				fail('INVALID_XML', 'Invalid XML character reference');
		}
	};
	while (at < xml.length) {
		check();
		if (xml[at] !== '<') {
			const stop = xml.indexOf('<', at),
				end = stop < 0 ? xml.length : stop,
				value = xml.slice(at, end);
			if ((!stack.length && !xmlWhitespace(value)) || value.includes(']]>'))
				fail('INVALID_XML', 'Text outside XML root or invalid text');
			text(value);
			node();
			at = end;
			continue;
		}
		node();
		if (xml.startsWith('<!--', at) || xml.startsWith('<![CDATA[', at) || xml.startsWith('<?', at)) {
			const comment = xml.startsWith('<!--', at),
				cdata = xml.startsWith('<![CDATA[', at);
			const start = at + (comment ? 4 : cdata ? 9 : 2),
				token = comment ? '-->' : cdata ? ']]>' : '?>';
			const end = xml.indexOf(token, start);
			if (end < 0) fail('INVALID_XML', 'Unterminated XML construct');
			const value = xml.slice(start, end);
			if ((comment && (value.includes('--') || value.endsWith('-'))) || (cdata && !stack.length))
				fail('INVALID_XML', 'Invalid XML comment or CDATA');
			if (
				!comment &&
				!cdata &&
				(!tagName.test(value) ||
					(/^xml(?:[ \t\r\n]|$)/i.test(value) && (at !== 0 || !/^xml[ \t\r\n]/.test(value))))
			)
				fail('INVALID_XML', 'Invalid XML processing instruction');
			if (!comment && !cdata && /^xml[ \t\r\n]/.test(value)) inspectDeclaration(value, encoding);
			at = end + token.length;
			continue;
		}
		let end = at + 1,
			quote = '';
		for (; end < xml.length; end++) {
			const char = xml[end] ?? '';
			if (quote) {
				if (char === quote) quote = '';
			} else if (char === '"' || char === "'") quote = char;
			else if (char === '>') break;
		}
		if (end === xml.length) fail('INVALID_XML', 'Unterminated XML tag');
		let body = xml.slice(at + 1, end);
		const closing = body.startsWith('/'),
			empty = body.endsWith('/');
		if (closing) body = body.slice(1);
		if (empty) body = body.slice(0, -1);
		const name = tagName.exec(body)?.[1];
		if (!name) fail('INVALID_XML', 'Invalid XML element name');
		let cursor = name.length;
		if (closing) {
			if (empty || !xmlWhitespace(body.slice(cursor)) || stack.pop()?.name !== name)
				fail('INVALID_XML', 'Mismatched XML closing tag');
		} else {
			const attrs = new Set<string>();
			const namespaces = Object.create(stack.at(-1)?.namespaces ?? null) as Record<string, string>;
			namespaces['xml'] = 'http://www.w3.org/XML/1998/namespace';
			namespaces['xmlns'] = 'http://www.w3.org/2000/xmlns/';
			while (!xmlWhitespace(body.slice(cursor))) {
				attribute.lastIndex = cursor;
				const match = attribute.exec(body);
				if (!match || attrs.has(match[1] ?? ''))
					fail('INVALID_XML', 'Malformed or duplicate XML attribute');
				attrs.add(match[1] ?? '');
				const value = match[2] ?? match[3] ?? '';
				text(value);
				if (match[1]?.startsWith('xmlns:'))
					namespaces[match[1].slice(6)] = decodeXmlAttribute(value);
				node();
				cursor = attribute.lastIndex;
			}
			const expanded = new Set<string>();
			for (const attr of attrs) {
				const colon = attr.indexOf(':');
				const uri = colon < 0 ? '' : namespaces[attr.slice(0, colon)];
				if (uri === undefined) fail('INVALID_XML', 'Unbound XML attribute prefix');
				const key = `${uri}\u0000${attr.slice(colon + 1)}`;
				if (expanded.has(key)) fail('INVALID_XML', 'Duplicate expanded XML attribute name');
				expanded.add(key);
			}
			if (!stack.length && ++roots > 1) fail('INVALID_XML', 'Multiple XML roots');
			if (stack.length + 1 > limits.maxXmlDepth) fail('LIMIT_XML', 'XML depth limit exceeded');
			if (!empty) stack.push({ name, namespaces });
		}
		at = end + 1;
	}
	if (stack.length || roots !== 1) fail('INVALID_XML', 'Missing or unclosed XML root');
	return nodes;
}

function inspectDeclaration(value: string, encoding: string): void {
	const declarations: [string, string][] = [];
	let cursor = 3;
	while (!xmlWhitespace(value.slice(cursor))) {
		attribute.lastIndex = cursor;
		const match = attribute.exec(value);
		if (!match) fail('INVALID_XML', 'Malformed XML declaration');
		declarations.push([match[1] ?? '', match[2] ?? match[3] ?? '']);
		cursor = attribute.lastIndex;
	}
	if (declarations.shift()?.join('=') !== 'version=1.0')
		fail('INVALID_XML', 'Only XML version 1.0 is supported');
	if (declarations[0]?.[0] === 'encoding') {
		const declared = (declarations.shift()?.[1] ?? '').toLowerCase().replace(/-/g, '');
		const actual = encoding.replace(/-/g, '');
		if (declared !== actual && !(declared === 'utf16' && actual.startsWith('utf16')))
			fail('INVALID_XML', 'Unsupported or mismatched XML encoding declaration');
	}
	if (declarations[0]?.[0] === 'standalone') {
		const standalone = declarations.shift()?.[1];
		if (standalone !== 'yes' && standalone !== 'no')
			fail('INVALID_XML', 'Invalid standalone XML declaration');
	}
	if (declarations.length) fail('INVALID_XML', 'Invalid XML declaration attributes');
}

function decodeXmlAttribute(value: string): string {
	const entities: Record<string, string> = { amp: '&', lt: '<', gt: '>', apos: "'", quot: '"' };
	return value
		.replace(/[\t\r\n]/g, ' ')
		.replace(/&(amp|lt|gt|apos|quot|#x[\da-fA-F]+|#\d+);/g, (_, entity: string) => {
			if (entity.startsWith('#x')) return String.fromCodePoint(parseInt(entity.slice(2), 16));
			if (entity.startsWith('#')) return String.fromCodePoint(Number(entity.slice(1)));
			return entities[entity] ?? '';
		});
}

export function inspectNamespaces(root: Element, check: () => void): void {
	const xml = 'http://www.w3.org/XML/1998/namespace';
	const xmlns = 'http://www.w3.org/2000/xmlns/';
	const pending: Element[] = [root];
	while (pending.length) {
		check();
		const element = pending.pop();
		if (!element) break;
		const seen = new Set<string>();
		for (const attr of Array.from(element.attributes)) {
			const key = `${attr.namespaceURI ?? ''}\u0000${attr.localName}`;
			if (seen.has(key)) fail('INVALID_XML', 'Duplicate expanded XML attribute name');
			seen.add(key);
			if (
				attr.namespaceURI === xmlns &&
				(attr.name === 'xmlns:xmlns' ||
					attr.value === xmlns ||
					(attr.name === 'xmlns:xml' ? attr.value !== xml : attr.value === xml) ||
					(attr.name !== 'xmlns' && !attr.value))
			)
				fail('INVALID_XML', 'Invalid reserved XML namespace binding');
		}
		for (const child of Array.from(element.childNodes))
			if (child.nodeType === 1) pending.push(child as Element);
	}
}
