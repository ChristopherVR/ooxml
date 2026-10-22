// A text-level scanner for the few XML containers the Visio mapping splits (`Pages`, `Shapes`,
// `Connects`). It returns spans into the original text instead of re-serializing a DOM, so the
// shapes a peer shares are byte-for-byte what the editor wrote.

export interface ElementSpan {
	/** Local name (prefix removed). */
	name: string;
	/** Offset of `<`. */
	start: number;
	/** Offset just past the closing `>` (or of the self-closing tag). */
	end: number;
	/** Offset just past the opening tag's `>`. */
	openEnd: number;
	/** Offset of the closing tag's `<`; equals `openEnd` for a self-closing element. */
	closeStart: number;
	selfClosing: boolean;
}

export interface Container extends ElementSpan {
	children: ElementSpan[];
}

interface Tag {
	name: string;
	closing: boolean;
	selfClosing: boolean;
	start: number;
	end: number;
}

function* tags(xml: string): Generator<Tag> {
	let i = 0;
	while (i < xml.length) {
		const lt = xml.indexOf('<', i);
		if (lt < 0) return;
		if (xml.startsWith('<!--', lt)) {
			const close = xml.indexOf('-->', lt + 4);
			i = close < 0 ? xml.length : close + 3;
			continue;
		}
		if (xml.startsWith('<![CDATA[', lt)) {
			const close = xml.indexOf(']]>', lt + 9);
			i = close < 0 ? xml.length : close + 3;
			continue;
		}
		if (xml[lt + 1] === '?' || xml[lt + 1] === '!') {
			const close = xml.indexOf('>', lt + 2);
			i = close < 0 ? xml.length : close + 1;
			continue;
		}
		let j = lt + 1;
		const closing = xml[j] === '/';
		if (closing) j++;
		const nameStart = j;
		while (j < xml.length && !/[\s/>]/.test(xml[j]!)) j++;
		const qualified = xml.slice(nameStart, j);
		let quote = '';
		while (j < xml.length) {
			const ch = xml[j]!;
			if (quote) {
				if (ch === quote) quote = '';
			} else if (ch === '"' || ch === "'") quote = ch;
			else if (ch === '>') break;
			j++;
		}
		if (j >= xml.length) return;
		const colon = qualified.indexOf(':');
		yield {
			name: colon < 0 ? qualified : qualified.slice(colon + 1),
			closing,
			selfClosing: !closing && xml[j - 1] === '/',
			start: lt,
			end: j + 1,
		};
		i = j + 1;
	}
}

/**
 * The first element named `parent` (any depth) with its direct child elements, or undefined when
 * it is missing or malformed. Child spans cover the whole element including its subtree.
 */
export function scanContainer(xml: string, parent: string): Container | undefined {
	const stack: { name: string; start: number; openEnd: number }[] = [];
	let found: { start: number; openEnd: number; depth: number } | undefined;
	const children: ElementSpan[] = [];
	for (const tag of tags(xml)) {
		if (tag.closing) {
			const open = stack.pop();
			if (!open || open.name !== tag.name) return undefined;
			if (!found) continue;
			if (stack.length === found.depth)
				return {
					name: parent,
					start: found.start,
					openEnd: found.openEnd,
					closeStart: tag.start,
					end: tag.end,
					selfClosing: false,
					children,
				};
			if (stack.length === found.depth + 1)
				children.push({
					name: tag.name,
					start: open.start,
					openEnd: open.openEnd,
					closeStart: tag.start,
					end: tag.end,
					selfClosing: false,
				});
			continue;
		}
		if (!found && tag.name === parent) {
			if (tag.selfClosing)
				return {
					name: parent,
					start: tag.start,
					openEnd: tag.end,
					closeStart: tag.end,
					end: tag.end,
					selfClosing: true,
					children: [],
				};
			found = { start: tag.start, openEnd: tag.end, depth: stack.length };
			stack.push({ name: tag.name, start: tag.start, openEnd: tag.end });
			continue;
		}
		if (found && stack.length === found.depth + 1 && tag.selfClosing)
			children.push({
				name: tag.name,
				start: tag.start,
				openEnd: tag.end,
				closeStart: tag.end,
				end: tag.end,
				selfClosing: true,
			});
		if (!tag.selfClosing) stack.push({ name: tag.name, start: tag.start, openEnd: tag.end });
	}
	return undefined;
}

/** The value of attribute `name` in the opening tag text `tag`, or undefined. */
export function attrOf(tag: string, name: string): string | undefined {
	const match = new RegExp(String.raw`\s${name}\s*=\s*(?:"([^"]*)"|'([^']*)')`).exec(tag);
	return match ? (match[1] ?? match[2]) : undefined;
}

/** XML-escapes an attribute value written between double quotes. */
export const escapeAttr = (value: string): string =>
	value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

/** Unescapes the five predefined entities and numeric references of an attribute value. */
export const unescapeAttr = (value: string): string =>
	value.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (_m, entity: string) => {
		switch (entity.toLowerCase()) {
			case 'amp':
				return '&';
			case 'lt':
				return '<';
			case 'gt':
				return '>';
			case 'quot':
				return '"';
			case 'apos':
				return "'";
		}
		const code =
			entity[1]?.toLowerCase() === 'x' ? parseInt(entity.slice(2), 16) : +entity.slice(1);
		return Number.isFinite(code) ? String.fromCodePoint(code) : '';
	});
