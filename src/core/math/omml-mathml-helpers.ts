import type { OmmlNode } from './omml-node.js';
export type ChildrenConverter = (node: OmmlNode) => string;
// ── Helpers ────────────────────────────────────────────────────────────────

/** Safely retrieve a child node, always returning an object (never undefined). */
export function child(node: OmmlNode | undefined, key: string): OmmlNode {
	if (!node) {
		return {};
	}
	const v = node[key];
	if (v && typeof v === 'object' && !Array.isArray(v)) {
		return v as OmmlNode;
	}
	return {};
}

/** Ensure a value is an array of OmmlNode. */
export function ensureArray(value: OmmlNode[keyof OmmlNode]): OmmlNode[] {
	if (value === undefined || value === null) {
		return [];
	}
	if (Array.isArray(value)) {
		return value as OmmlNode[];
	}
	if (typeof value === 'object') {
		return [value as OmmlNode];
	}
	return [];
}

/** Read a string attribute from a node. */
export function attr(node: OmmlNode | undefined, name: string): string {
	if (!node) {
		return '';
	}
	const v = node[`@_${name}`];
	return typeof v === 'string' ? v : v !== undefined ? String(v) : '';
}

/** Read the `@_val` attribute (extremely common in OMML property nodes). */
export function val(node: OmmlNode | undefined): string {
	return attr(node, 'val');
}

/** Escape angle brackets and ampersands for safe embedding in MathML. */
export function escapeXml(text: string): string {
	return text.replace(/&/gu, '&amp;').replace(/</gu, '&lt;').replace(/>/gu, '&gt;');
}

/** Read an `m:t` text run value (object/string/number) as a string. */
export function readText(node: OmmlNode): string {
	const textNode = node['m:t'];
	if (typeof textNode === 'string') {
		return textNode;
	}
	if (typeof textNode === 'number' || typeof textNode === 'boolean') {
		return String(textNode);
	}
	if (textNode && typeof textNode === 'object' && !Array.isArray(textNode)) {
		const inner = (textNode as OmmlNode)['#text'];
		return typeof inner === 'string' ? inner : '';
	}
	return '';
}

// ── Character classification ────────────────────────────────────────────────

/** Set of characters treated as mathematical operators. */
export const OPERATOR_CHARS = new Set([
	'+',
	'-',
	'−',
	'±',
	'∓',
	'×',
	'÷',
	'·',
	'=',
	'≠',
	'≈',
	'≡',
	'≤',
	'≥',
	'<',
	'>',
	'≪',
	'≫',
	'∈',
	'∉',
	'⊂',
	'⊃',
	'⊆',
	'⊇',
	'∪',
	'∩',
	'→',
	'←',
	'↔',
	'⇒',
	'⇐',
	'⇔',
	'∞',
	'∴',
	'∵',
	'∝',
	'∀',
	'∃',
	',',
	';',
	':',
	'!',
	'?',
	'.',
	'|',
	'/',
	'\\',
	"'",
	'(',
	')',
	'[',
	']',
	'{',
	'}',
	'⟨',
	'⟩',
]);

export function isOperator(ch: string): boolean {
	return OPERATOR_CHARS.has(ch.trim());
}

export function isNumeric(text: string): boolean {
	return /^[0-9]+(?:[.,][0-9]+)?$/u.test(text.trim());
}

// ── Unicode accent map (m:acc) ──────────────────────────────────────────────

export const ACCENT_MAP: Record<string, string> = {
	'̂': '^', // combining circumflex → hat
	'̃': '~', // combining tilde
	'̄': '¯', // combining macron → bar
	'̅': '¯', // combining overline → bar
	'̇': '˙', // combining dot above
	'̈': '¨', // combining diaeresis
	'̌': 'ˇ', // combining caron
	'̲': '_', // combining underbar
	'⃗': '→', // combining right arrow above → →
	'^': '^',
	'~': '~',
	'¯': '¯',
	'˙': '˙',
	'¨': '¨',
	ˇ: 'ˇ',
};

// ── Nary operator map ───────────────────────────────────────────────────────

export const NARY_CHAR_MAP: Record<string, string> = {
	'∑': '∑', // ∑
	'∏': '∏', // ∏
	'∫': '∫', // ∫
	'∬': '∬', // ∬
	'∭': '∭', // ∭
	'∮': '∮', // ∮
	'∐': '∐', // ∐
	'⋀': '⋀', // ⋀
	'⋁': '⋁', // ⋁
	'⋂': '⋂', // ⋂
	'⋃': '⋃', // ⋃
};

// ── Delimiter bracket maps ──────────────────────────────────────────────────

export const DELIM_BEGIN_MAP: Record<string, string> = {
	'(': '(',
	'[': '[',
	'{': '{',
	'|': '|',
	'‖': '‖',
	'⟨': '⟨',
	'⌈': '⌈',
	'⌊': '⌊',
};

export const DELIM_END_MAP: Record<string, string> = {
	')': ')',
	']': ']',
	'}': '}',
	'|': '|',
	'‖': '‖',
	'⟩': '⟩',
	'⌉': '⌉',
	'⌋': '⌋',
};
