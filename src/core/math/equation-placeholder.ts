/**
 * Equation placeholders: the dotted boxes an empty argument shows while it waits for input (the
 * numerator of a new fraction, the body of a new integral). Linear text spells them `⬚`
 * (U+2B1A), as Word's linear format does; saved Office Math keeps the argument empty.
 */
import type { OmmlNode } from './omml-node';
import { stripXmlOrderSuffix } from './omml-node';
import { childNode, contentKeys, flagOn, readRunText } from './omml-to-latex-helpers';

export const EQUATION_PLACEHOLDER = '⬚';

/** Arguments that show a placeholder when they hold nothing. */
const ARGUMENTS = new Set(['m:e', 'm:num', 'm:den', 'm:sub', 'm:sup', 'm:deg', 'm:lim', 'm:fName']);

function hidden(tag: string, node: OmmlNode, arg: string): boolean {
	const on = (pr: string, flag: string) => flagOn(childNode(node, pr), flag);
	if (tag === 'm:rad' && arg === 'm:deg') return on('m:radPr', 'm:degHide');
	if (tag === 'm:nary' && arg === 'm:sub') return on('m:naryPr', 'm:subHide');
	if (tag === 'm:nary' && arg === 'm:sup') return on('m:naryPr', 'm:supHide');
	return false;
}

const isEmpty = (value: unknown) =>
	value === '' ||
	value === undefined ||
	(typeof value === 'object' && value !== null && contentKeys(value as OmmlNode).length === 0);

function fill(tag: string, node: OmmlNode): OmmlNode {
	const out: OmmlNode = {};
	for (const [key, value] of Object.entries(node)) {
		const child = stripXmlOrderSuffix(key);
		if (key.startsWith('@_') || key === '#text' || child.endsWith('Pr') || child === 'm:t') {
			out[key] = value;
			continue;
		}
		const items = Array.isArray(value) ? value : [value];
		const next = items.map((item) => {
			if (ARGUMENTS.has(child) && tag !== 'm:oMath' && isEmpty(item) && !hidden(tag, node, child))
				return { 'm:r': { 'm:t': EQUATION_PLACEHOLDER } };
			return item && typeof item === 'object' ? fill(child, item as OmmlNode) : item;
		});
		out[key] = Array.isArray(value) ? next : next[0];
	}
	return out;
}

/** A copy of `tree` with a `⬚` run in every empty argument that Word would show a box for. */
export function fillPlaceholders(tree: OmmlNode): OmmlNode {
	return fill('', tree);
}

/** A copy of `tree` without placeholder runs: what is saved for the arguments still empty. */
export function stripPlaceholders(tree: OmmlNode): OmmlNode {
	const out: OmmlNode = {};
	for (const [key, value] of Object.entries(tree)) {
		const tag = stripXmlOrderSuffix(key);
		const items = (Array.isArray(value) ? value : [value]).flatMap((item) => {
			if (tag === 'm:r' && item && typeof item === 'object') {
				const text = readRunText(item as OmmlNode);
				if (!text.includes(EQUATION_PLACEHOLDER)) return [item];
				const rest = text.split(EQUATION_PLACEHOLDER).join('');
				return rest ? [{ ...(item as OmmlNode), 'm:t': rest }] : [];
			}
			if (key.startsWith('@_') || !item || typeof item !== 'object') return [item];
			return [stripPlaceholders(item as OmmlNode)];
		});
		if (items.length === 0 && tag === 'm:r') continue;
		out[key] = Array.isArray(value) ? items : (items[0] ?? {});
	}
	return out;
}
