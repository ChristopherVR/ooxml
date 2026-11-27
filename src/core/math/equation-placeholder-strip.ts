/** Removing placeholders before an equation is saved: part of the lazily loaded equation tools. */
import { EQUATION_PLACEHOLDER, readRunText, stripXmlOrderSuffix, type OmmlNode } from './index';

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
