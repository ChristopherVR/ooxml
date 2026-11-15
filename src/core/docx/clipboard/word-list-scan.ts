/** The lists a Word clipboard fragment uses, with their levels. */
import type { NumberingLevelDefinition } from '../numbering-model';
import { parseDeclarations, firstFontFamily } from './css';
import { tokensText, type HtmlToken } from './html-tokens';
import { levelFromMarker, levelFromRule } from './word-lists';
import type { WordStyleSheet } from './word-styles';

/** One pasted list: Word's list definition (`l0`) as used by one list instance (`lfo1`). */
export interface WordClipboardList {
	/** `l0:lfo1`, unique within one clipboard fragment. */
	key: string;
	levels: NumberingLevelDefinition[];
}

export interface WordListReference {
	key: string;
	/** 0-based `w:ilvl`. */
	level: number;
}

/** A paragraph's `mso-list:l0 level2 lfo1`, or `undefined` for none (`mso-list:none`). */
export function listReference(style: string | undefined): WordListReference | undefined {
	const value = parseDeclarations(style).get('mso-list');
	const match = value && /^(l\d+)\s+level(\d)\s+lfo(\d+)/i.exec(value);
	return match
		? { key: `${match[1]!.toLowerCase()}:lfo${match[3]}`, level: Number(match[2]) - 1 }
		: undefined;
}

/** The end index (exclusive) of the element opened at `start`, by matching its name. */
export function elementEnd(tokens: readonly HtmlToken[], start: number): number {
	const open = tokens[start];
	if (open?.kind !== 'open' || open.selfClosing) return start + 1;
	let depth = 0;
	for (let index = start; index < tokens.length; index++) {
		const token = tokens[index]!;
		if (token.kind === 'open' && token.name === open.name && !token.selfClosing) depth++;
		else if (token.kind === 'close' && token.name === open.name && --depth === 0) return index + 1;
	}
	return tokens.length;
}

interface Usage {
	listId: string;
	className: string | undefined;
	markers: Map<number, { text: string; font: string | undefined }>;
	used: Set<number>;
}

/**
 * Every list the fragment's paragraphs reference. Levels come from the stylesheet's `@list`
 * rules when present, otherwise from the marker text Word wrote for browsers (`mso-list:Ignore`).
 */
export function scanWordLists(
	tokens: readonly HtmlToken[],
	sheet: WordStyleSheet,
): Map<string, WordClipboardList> {
	const usages = new Map<string, Usage>();
	let current: { usage: Usage; level: number } | undefined;
	const fonts: Array<string | undefined> = [];
	for (let index = 0; index < tokens.length; index++) {
		const token = tokens[index]!;
		if (token.kind === 'close' && token.name === 'span') fonts.pop();
		if (token.kind !== 'open') continue;
		const style = token.attrs.get('style');
		if (/^(p|h[1-6]|li)$/.test(token.name)) {
			const reference = listReference(style);
			if (!reference) {
				current = undefined;
				continue;
			}
			const usage = usages.get(reference.key) ?? {
				listId: reference.key.split(':')[0]!,
				className: token.attrs.get('class'),
				markers: new Map(),
				used: new Set<number>(),
			};
			usage.used.add(reference.level);
			usages.set(reference.key, usage);
			current = { usage, level: reference.level };
			continue;
		}
		if (token.name !== 'span' || token.selfClosing) continue;
		const declarations = parseDeclarations(style);
		if (declarations.get('mso-list')?.toLowerCase() !== 'ignore') {
			fonts.push(firstFontFamily(declarations.get('font-family')));
			continue;
		}
		const end = elementEnd(tokens, index);
		if (current && !current.usage.markers.has(current.level))
			current.usage.markers.set(current.level, {
				text: tokensText(tokens.slice(index + 1, end - 1)),
				font: [...fonts].reverse().find((font) => font !== undefined),
			});
		index = end - 1;
	}
	const lists = new Map<string, WordClipboardList>();
	for (const [key, usage] of usages) {
		const rules = sheet.lists.get(usage.listId);
		const inherited = usage.className ? sheet.classDeclarations.get(usage.className) : undefined;
		const levels = rules?.size
			? [...rules].map(([level, rule]) => levelFromRule(level, rule, inherited))
			: [...usage.markers].map(([level, marker]) =>
					levelFromMarker(level, marker.text, marker.font),
				);
		// A level a paragraph uses but neither source describes still needs a definition.
		const missing = [...usage.used].filter((level) => !levels.some((l) => l.level === level));
		for (const level of missing) levels.push(levelFromMarker(level, '', undefined));
		levels.sort((a, b) => a.level - b.level);
		if (levels.length) lists.set(key, { key, levels });
	}
	return lists;
}
