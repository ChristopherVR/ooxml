/** Word's clipboard stylesheet: which class is which style, and its `@list` definitions. */
import { cssStringValue, readStyleRules, type CssRule } from './css';

export type WordStyleKind = 'paragraph' | 'character';

export interface WordStyleSheet {
	/** Class name to Word style name and kind (`MsoQuote` is the paragraph style `Quote`). */
	classes: Map<string, { name: string; kind: WordStyleKind }>;
	/** Paragraph style formatting by class, for list levels that inherit their indent. */
	classDeclarations: Map<string, Map<string, string>>;
	/** `@list l0:level1` declarations: list id, then the 0-based level. */
	lists: Map<string, Map<number, Map<string, string>>>;
}

const CONTEXTUAL_SPACING = /CxSp(?:First|Middle|Last)$/;

/**
 * The style name a class stands for when the stylesheet does not say (`mso-style-name` is only
 * written for styles whose name is not the class): Word's built-in classes are `Mso` plus the
 * name without spaces, and list and quote paragraphs add `CxSpFirst`/`Middle`/`Last` for
 * contextual spacing.
 */
export function styleNameFromClass(className: string): string {
	const base = className.replace(CONTEXTUAL_SPACING, '').replace(/^Mso/, '');
	return base.replace(/([a-z\d])([A-Z])/g, '$1 $2').replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2');
}

function classSelectors(rule: CssRule): Array<{ element: string; className: string }> {
	return rule.selector
		.split(',')
		.map((part) => /^([a-z]+\d?)?\.([\w-]+)$/i.exec(part.trim()))
		.filter((match): match is RegExpExecArray => Boolean(match))
		.map((match) => ({ element: (match[1] ?? '').toLowerCase(), className: match[2]! }));
}

/** Reads the `<style>` blocks of a Word clipboard document. */
export function readWordStyleSheet(html: string): WordStyleSheet {
	const sheet: WordStyleSheet = {
		classes: new Map(),
		classDeclarations: new Map(),
		lists: new Map(),
	};
	for (const rule of readStyleRules(html)) {
		const list = /^@list\s+(l\d+):level(\d)$/i.exec(rule.selector);
		if (list) {
			const levels = sheet.lists.get(list[1]!) ?? new Map<number, Map<string, string>>();
			levels.set(Number(list[2]) - 1, rule.declarations);
			sheet.lists.set(list[1]!, levels);
			continue;
		}
		for (const { element, className } of classSelectors(rule)) {
			const named = rule.declarations.get('mso-style-name');
			const kind: WordStyleKind = element === 'span' ? 'character' : 'paragraph';
			if (!sheet.classes.has(className))
				sheet.classes.set(className, {
					name: named ? cssStringValue(named) : styleNameFromClass(className),
					kind,
				});
			if (kind === 'paragraph') sheet.classDeclarations.set(className, rule.declarations);
		}
	}
	return sheet;
}

/** The paragraph style a class or heading tag names, falling back to the class itself. */
export function paragraphStyleName(
	sheet: WordStyleSheet,
	className: string | undefined,
	tag: string,
): string | undefined {
	if (className) {
		const known = sheet.classes.get(className);
		if (known?.kind === 'paragraph') return known.name;
		if (!known) return styleNameFromClass(className);
	}
	const heading = /^h([1-6])$/.exec(tag);
	return heading ? `heading ${heading[1]}` : undefined;
}

const styleKey = (value: string): string => value.toLowerCase().replace(/[\s_-]+/g, '');

/**
 * The id of the style in `styles` that a pasted style name means: the same name, ignoring case
 * and spaces (Word writes `heading 1` in styles.xml and `Heading 1` on the clipboard), or a
 * style whose id is that name without spaces (`Heading1`).
 */
export function styleIdForName(
	styles: Readonly<Record<string, { id: string; name?: string }>> | undefined,
	name: string,
): string | undefined {
	const key = styleKey(name);
	const all = Object.values(styles ?? {});
	return (
		all.find((style) => style.name !== undefined && styleKey(style.name) === key)?.id ??
		all.find((style) => styleKey(style.id) === key)?.id
	);
}

/** The character style a span's class names, when the stylesheet declares it as one. */
export function characterStyleName(
	sheet: WordStyleSheet,
	className: string | undefined,
): string | undefined {
	const known = className ? sheet.classes.get(className) : undefined;
	return known?.kind === 'character' ? known.name : undefined;
}
