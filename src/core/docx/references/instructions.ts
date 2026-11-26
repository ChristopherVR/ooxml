import { fieldTokens } from '../field-instruction';
import type { TextRun } from '../model';

export const REFERENCE_LCID: Readonly<Record<string, number>> = {
	en: 1033,
	de: 1031,
	es: 3082,
	fr: 1036,
	'zh-CN': 2052,
	sv: 1053,
};

export function quoteReference(text: string): string {
	return '"' + text.replace(/[\\"]/g, '\\$&') + '"';
}

export function referenceInstruction(instruction: string) {
	if (instruction.length > 8192) return undefined;
	const [name, ...tokens] = fieldTokens(instruction);
	if (name?.kind !== 'text') return undefined;
	const args: string[] = [];
	const switches = new Map<string, (string | undefined)[]>();
	for (let i = 0; i < tokens.length; i++) {
		const token = tokens[i];
		if (!token) continue;
		const next = tokens[i + 1];
		if (token.kind === 'text') args.push(token.value);
		else {
			const value = next?.kind === 'text' ? next.value : undefined;
			const list = switches.get(token.value) ?? [];
			list.push(value);
			switches.set(token.value, list);
			if (value !== undefined) i++;
		}
	}
	return {
		name: name.value.toUpperCase(),
		args,
		has: (key: string) => switches.has(key),
		value: (key: string) => switches.get(key)?.[0],
		values: (key: string) => switches.get(key)?.flatMap((item) => item ?? []) ?? [],
	};
}

export function indexEntryCode(options: {
	main: string;
	subentry: string;
	crossReference: string;
	bookmark: string;
	bold: boolean;
	italic: boolean;
}): string {
	const levels = [options.main, options.subentry]
		.filter(Boolean)
		.map((text) => text.replace(/[\\:]/g, '\\$&'));
	return [
		'XE',
		quoteReference(levels.join(':')),
		options.crossReference ? `\\t ${quoteReference(options.crossReference)}` : '',
		options.bookmark ? `\\r ${quoteReference(options.bookmark)}` : '',
		options.bold ? '\\b' : '',
		options.italic ? '\\i' : '',
	]
		.filter(Boolean)
		.join(' ');
}

export function authorityEntryCode(long: string, short: string, category: number): string {
	return `TA \\l ${quoteReference(long)} \\s ${quoteReference(short)} \\c ${category}`;
}

export function referenceFieldRuns(instruction: string, result?: readonly TextRun[]): TextRun[] {
	return [
		{ text: '', fieldChar: 'begin' },
		{ text: '', fieldCode: ` ${instruction.trim()} `, ...(result ? {} : { hidden: true }) },
		...(result
			? [
					{ text: '', fieldChar: 'separate' as const },
					...result.map((run) => ({ ...run, field: { instr: instruction.trim() } })),
				]
			: []),
		{ text: '', fieldChar: 'end' },
	];
}
