// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// STYLEREF (ECMA-376 Part 1, 17.16.5; Word's reading in MS-OE376 2.1.530): the text of the
// nearest paragraph in a paragraph style, or the nearest text in a character style.
import {
	applyFieldTextCase,
	fieldTokens,
	isFieldTextCase,
	type FieldTextCase,
} from './field-instruction';
import { cellSourceBlocks } from './cell-blocks';
import type { Block, DocumentModel, Paragraph, TextRun } from './model';
import { computeListLabels } from './numbering-format';
import { resolveNumberingLevel } from './numbering-parse';

/** A STYLEREF instruction this code evaluates. */
export interface StyleReference {
	/** The style argument: a style name, a style id, or `1` to `9` for the heading styles. */
	style: string;
	/** `\l`: in a header or footer, the last matching text on the page instead of the first. */
	last: boolean;
	/** `\n`: the paragraph number of the matching paragraph instead of its text. */
	number: boolean;
	/**
	 * `\t` with `\n`: the number without the text around it (`1.01` for `Section 1.01`). Word writes
	 * `\s` for the chapter number of a caption, which reads as `\n \t`.
	 */
	numberOnly: boolean;
	/** `\* Upper`, `Lower`, `FirstCap` or `Caps`. */
	textCase?: FieldTextCase;
}

/**
 * `\*` formats that leave the text as it is and only say which formatting the result keeps. The
 * result keeps that of the previous result in both cases (`CHARFORMAT`'s is not applied).
 */
const FORMATTING = new Set(['mergeformat', 'charformat']);

/**
 * Reads a STYLEREF instruction. Undefined for another field, a STYLEREF without a style, and
 * one with a switch that is not evaluated here (`\p`, `\r`, `\w`, number or date pictures,
 * number formats): such fields keep the result they were saved with.
 */
export function parseStyleReference(instr: string): StyleReference | undefined {
	const tokens = fieldTokens(instr);
	if (tokens[0]?.kind !== 'text' || tokens[0].value.toUpperCase() !== 'STYLEREF') return undefined;
	const style = tokens[1]?.kind === 'text' ? tokens[1].value : undefined;
	if (!style) return undefined;
	const reference: StyleReference = {
		style,
		last: false,
		number: false,
		numberOnly: false,
	};
	for (let index = 2; index < tokens.length; index++) {
		const token = tokens[index]!;
		if (token.kind === 'text') return undefined;
		if (token.value === 'l') reference.last = true;
		else if (token.value === 'n') reference.number = true;
		else if (token.value === 't') reference.numberOnly = true;
		else if (token.value === 's') {
			reference.number = true;
			reference.numberOnly = true;
		} else if (token.value === '*') {
			const format = tokens[++index];
			if (format?.kind !== 'text') return undefined;
			const name = format.value.toLowerCase();
			if (isFieldTextCase(name)) reference.textCase = name;
			else if (!FORMATTING.has(name)) return undefined;
		} else return undefined;
	}
	return reference;
}

/** The style a STYLEREF argument names, and whether it styles paragraphs or runs. */
export interface StyleTarget {
	kind: 'paragraph' | 'character';
	id: string;
	/** The document's default paragraph style: paragraphs without a style are in it. */
	isDefault?: boolean;
}

/**
 * Resolves a style argument as Word does: by style name in any case (`1` to `9` stand for the
 * built-in `heading 1` to `heading 9`), then, for files written by other tools, by style id.
 */
export function styleTarget(model: DocumentModel, style: string): StyleTarget | undefined {
	const name = /^[1-9]$/.test(style) ? `heading ${style}` : style;
	const key = name.toLowerCase();
	const paragraphs = Object.values(model.paragraphStyles?.styles ?? {});
	const characters = Object.values(model.characterStyles?.styles ?? {}).filter(
		(definition) => definition.type === 'character',
	);
	const paragraph = (
		definition: (typeof paragraphs)[number] | undefined,
	): StyleTarget | undefined =>
		definition && {
			kind: 'paragraph',
			id: definition.id,
			...(definition.isDefault ? { isDefault: true } : {}),
		};
	const character = (
		definition: (typeof characters)[number] | undefined,
	): StyleTarget | undefined => definition && { kind: 'character', id: definition.id };
	const named = <T extends { name?: string }>(list: T[]) =>
		list.find((definition) => definition.name?.toLowerCase() === key);
	const withId = <T extends { id: string }>(list: T[]) =>
		list.find((definition) => definition.id === name) ??
		list.find((definition) => definition.id.toLowerCase() === key);
	return (
		paragraph(named(paragraphs)) ??
		character(named(characters)) ??
		paragraph(withId(paragraphs)) ??
		character(withId(characters))
	);
}

/**
 * The paragraphs of the main document in reading order: body paragraphs, and each table's cells
 * row by row, a cell's paragraphs and the tables nested in it in the order they come. Headers,
 * footers, notes and text boxes are not searched, as in Word.
 */
export function documentParagraphs(blocks: readonly Block[]): Paragraph[] {
	const result: Paragraph[] = [];
	const visit = (items: readonly Block[]) => {
		for (const block of items) {
			if (block.type === 'paragraph') result.push(block);
			else for (const row of block.rows) for (const cell of row) visit(cellSourceBlocks(cell));
		}
	};
	visit(blocks);
	return result;
}

/** Text a STYLEREF can pick up: hidden text and text a revision deletes or moves away are not. */
const shownText = (run: TextRun): string =>
	run.vanish || run.revision?.kind === 'delete' || run.revision?.kind === 'moveFrom'
		? ''
		: run.text;

/** Runs of a paragraph, `from` to `to` (both included). */
export interface RunSpan {
	from: number;
	to: number;
}

/** Text a STYLEREF can take, and the runs of its paragraph it comes from. */
export interface StyleRefPortion extends RunSpan {
	/** Undefined for a match this code cannot evaluate: the field keeps its saved result. */
	text: string | undefined;
}

/**
 * Each stretch of runs in the character style; runs without text do not break one, the runs of
 * `skip` (a field's own) do and are not read.
 */
function characterPortions(paragraph: Paragraph, id: string, skip?: RunSpan): StyleRefPortion[] {
	const portions: StyleRefPortion[] = [];
	let open: StyleRefPortion | undefined;
	paragraph.runs.forEach((run, index) => {
		if (skip && index >= skip.from && index <= skip.to) {
			open = undefined;
			return;
		}
		const text = shownText(run);
		if (!text) return;
		if (run.style !== id) open = undefined;
		else if (open) {
			open.text += text;
			open.to = index;
		} else portions.push((open = { text, from: index, to: index }));
	});
	return portions;
}

const inParagraphStyle = (paragraph: Paragraph, target: StyleTarget) =>
	paragraph.style === target.id || (paragraph.style === undefined && target.isDefault === true);

/** Evaluates STYLEREF fields of one document against its paragraphs. */
export interface StyleRefSource {
	readonly paragraphs: readonly Paragraph[];
	/** Where a paragraph sits in `paragraphs`. */
	readonly index: ReadonlyMap<string, number>;
	/** The style an argument names, or undefined when the document has no such style. */
	target(style: string): StyleTarget | undefined;
	/**
	 * Every result the field could take from `paragraphs[index]`, first to last, with the runs it
	 * comes from; empty when that paragraph has no text in the style (or, with `\n`, no number).
	 * A paragraph-style match has one result; a character style has one per stretch of runs in it,
	 * leaving out the runs of `skip`.
	 */
	portions(
		index: number,
		reference: StyleReference,
		target: StyleTarget,
		skip?: RunSpan,
	): StyleRefPortion[];
	/** The texts of `portions` (undefined where the field keeps its saved result). */
	results(index: number, reference: StyleReference, target: StyleTarget): (string | undefined)[];
}

/** Reads one document for STYLEREF fields. The model must not change while this is used. */
export function styleRefSource(model: DocumentModel): StyleRefSource {
	const paragraphs = documentParagraphs(model.blocks);
	const index = new Map(paragraphs.map((paragraph, at) => [paragraph.id, at]));
	const targets = new Map<string, StyleTarget | undefined>();
	let labels: ReturnType<typeof computeListLabels> | undefined;
	/** The paragraph's number; undefined without one, null when `\t` cannot be applied to it. */
	const number = (paragraph: Paragraph, numberOnly: boolean): string | null | undefined => {
		labels ??= computeListLabels(model);
		const label = labels.get(paragraph.id);
		if (!label || !model.numberingCatalog) return undefined;
		const level = resolveNumberingLevel(model.numberingCatalog, label.numId, label.level);
		if (!level || level.numFmt === 'bullet') return undefined;
		let text = label.text;
		if (numberOnly) {
			const first = level.lvlText.indexOf('%');
			const end = level.lvlText.search(/%[1-9](?!.*%[1-9])/);
			const prefix = first < 0 ? '' : level.lvlText.slice(0, first);
			const suffix = end < 0 ? '' : level.lvlText.slice(end + 2);
			// Words between the levels (`Part %1, Section %2`): what Word drops of them is unknown.
			const inner = first < 0 || end < 0 ? '' : level.lvlText.slice(first, end + 2);
			if (/\p{L}/u.test(inner.replace(/%[1-9]/g, ''))) return null;
			if (prefix && text.startsWith(prefix)) text = text.slice(prefix.length);
			if (suffix && text.endsWith(suffix)) text = text.slice(0, -suffix.length);
		}
		return text.trim().replace(/\.+$/, '');
	};
	const portions = (
		at: number,
		reference: StyleReference,
		target: StyleTarget,
		skip?: RunSpan,
	): StyleRefPortion[] => {
		const paragraph = paragraphs[at];
		if (!paragraph) return [];
		let found: StyleRefPortion[];
		if (target.kind === 'character') found = characterPortions(paragraph, target.id, skip);
		else
			// An empty paragraph in the style is a match too, which shows nothing, as in Word.
			found = inParagraphStyle(paragraph, target)
				? [{ text: paragraph.runs.map(shownText).join(''), from: 0, to: paragraph.runs.length - 1 }]
				: [];
		if (reference.number && found.length) {
			const label = number(paragraph, reference.numberOnly);
			found =
				label === undefined
					? []
					: found.map((portion) => ({ ...portion, text: label === null ? undefined : label }));
		}
		return reference.textCase
			? found.map((portion) => ({
					...portion,
					text:
						portion.text === undefined
							? undefined
							: applyFieldTextCase(portion.text, reference.textCase!),
				}))
			: found;
	};
	return {
		paragraphs,
		index,
		target(style) {
			if (!targets.has(style)) targets.set(style, styleTarget(model, style));
			return targets.get(style);
		},
		portions,
		results: (at, reference, target) =>
			portions(at, reference, target).map((portion) => portion.text),
	};
}
