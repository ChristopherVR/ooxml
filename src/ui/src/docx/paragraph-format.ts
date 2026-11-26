import {
	resolveParagraphFormatting,
	type DocumentModel,
	type ParagraphFormatting,
} from 'ooxml-core/docx';
import { paragraphFromAttrs } from 'ooxml-core/docx/ui';
import { tableStyleAt } from './styled-paragraphs';
import { closeHistory } from 'prosemirror-history';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

export type Alignment = 'left' | 'center' | 'right' | 'justify';
export type Special = 'none' | 'firstLine' | 'hanging';
export type LineRule = 'single' | 'oneAndHalf' | 'double' | 'atLeast' | 'exactly' | 'multiple';

/** Everything Word's Paragraph dialog edits, in the units it shows (inches for indents, points for spacing). */
export interface ParagraphFormat {
	align: Alignment;
	leftInches: number;
	rightInches: number;
	special: Special;
	specialInches: number;
	beforePt: number;
	afterPt: number;
	lineRule: LineRule;
	/** Points for `atLeast` and `exactly`, a line multiple for `multiple`. */
	lineAt: number;
	contextualSpacing: boolean;
	widowControl: boolean;
	keepNext: boolean;
	keepLines: boolean;
	pageBreakBefore: boolean;
	suppressLineNumbers: boolean;
}

/** A selection's paragraph format, with `null` for a field that differs between paragraphs. */
export type ParagraphFormatState = { [K in keyof ParagraphFormat]: ParagraphFormat[K] | null };

const FIELDS = [
	'align',
	'leftInches',
	'rightInches',
	'special',
	'specialInches',
	'beforePt',
	'afterPt',
	'lineRule',
	'lineAt',
	'contextualSpacing',
	'widowControl',
	'keepNext',
	'keepLines',
	'pageBreakBefore',
	'suppressLineNumbers',
] as const satisfies readonly (keyof ParagraphFormat)[];

const TWIPS_PER_INCH = 1440;

type Attrs = Record<string, unknown>;
const num = (value: unknown): number =>
	typeof value === 'number' && Number.isFinite(value) ? value : 0;

/** Direct paragraph attributes with the style chain and document defaults resolved beneath them. */
export function resolvedAttrs(
	attrs: Attrs,
	model: DocumentModel | undefined,
	tableParagraph?: ParagraphFormatting,
): Attrs {
	const direct = Object.fromEntries(
		Object.entries(attrs).filter(
			([key, value]) => value != null && (key !== 'style' || value !== ''),
		),
	);
	const catalog = model?.paragraphStyles;
	if (!catalog && !tableParagraph) return direct;
	return {
		...direct,
		...resolveParagraphFormatting(
			paragraphFromAttrs(direct, String(direct.id ?? ''), []),
			catalog,
			tableParagraph,
			model?.overrideTableStyleFontSizeAndJustification,
		),
	};
}

function formatOf(attrs: Attrs): ParagraphFormat {
	const align = attrs.align;
	const firstLine = num(attrs.firstLineTwips);
	const hanging = num(attrs.hangingTwips);
	const rule = attrs.lineSpacingRule;
	const twips = num(attrs.lineSpacingTwips) || 240;
	let lineRule: LineRule = 'single';
	let lineAt = 1;
	if (rule === 'exact' || rule === 'atLeast') {
		lineRule = rule === 'exact' ? 'exactly' : 'atLeast';
		lineAt = twips / 20;
	} else if (twips === 240) lineRule = 'single';
	else if (twips === 360) lineRule = 'oneAndHalf';
	else if (twips === 480) lineRule = 'double';
	else {
		lineRule = 'multiple';
		lineAt = twips / 240;
	}
	return {
		align:
			align === 'center' || align === 'right' || align === 'justify'
				? align
				: attrs.direction === 'rtl' && align !== 'left'
					? 'right'
					: 'left',
		leftInches: num(attrs.indentStartTwips ?? attrs.indentLeftTwips) / TWIPS_PER_INCH,
		rightInches: num(attrs.indentEndTwips ?? attrs.indentRightTwips) / TWIPS_PER_INCH,
		special: hanging > 0 ? 'hanging' : firstLine > 0 ? 'firstLine' : 'none',
		specialInches: (hanging || firstLine) / TWIPS_PER_INCH,
		beforePt: num(attrs.spacingBeforeTwips) / 20,
		afterPt: num(attrs.spacingAfterTwips) / 20,
		lineRule,
		lineAt,
		contextualSpacing: attrs.contextualSpacing === true,
		widowControl: attrs.widowControl !== false,
		keepNext: attrs.keepNext === true,
		keepLines: attrs.keepLines === true,
		pageBreakBefore: attrs.pageBreakBefore === true,
		suppressLineNumbers: attrs.suppressLineNumbers === true,
	};
}

function selectedParagraphs(state: EditorState): Map<number, ProseMirrorNode> {
	const found = new Map<number, ProseMirrorNode>();
	for (const { $from, $to } of state.selection.ranges) {
		if ($from.parent.type.name === 'paragraph') found.set($from.before(), $from.parent);
		state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
			if (node.type.name === 'paragraph') found.set(pos, node);
		});
	}
	return found;
}

/** The selected paragraphs' format, with `null` where they disagree. */
export function readParagraphFormat(
	state: EditorState,
	model: DocumentModel | undefined,
): ParagraphFormatState {
	const formats = [...selectedParagraphs(state)].map(([pos, node]) =>
		formatOf(resolvedAttrs(node.attrs, model, tableStyleAt(state.doc, pos, model)?.paragraph)),
	);
	const result = {} as Record<string, unknown>;
	for (const field of FIELDS) {
		const first = formats[0]?.[field];
		result[field] = formats.every((format) => format[field] === first) ? first : null;
	}
	return result as ParagraphFormatState;
}

const twips = (value: number, perUnit: number) => Math.max(0, Math.round(value * perUnit));

/** Attribute changes for one paragraph. The indent side follows the paragraph's logical/physical attribute. */
function attrChanges(node: ProseMirrorNode, changes: Partial<ParagraphFormat>): Attrs {
	const out: Attrs = {};
	if (changes.align !== undefined) out.align = out.justification = changes.align;
	if (changes.leftInches !== undefined)
		out[node.attrs.indentStartTwips != null ? 'indentStartTwips' : 'indentLeftTwips'] = twips(
			changes.leftInches,
			TWIPS_PER_INCH,
		);
	if (changes.rightInches !== undefined)
		out[node.attrs.indentEndTwips != null ? 'indentEndTwips' : 'indentRightTwips'] = twips(
			changes.rightInches,
			TWIPS_PER_INCH,
		);
	if (changes.special !== undefined || changes.specialInches !== undefined) {
		const kind = changes.special ?? (num(node.attrs.hangingTwips) > 0 ? 'hanging' : 'firstLine');
		const by = twips(changes.specialInches ?? 0, TWIPS_PER_INCH);
		out.firstLineTwips = kind === 'firstLine' ? by : null;
		out.hangingTwips = kind === 'hanging' ? by : null;
	}
	if (changes.beforePt !== undefined) out.spacingBeforeTwips = twips(changes.beforePt, 20);
	if (changes.afterPt !== undefined) out.spacingAfterTwips = twips(changes.afterPt, 20);
	if (changes.lineRule !== undefined || changes.lineAt !== undefined) {
		const rule = changes.lineRule ?? 'multiple';
		if (rule === 'atLeast' || rule === 'exactly') {
			out.lineSpacingRule = rule === 'exactly' ? 'exact' : 'atLeast';
			out.lineSpacingTwips = Math.max(20, twips(changes.lineAt ?? 12, 20));
		} else {
			out.lineSpacingRule = 'auto';
			out.lineSpacingTwips =
				rule === 'single'
					? 240
					: rule === 'oneAndHalf'
						? 360
						: rule === 'double'
							? 480
							: Math.max(24, twips(changes.lineAt ?? 1, 240));
		}
	}
	// Explicit values, so turning an option off also overrides one a style turns on.
	for (const key of [
		'contextualSpacing',
		'widowControl',
		'keepNext',
		'keepLines',
		'pageBreakBefore',
		'suppressLineNumbers',
	] as const)
		if (changes[key] !== undefined) out[key] = changes[key];
	return out;
}

/** Applies the fields present in `changes` to every selected paragraph as one undoable step. */
export function applyParagraphFormat(view: EditorView, changes: Partial<ParagraphFormat>): boolean {
	if (!view.editable) return false;
	let tr = view.state.tr;
	for (const [pos, node] of selectedParagraphs(view.state)) {
		for (const [attr, value] of Object.entries(attrChanges(node, changes)))
			if (node.attrs[attr] !== value) tr = tr.setNodeAttribute(pos, attr, value);
	}
	if (!tr.docChanged) return false;
	view.dispatch(closeHistory(tr).scrollIntoView());
	return true;
}
