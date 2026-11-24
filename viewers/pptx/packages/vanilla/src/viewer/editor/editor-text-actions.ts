import type { ChangeCaseMode, ListConversionScope } from 'ooxml-ui/pptx';
import { stepFontSizePt, tableTextShadowUpdate, transformInlineListCase } from 'ooxml-ui/pptx';
import type { PptxThemeColorRef, TextStyle } from 'pptx-viewer-core';

import type { Store, ViewerState } from '../state';
import type { ApplyToSelected } from './editor-apply-to-selected';
import {
	adjustFontSize,
	changeTextCase,
	clearFormatting,
	patchTextStyle,
	setCharacterSpacing,
	setFontFamily,
	setFontSize,
	setHighlightColor,
	setTextColor,
	stepFontSize,
	toggleTextProp,
	toggleTextShadow,
} from './editor-format-mutations';
import {
	adjustIndent,
	setColumnCount,
	setLineSpacing,
	setTextAlign,
	setTextDirection,
	toggleListType,
} from './editor-paragraph-mutations';
import { recordRecentColor } from './editor-recent-colors';
import { tableTextChangeCasePatch, tableTextFormatPatch } from './editor-table-text-format';

/**
 * Character + paragraph formatting actions for the ribbon's Home > Font and
 * Home > Paragraph groups. Every method is a thin `applyToSelected` wrapper
 * around the pure builders in `editor-format-mutations.ts` /
 * `editor-paragraph-mutations.ts`; this file owns none of the mutation logic
 * itself, only the selection/history wiring shared by the whole action set.
 */
export interface TextActions {
	toggleBold(): void;
	toggleItalic(): void;
	toggleUnderline(): void;
	toggleStrikethrough(): void;
	toggleTextShadow(): void;
	changeFontSize(delta: number): void;
	/** Ctrl+Shift+>/< and Ctrl+]/[: step through PowerPoint's font-size ladder. */
	stepFontSize(direction: 'increase' | 'decrease'): void;
	setFontSize(size: number): void;
	setFontFamily(family: string): void;
	setTextColor(color: string, ref?: PptxThemeColorRef): void;
	setHighlightColor(color: string): void;
	setCharacterSpacing(value: number): void;
	changeCase(mode: ChangeCaseMode): void;
	clearFormatting(): void;
	/** Set (not toggle) character/paragraph style fields on the selection, e.g. the title-bar search. */
	patchSelectedTextStyle(patch: Partial<TextStyle>): void;
	toggleBulletList(): void;
	toggleNumberedList(): void;
	increaseIndent(): void;
	decreaseIndent(): void;
	setTextAlign(align: TextStyle['align']): void;
	setLineSpacing(value: number): void;
	setTextDirection(direction: TextStyle['textDirection']): void;
	setColumnCount(count: number): void;
}

/**
 * `getListConversionScope` is the live host `listConversionScope` option, read
 * on every Bullets / Numbering press (unset keeps PowerPoint's single paragraph).
 */
export function createTextActions(
	store: Store<ViewerState>,
	applyToSelected: ApplyToSelected,
	getListConversionScope?: () => ListConversionScope | undefined,
): TextActions {
	const listOptions = () => ({ listConversionScope: getListConversionScope?.() });
	// A selected table formats every run of its selected cells (shared op).
	const table = tableTextFormatPatch.bind(null, store);
	const toggle = (key: 'bold' | 'italic' | 'underline' | 'strikethrough') =>
		applyToSelected(
			(el) => table(el, (current) => ({ [key]: !current[key] })) ?? toggleTextProp(el, key),
		);
	const tableSize = (pt: number) => ({ fontSize: Math.min(400, Math.max(1, pt)) });
	return {
		toggleBold: () => toggle('bold'),
		toggleItalic: () => toggle('italic'),
		toggleUnderline: () => toggle('underline'),
		toggleStrikethrough: () => toggle('strikethrough'),
		toggleTextShadow: () =>
			applyToSelected(
				(el) =>
					table(el, (current) => tableTextShadowUpdate(Boolean(current.textShadowColor))) ??
					toggleTextShadow(el),
			),
		changeFontSize: (delta) =>
			applyToSelected(
				(el) =>
					table(el, (current) => tableSize((current.fontSize ?? 18) + delta)) ??
					adjustFontSize(el, delta),
			),
		stepFontSize: (direction) =>
			applyToSelected(
				(el) =>
					table(el, (current) => tableSize(stepFontSizePt(current.fontSize ?? 18, direction))) ??
					stepFontSize(el, direction),
			),
		setFontSize: (size) =>
			applyToSelected((el) => table(el, () => tableSize(size)) ?? setFontSize(el, size)),
		setFontFamily: (family) =>
			applyToSelected(
				(el) => table(el, () => ({ fontFamily: family })) ?? setFontFamily(el, family),
			),
		setTextColor: (color, ref) => {
			recordRecentColor(store, color);
			applyToSelected((el) => table(el, () => ({ color })) ?? setTextColor(el, color, ref));
		},
		setHighlightColor: (color) => {
			recordRecentColor(store, color);
			applyToSelected(
				(el) => table(el, () => ({ highlightColor: color })) ?? setHighlightColor(el, color),
			);
		},
		setCharacterSpacing: (value) =>
			applyToSelected(
				(el) => table(el, () => ({ characterSpacing: value })) ?? setCharacterSpacing(el, value),
			),
		changeCase: (mode) =>
			applyToSelected((el, snapshot) => {
				const tablePatch = tableTextChangeCasePatch(store, el, mode);
				if (tablePatch) {
					return tablePatch;
				}
				if (!snapshot) {
					return changeTextCase(el, mode);
				}
				const { text, textSegments } = transformInlineListCase(snapshot, null, mode);
				return { text, textSegments };
			}),
		clearFormatting: () =>
			applyToSelected(
				(el) =>
					table(el, () => ({
						bold: false,
						italic: false,
						underline: false,
						strikethrough: false,
						highlightColor: undefined,
						characterSpacing: 0,
						...tableTextShadowUpdate(true),
					})) ?? clearFormatting(el),
			),
		patchSelectedTextStyle: (patch) => applyToSelected((el) => patchTextStyle(el, patch)),
		toggleBulletList: () =>
			applyToSelected((el, snapshot) => toggleListType(el, 'bullet', snapshot, listOptions())),
		toggleNumberedList: () =>
			applyToSelected((el, snapshot) => toggleListType(el, 'numbered', snapshot, listOptions())),
		increaseIndent: () => applyToSelected((el) => adjustIndent(el, 1)),
		decreaseIndent: () => applyToSelected((el) => adjustIndent(el, -1)),
		setTextAlign: (align) => applyToSelected((el) => setTextAlign(el, align)),
		setLineSpacing: (value) => applyToSelected((el) => setLineSpacing(el, value)),
		setTextDirection: (direction) => applyToSelected((el) => setTextDirection(el, direction)),
		setColumnCount: (count) => applyToSelected((el) => setColumnCount(el, count)),
	};
}
