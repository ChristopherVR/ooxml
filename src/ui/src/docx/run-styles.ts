import { Plugin, PluginKey, type EditorState } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import {
	resolveRunFormatting,
	resolveTableStyleFormatting,
	resolveThemeColorReference,
	reviewRunFormatting,
	type DocumentModel,
	type RunFormatting,
	type Table,
	type TextRun,
} from 'ooxml-core/docx';
import { appendInlineNode } from './run-adapter';
import { ligatureStyle } from 'ooxml-core/docx/ui';
import { scaledSegments, scaleMeasurer } from './run-scale';
import { scaledParagraphBreaks } from './scaled-paragraph-breaks';
import type { ReviewDisplayMode } from './review-display';
import { displayParagraph } from './review-paragraph-display';
import { reviewRunCss } from './review-run-css';

type Theme = NonNullable<DocumentModel['theme']>;

export function themeFontOf(
	formatting: RunFormatting,
	theme: Theme | undefined,
): string | undefined {
	const role = formatting.fontTheme?.ascii ?? formatting.fontTheme?.hAnsi;
	return role && theme ? theme.fonts[role]?.latin : undefined;
}

const quoteFont = (family: string) => `"${family.replace(/["\\]/g, '')}"`;
const safeColor = (value: string | undefined) =>
	value && /^#[0-9a-f]{6}$/i.test(value) ? value : undefined;

/** CSS for inherited run formatting; direct formatting still renders through the run's own marks. */
export function runFormattingCss(
	formatting: RunFormatting,
	theme?: Theme,
	direct: RunFormatting = {},
): string {
	const css: string[] = [];
	const script =
		formatting.verticalAlign === 'superscript' || formatting.verticalAlign === 'subscript';
	if (formatting.ligatures) css.push(ligatureStyle(formatting.ligatures));
	// Word toggle properties cancel out when set at an odd number of levels; a direct toggle that the
	// hierarchy cancels still renders its mark, so the decoration (inside the mark) resets it.
	if (formatting.bold) css.push('font-weight:700');
	else if (direct.bold) css.push('font-weight:400');
	if (formatting.italic) css.push('font-style:italic');
	else if (direct.italic) css.push('font-style:normal');
	const lines = [
		formatting.underline ? 'underline' : '',
		formatting.strike || formatting.doubleStrike ? 'line-through' : '',
	].filter(Boolean);
	if (lines.length) css.push(`text-decoration-line:${lines.join(' ')}`);
	if (formatting.caps) css.push('text-transform:uppercase');
	else if (direct.caps) css.push('text-transform:none');
	if (formatting.smallCaps) css.push('font-variant:small-caps');
	else if (direct.smallCaps) css.push('font-variant:normal');
	if (formatting.fontSize) css.push(`font-size:${formatting.fontSize * (script ? 0.65 : 1)}pt`);
	if (formatting.verticalAlign && direct.verticalAlign === undefined)
		css.push(
			`vertical-align:${formatting.verticalAlign === 'baseline' ? 'baseline' : formatting.verticalAlign === 'superscript' ? 'super' : 'sub'}`,
		);
	const family = formatting.fontFamily ?? themeFontOf(formatting, theme);
	if (family) css.push(`font-family:${quoteFont(family)}`);
	const color =
		safeColor(formatting.color) ??
		(formatting.colorTheme && theme
			? safeColor(resolveThemeColorReference(formatting.colorTheme, theme))
			: undefined);
	if (color) css.push(`color:${color}`);
	if (formatting.characterSpacingTwips !== undefined)
		css.push(`letter-spacing:${formatting.characterSpacingTwips / 20}pt`);
	// Direct position lives on its mark, including in read-only previews. Applying it to an inner
	// decoration too would move the baseline twice.
	if (formatting.positionHalfPoints !== undefined && direct.positionHalfPoints === undefined) {
		if (script && direct.verticalAlign === undefined)
			css.push('position:relative', `top:${-formatting.positionHalfPoints / 2}pt`);
		else css.push(`vertical-align:${formatting.positionHalfPoints / 2}pt`);
	}
	if (formatting.kerningHalfPoints !== undefined) {
		const threshold = formatting.kerningHalfPoints / 2;
		css.push(
			`font-kerning:${threshold > 0 && (formatting.fontSize ?? 11) >= threshold ? 'normal' : 'none'}`,
		);
	}
	return css.join(';');
}

/** The run an inline node stands for, with its marks or attributes as direct formatting. */
export function runOf(node: ProseMirrorNode): TextRun | undefined {
	const runs: TextRun[] = [];
	appendInlineNode(runs, node);
	return runs[0];
}

/**
 * Renders formatting a run inherits from document defaults, its paragraph style, its table's style
 * (e.g. a bold header row), its character style and the theme. It is display-only: decorations
 * never enter the model or collaboration steps.
 */
export const runStylesKey = new PluginKey('dve-run-styles');

/** The document model the editor's run-style decorations resolve against, if the plugin is on. */
export function styleModelOf(state: EditorState): DocumentModel | undefined {
	return (
		runStylesKey.get(state)?.spec as { getModel?: () => DocumentModel } | undefined
	)?.getModel?.();
}

export function runStylesPlugin(
	getModel: () => DocumentModel,
	getMode: () => ReviewDisplayMode = () => 'all',
) {
	let measurer = scaleMeasurer();
	return new Plugin({
		key: runStylesKey,
		getModel,
		view(view) {
			const fonts = view.dom.ownerDocument.fonts;
			const refresh = () => {
				measurer = scaleMeasurer();
				view.dispatch(view.state.tr.setMeta('addToHistory', false).setMeta(runStylesKey, 'fonts'));
			};
			fonts?.addEventListener?.('loadingdone', refresh);
			return { destroy: () => fonts?.removeEventListener?.('loadingdone', refresh) };
		},
		props: {
			decorations(state) {
				const model = getModel();
				const mode = getMode();
				const tables = new Map<string, Table>();
				for (const block of model.blocks) if (block.type === 'table') tables.set(block.id, block);
				const decorations: Decoration[] = [];
				const visitParagraph = (
					paragraph: ProseMirrorNode,
					start: number,
					tableStyleRun?: RunFormatting,
				) => {
					const paragraphStyleId = displayParagraph(paragraph, mode).value.style;
					let hasScale = false;
					paragraph.forEach((child, offset) => {
						if (!child.isInline || child.type.name === 'equation') return;
						const current = runOf(child);
						if (!current) return;
						const projected = reviewRunFormatting(current, mode);
						const run = projected.value;
						const resolved = resolveRunFormatting(run, {
							runCatalog: model.characterStyles,
							paragraphCatalog: model.paragraphStyles,
							paragraphStyleId,
							...(tableStyleRun ? { tableStyleRun } : {}),
						});
						const style =
							mode === 'original'
								? `${runFormattingCss(resolved, model.theme)};${reviewRunCss(resolved)}`
								: runFormattingCss(resolved, model.theme, run);
						const originalAttrs =
							mode === 'original'
								? {
										...(run.language ? { lang: run.language } : {}),
										...(run.rtl !== undefined ? { dir: run.rtl ? 'rtl' : 'ltr' } : {}),
										...(projected.error ? { 'data-review-format-error': projected.error } : {}),
									}
								: {};
						const hidden = resolved.vanish === true;
						if (!child.isText) {
							decorations.push(
								Decoration.node(start + offset, start + offset + child.nodeSize, {
									style: `${runFormattingCss(resolved, model.theme)};${reviewRunCss(resolved)}`,
									...(run.language ? { lang: run.language } : {}),
									...(run.rtl !== undefined ? { dir: run.rtl ? 'rtl' : 'ltr' } : {}),
									...(projected.error ? { 'data-review-format-error': projected.error } : {}),
									...(hidden ? { class: 'dve-hidden-text' } : {}),
								}),
							);
							return;
						}
						const segments = scaledSegments(
							child.text!,
							resolved,
							resolved.fontFamily ?? themeFontOf(resolved, model.theme),
							measurer,
						);
						if (segments.length) {
							hasScale = true;
							for (const segment of segments)
								decorations.push(
									Decoration.inline(
										start + offset + segment.from,
										start + offset + segment.to,
										{
											style: `${style};${segment.css}`,
											...originalAttrs,
											class: `dve-scaled-text${hidden ? ' dve-hidden-text' : ''}`,
										},
										{ segment: segment.from },
									),
								);
							return;
						}
						if (style || hidden)
							decorations.push(
								Decoration.inline(start + offset, start + offset + child.nodeSize, {
									...originalAttrs,
									...(style ? { style } : {}),
									...(hidden ? { class: 'dve-hidden-text' } : {}),
								}),
							);
					});
					if (hasScale) decorations.push(...scaledParagraphBreaks(paragraph, start));
				};
				state.doc.forEach((block, blockOffset) => {
					if (block.type.name === 'paragraph') return visitParagraph(block, blockOffset + 1);
					if (block.type.name !== 'table') return;
					const table = tables.get(String(block.attrs.id));
					let columnCount = 1;
					block.forEach((row) => {
						let width = 0;
						row.forEach((cell) => (width += Number(cell.attrs.colspan) || 1));
						columnCount = Math.max(columnCount, width);
					});
					block.forEach((row, rowOffset, rowIndex) => {
						let column = 0;
						const rowStart = blockOffset + 1 + rowOffset + 1;
						row.forEach((cell, cellOffset) => {
							const tableStyleRun = table?.style
								? resolveTableStyleFormatting(
										table.style,
										model.tableStyles,
										table.look,
										rowIndex,
										block.childCount,
										column,
										columnCount,
									).run
								: undefined;
							cell.forEach((paragraph, paragraphOffset) => {
								if (paragraph.type.name === 'paragraph')
									visitParagraph(
										paragraph,
										rowStart + cellOffset + 1 + paragraphOffset + 1,
										tableStyleRun,
									);
							});
							column += Number(cell.attrs.colspan) || 1;
						});
					});
				});
				return DecorationSet.create(state.doc, decorations);
			},
		},
	});
}
