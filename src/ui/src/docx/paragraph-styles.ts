import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { Plugin } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import {
	contextualSpacingTwips,
	contextualStyleId,
	resolveParagraphFormatting,
	computeListLabels,
	displayListLabel,
	type DocumentModel,
	type Paragraph,
	type ParagraphFormatting,
	type TextRun,
} from 'ooxml-core/docx';
import { paragraphStyle } from './schema';
import { listMarkerDisplay } from './list-marker-display';
import { scaleMeasurer } from './run-scale';
import type { ReviewDisplayMode } from './review-display';
import { displayParagraph, ORIGINAL_PARAGRAPH_RESET } from './review-paragraph-display';
import { AttributeCache, DecorationCache } from './decoration-cache';
import { visitStyledParagraphs } from './styled-paragraphs';
import { appendInlineNode } from './run-adapter';

/** Derived display formatting stays out of document attributes and collaboration steps. */
export function paragraphStylesPlugin(
	getModel: () => DocumentModel,
	getMode: () => ReviewDisplayMode = () => 'all',
	pictureUrl?: (partName: string, contentType: string) => string | undefined,
) {
	let measurer = scaleMeasurer();
	const cache = new DecorationCache();
	const projections = new AttributeCache<ReturnType<typeof displayParagraph>>();
	const attributes = new AttributeCache<Parameters<typeof Decoration.node>[2]>();
	const formatting = new AttributeCache<ReturnType<typeof resolveParagraphFormatting>>();
	const project = (node: Parameters<typeof displayParagraph>[0], mode: ReviewDisplayMode) =>
		projections.resolve(node.attrs, [mode], () => displayParagraph(node, mode));
	return new Plugin({
		view(view) {
			const reload = () => {
				measurer = scaleMeasurer();
				cache.clear();
				view.dispatch(
					view.state.tr.setMeta('list-marker-fonts', true).setMeta('addToHistory', false),
				);
			};
			document.fonts?.addEventListener('loadingdone', reload);
			return {
				destroy() {
					document.fonts?.removeEventListener('loadingdone', reload);
				},
			};
		},
		props: {
			decorations(state) {
				const model = getModel();
				const mode = getMode();
				const inputs = [
					model,
					model.paragraphStyles,
					model.tableStyles,
					model.overrideTableStyleFontSizeAndJustification,
					model.numberingCatalog,
					model.theme,
					model.characterStyles,
					mode,
				];
				const cached = cache.get(state.doc, inputs);
				if (cached) return cached;
				const catalog = model.paragraphStyles;
				const paragraphs: Paragraph[] = [];
				const sectionBreaks = new Set(
					model.sections?.slice(0, -1).map((section) => section.endsAtBlockId),
				);
				state.doc.descendants((node) => {
					if (node.type.name !== 'paragraph') return;
					const paragraph = project(node, mode).value;
					if (sectionBreaks.has(paragraph.id) && node.childCount) {
						const runs: TextRun[] = [];
						node.forEach((child) => appendInlineNode(runs, child));
						paragraphs.push({ ...paragraph, runs });
					} else paragraphs.push(paragraph);
				});
				const labels = computeListLabels({ ...model, blocks: paragraphs });
				const decorations: Decoration[] = [];
				const styled: {
					node: ProseMirrorNode;
					pos: number;
					tableParagraph: ParagraphFormatting | undefined;
					effective: ParagraphFormatting;
					styleId: string | undefined;
				}[] = [];
				visitStyledParagraphs(state.doc, model, (node, pos, tableStyle) => {
					const tableParagraph = tableStyle?.paragraph;
					const effective = formatting.resolve(
						node.attrs,
						[
							catalog,
							JSON.stringify(tableParagraph),
							model.overrideTableStyleFontSizeAndJustification,
							mode,
						],
						() => {
							const paragraph = project(node, mode).value;
							return catalog || tableParagraph
								? resolveParagraphFormatting(
										paragraph,
										catalog,
										tableParagraph,
										model.overrideTableStyleFontSizeAndJustification,
									)
								: paragraph;
						},
					);
					const styleId = contextualStyleId(project(node, mode).value, catalog);
					styled.push({ node, pos, tableParagraph, effective, styleId });
				});
				// Contextual spacing compares a paragraph with its siblings: adjacent positions.
				const spaced = (index: number) => {
					const entry = styled[index];
					return entry && { ...entry.effective, styleId: entry.styleId };
				};
				styled.forEach(({ node, pos, tableParagraph, effective }, index) => {
					const before = styled[index - 1];
					const after = styled[index + 1];
					const contextual = contextualSpacingTwips(
						before && before.pos + before.node.nodeSize === pos ? spaced(index - 1) : undefined,
						{ ...effective, styleId: styled[index]?.styleId },
						after && pos + node.nodeSize === after.pos ? spaced(index + 1) : undefined,
					);
					const label = labels.get(String(node.attrs.id));
					const dependencies = [
						catalog,
						JSON.stringify(tableParagraph),
						model.overrideTableStyleFontSizeAndJustification,
						model.characterStyles,
						model.theme,
						mode,
						measurer,
						JSON.stringify(label),
						pictureUrl,
						label ? node.firstChild : undefined,
						contextual.beforeTwips,
						contextual.afterTwips,
					];
					const rendered = attributes.resolve(node.attrs, dependencies, () => {
						const projected = project(node, mode);
						const paragraph = projected.value;
						const reset =
							mode === 'original' && node.attrs.formatRevision && !projected.error
								? `${ORIGINAL_PARAGRAPH_RESET};`
								: '';
						const marker = listMarkerDisplay(
							node,
							{ ...paragraph, ...effective },
							label,
							model,
							measurer,
							pictureUrl,
						);
						const list = label
							? {
									listIndentLeftTwips: label.indentLeftTwips ?? null,
									listHangingTwips: label.hangingTwips ?? null,
									listFirstLineTwips: label.firstLineTwips ?? null,
								}
							: { listIndentLeftTwips: null, listHangingTwips: null, listFirstLineTwips: null };
						const spacing = {
							...(contextual.beforeTwips !== undefined
								? { spacingBeforeTwips: contextual.beforeTwips }
								: {}),
							...(contextual.afterTwips !== undefined
								? { spacingAfterTwips: contextual.afterTwips }
								: {}),
						};
						return {
							style: `${reset}${paragraphStyle({ ...effective, ...spacing, ...list })};${marker.css}`,
							...(projected.error ? { 'data-review-format-error': projected.error } : {}),
							...marker.attributes,
							...(label || node.attrs.listLabelText != null
								? {
										'data-list-label': label
											? `${displayListLabel(label.text, label.markerFormat?.fontFamily)}${label.suffix === 'space' ? ' ' : label.suffix === 'none' ? '' : '\t'}`
											: '',
									}
								: {}),
							...(effective.direction || reset ? { dir: effective.direction ?? 'ltr' } : {}),
						};
					});
					decorations.push(Decoration.node(pos, pos + node.nodeSize, rendered));
				});
				return cache.set(state.doc, inputs, DecorationSet.create(state.doc, decorations));
			},
		},
	});
}

export { resetStylePicker, syncStylePicker } from './styles-group';
