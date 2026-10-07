import { Plugin } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import {
	resolveParagraphFormatting,
	computeListLabels,
	displayListLabel,
	type DocumentModel,
	type Paragraph,
} from 'ooxml-core/docx';
import { paragraphStyle } from './schema';
import { listMarkerDisplay } from './list-marker-display';
import { scaleMeasurer } from './run-scale';
import type { ReviewDisplayMode } from './review-display';
import { displayParagraph, ORIGINAL_PARAGRAPH_RESET } from './review-paragraph-display';

/** Derived display formatting stays out of document attributes and collaboration steps. */
export function paragraphStylesPlugin(
	getModel: () => DocumentModel,
	getMode: () => ReviewDisplayMode = () => 'all',
) {
	let measurer = scaleMeasurer();
	return new Plugin({
		view(view) {
			const reload = () => {
				measurer = scaleMeasurer();
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
				const catalog = model.paragraphStyles;
				const paragraphs: Paragraph[] = [];
				state.doc.descendants((node) => {
					if (node.type.name === 'paragraph') paragraphs.push(displayParagraph(node, mode).value);
				});
				const labels = computeListLabels({ ...model, blocks: paragraphs });
				const decorations: Decoration[] = [];
				state.doc.descendants((node, pos) => {
					if (node.type.name !== 'paragraph') return;
					const projected = displayParagraph(node, mode);
					const paragraph = projected.value;
					const reset =
						mode === 'original' && node.attrs.formatRevision && !projected.error
							? `${ORIGINAL_PARAGRAPH_RESET};`
							: '';
					const effective = catalog ? resolveParagraphFormatting(paragraph, catalog) : paragraph;
					const label = labels.get(String(node.attrs.id));
					const marker = listMarkerDisplay(
						node,
						{ ...paragraph, ...effective },
						label,
						model,
						measurer,
					);
					const list = label
						? {
								listIndentLeftTwips: label.indentLeftTwips ?? null,
								listHangingTwips: label.hangingTwips ?? null,
								listFirstLineTwips: label.firstLineTwips ?? null,
							}
						: { listIndentLeftTwips: null, listHangingTwips: null, listFirstLineTwips: null };
					decorations.push(
						Decoration.node(pos, pos + node.nodeSize, {
							style: `${reset}${paragraphStyle({ ...effective, ...list })};${marker.css}`,
							...(projected.error ? { 'data-review-format-error': projected.error } : {}),
							...marker.attributes,
							...(label || node.attrs.listLabelText != null
								? {
										'data-list-label': label
											? `${displayListLabel(label.text)}${label.suffix === 'space' ? ' ' : label.suffix === 'none' ? '' : '\t'}`
											: '',
									}
								: {}),
							...(effective.direction || reset ? { dir: effective.direction ?? 'ltr' } : {}),
						}),
					);
				});
				return DecorationSet.create(state.doc, decorations);
			},
		},
	});
}

export { resetStylePicker, syncStylePicker } from './styles-group';
