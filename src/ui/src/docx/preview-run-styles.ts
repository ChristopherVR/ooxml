import {
	resolveRunFormatting,
	resolveParagraphFormatting,
	reviewParagraphFormatting,
	reviewRunFormatting,
	isRunHiddenForReview,
	type ReviewDisplayMode,
	type Block,
	type DocumentModel,
	type Paragraph,
} from 'ooxml-core/docx';
import { DOMSerializer, Fragment } from 'prosemirror-model';
import { runToInlineNodes } from './run-adapter';
import { runFormattingCss, themeFontOf } from './run-styles';
import { scaledSegments, scaleMeasurer } from './run-scale';
import { refreshPreviewScaleWithFonts } from './preview-scale-fonts';
import { paragraphStyle } from './schema';

/** Read-only story previews resolve styles for display without changing the source runs. */
export function stylePreviewRuns(
	element: HTMLElement,
	block: Block,
	model: DocumentModel,
	serializer: DOMSerializer,
	mode: ReviewDisplayMode = 'all',
): void {
	const measurer = scaleMeasurer();
	const paragraphs: Paragraph[] =
		block.type === 'paragraph'
			? [block]
			: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs));
	const elements =
		block.type === 'paragraph' ? [element] : [...element.querySelectorAll<HTMLElement>('p')];
	paragraphs.forEach((current, index) => {
		const target = elements[index];
		if (!target) return;
		const projected = reviewParagraphFormatting(current, mode);
		const paragraph = projected.value;
		const effective = model.paragraphStyles
			? resolveParagraphFormatting(paragraph, model.paragraphStyles)
			: paragraph;
		target.style.cssText = paragraphStyle({ ...paragraph, ...effective });
		if (projected.error) target.dataset.reviewFormatError = projected.error;
		target.replaceChildren();
		for (const currentRun of paragraph.runs) {
			if (isRunHiddenForReview(currentRun, mode)) continue;
			const projectedRun = reviewRunFormatting(currentRun, mode);
			const run = projectedRun.value;
			const resolved = resolveRunFormatting(run, {
				runCatalog: model.characterStyles,
				paragraphCatalog: model.paragraphStyles,
				paragraphStyleId: paragraph.style,
			});
			const contents = serializer.serializeFragment(Fragment.fromArray(runToInlineNodes(run)));
			if (
				run.text &&
				resolved.textScalePercent !== undefined &&
				resolved.textScalePercent !== 100
			) {
				const walker = document.createTreeWalker(contents, NodeFilter.SHOW_TEXT);
				const texts: Text[] = [];
				let text: Node | null;
				while ((text = walker.nextNode())) texts.push(text as Text);
				for (const text of texts) {
					const segments = scaledSegments(
						text.data,
						resolved,
						resolved.fontFamily ?? themeFontOf(resolved, model.theme),
						measurer,
					);
					if (!segments.length) continue;
					const fragment = document.createDocumentFragment();
					let from = 0;
					for (const segment of segments) {
						fragment.append(text.data.slice(from, segment.from));
						const span = document.createElement('span');
						span.className = 'dve-scaled-text';
						span.style.cssText = segment.css;
						span.textContent = text.data.slice(segment.from, segment.to);
						refreshPreviewScaleWithFonts(
							span,
							span.textContent,
							resolved,
							resolved.fontFamily ?? themeFontOf(resolved, model.theme),
						);
						fragment.append(span);
						from = segment.to;
					}
					fragment.append(text.data.slice(from));
					text.replaceWith(fragment);
				}
			}
			const css = runFormattingCss(resolved, model.theme, run);
			if (!css && !resolved.vanish && !projectedRun.error) {
				target.append(contents);
				continue;
			}
			const span = document.createElement('span');
			span.style.cssText = css;
			if (projectedRun.error) span.dataset.reviewFormatError = projectedRun.error;
			if (resolved.vanish && run.vanish === undefined) span.className = 'dve-hidden-text';
			span.append(contents);
			target.append(span);
		}
	});
}
