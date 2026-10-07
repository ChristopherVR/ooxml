import type { VisioDocument, VisioPage, VisioShape } from 'ooxml-core/visio';
import { visioRotationPreviewTransform } from 'ooxml-core/visio/ui';
import { renderPage, svgElement, type RenderResult } from './render-svg';
/** Render the temporary pose through the same geometry, paint, text and resource pipeline. */
export function createRotationPreview(
	model: VisioDocument,
	page: VisioPage,
	shape: VisioShape,
	source: SVGGElement,
) {
	const element = svgElement('g');
	element.classList.add('rotation-shape-preview');
	element.setAttribute('aria-hidden', 'true');
	element.style.pointerEvents = 'none';
	const visibility = source.style.visibility;
	let result: RenderResult | undefined;
	source.after(element);
	source.style.visibility = 'hidden';
	return {
		element,
		update(angle: number): void {
			const transform = visioRotationPreviewTransform(shape, angle);
			if (!transform) return;
			const previewPage = {
				...page,
				shapes: [{ ...shape, transform, rotation: { ...shape.rotation!, angle } }],
			};
			delete previewPage.backgroundPageId;
			const next = renderPage({ ...model, pages: [previewPage] }, previewPage, {
				interactive: false,
			});
			const defs = next.svg.querySelector(':scope > defs')!;
			const root = next.svg.querySelector(':scope > g')!;
			element.replaceChildren(defs, ...root.children);
			result?.dispose();
			result = next;
		},
		dispose(): void {
			element.remove();
			result?.dispose();
			result = undefined;
			source.style.visibility = visibility;
		},
	};
}
