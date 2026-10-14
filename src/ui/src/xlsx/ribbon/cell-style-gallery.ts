import type { Command, RibbonControl } from 'ooxml-core/xlsx/ui';
import type { OfficeUiGallery } from '../../ribbon/gallery';
import type { RenderScope } from './controls';
import { renderSharedGallery } from './shared-gallery';
import { refitRibbon } from './overflow';
import { tagCommand } from './controls';

type GalleryControl = Extract<RibbonControl, { kind: 'gallery' }>;
const INLINE_STYLES = ['Normal', 'Bad', 'Good', 'Neutral'];

/** Convert the existing cell-style CSS preview into the shared gallery's safe SVG format. */
function previewSvg(doc: Document, label: string, preview: string | undefined): string {
	const style = doc.createElement('span').style;
	style.cssText = preview ?? '';
	const ns = 'http://www.w3.org/2000/svg';
	const svg = doc.createElementNS(ns, 'svg');
	svg.setAttribute('xmlns', ns);
	svg.setAttribute('viewBox', '0 0 100 24');
	const fill = doc.createElementNS(ns, 'rect');
	fill.setAttribute('width', '100');
	fill.setAttribute('height', '24');
	fill.setAttribute('fill', style.backgroundColor || '#ffffff');
	const text = doc.createElementNS(ns, 'text');
	text.setAttribute('x', '5');
	text.setAttribute('y', '17');
	text.setAttribute('font-size', '12');
	text.setAttribute('font-family', 'Segoe UI, sans-serif');
	text.setAttribute('font-weight', style.fontWeight || '400');
	text.setAttribute('font-style', style.fontStyle || 'normal');
	text.setAttribute('fill', style.color || '#222222');
	text.textContent = label;
	svg.append(fill, text);
	return svg.outerHTML;
}

/** Excel's two-row Cell Styles strip, using the same gallery and picker as other products. */
export function renderCellStyleGallery(
	scope: RenderScope,
	control: GalleryControl,
	command: Command,
): HTMLElement {
	const gallery = renderSharedGallery(
		scope,
		{
			...control,
			sections: (ctx) => [
				{
					columns: 4,
					tileWidth: 100,
					tileHeight: 24,
					items: control.items(ctx).map((item) => ({
						...item,
						preview: previewSvg(scope.doc, ctx.t(item.label), item.preview),
					})),
				},
			],
		},
		command,
	) as OfficeUiGallery;
	gallery.classList.add('xve-cell-styles');
	scope.updates.push(() => {
		const state = gallery.state;
		if (!state) return;
		const items = state.sections.flatMap((section) => section.items);
		gallery.state = {
			...state,
			moreLabel: scope.ctx.t(command.label),
			inline: INLINE_STYLES.flatMap((id) => items.filter((item) => item.id === id)),
		};
		tagCommand(gallery, command);
		const trigger = gallery.querySelector<HTMLButtonElement>('.trigger');
		if (trigger) tagCommand(trigger, command);
	});
	return gallery;
}

/** Fold to the labelled Cell Styles button when the ribbon has no room for previews. */
export function wireCellStyleGalleries(root: HTMLElement): () => void {
	const fit = () => {
		const wide = root.getBoundingClientRect().width >= 1550;
		for (const gallery of root.querySelectorAll('.xve-cell-styles')) {
			if (wide) gallery.setAttribute('mode', 'inline');
			else gallery.removeAttribute('mode');
		}
		refitRibbon(root);
	};
	const Observer = root.ownerDocument.defaultView?.ResizeObserver;
	if (!Observer) return () => {};
	const observer = new Observer(fit);
	observer.observe(root);
	return () => observer.disconnect();
}
