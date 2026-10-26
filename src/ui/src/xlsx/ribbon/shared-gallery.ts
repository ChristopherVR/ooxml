import type { Command, RibbonControl } from 'ooxml-core/xlsx/ui';
import {
	defineGallery,
	type OfficeUiGallery,
	type OfficeGalleryPickEvent,
	type OfficeGallerySection,
} from '../../ribbon/gallery';
import { tagCommand, type RenderScope } from './controls';

/** Adapt the immutable shared state to the older mutable Excel ribbon contract. */
export function gallerySections(sections: readonly OfficeGallerySection[]) {
	return sections.map((section) => ({
		columns: section.columns,
		tileWidth: section.tileWidth,
		tileHeight: section.tileHeight,
		...(section.title !== undefined ? { title: section.title } : {}),
		items: section.items.map((item) => ({
			id: item.id,
			label: item.label,
			...(item.preview !== undefined ? { preview: item.preview } : {}),
			...(item.applied !== undefined ? { applied: item.applied } : {}),
		})),
	}));
}

/** Thin command adapter over the shared gallery used by PowerPoint. */
export function renderSharedGallery(
	scope: RenderScope,
	control: Extract<RibbonControl, { kind: 'gallery' }>,
	command: Command,
): HTMLElement {
	const { ctx, doc } = scope;
	defineGallery(doc.defaultView?.customElements);
	const gallery = doc.createElement('office-ui-gallery') as OfficeUiGallery;
	gallery.classList.add('xve-ribbon-gallery');
	if (command.icon) gallery.setAttribute('icon', command.icon);
	tagCommand(gallery, command);
	gallery.addEventListener('office-gallery-pick', (event) => {
		const pick = event as OfficeGalleryPickEvent;
		void ctx.commands.run(command.id, pick.detail.itemId);
	});
	scope.updates.push(() => {
		gallery.state = {
			id: command.id,
			label: ctx.t(command.label),
			disabled: !ctx.commands.isEnabled(command.id),
			sections: control.sections?.(ctx) ?? [],
		};
		tagCommand(gallery, command);
		const trigger = gallery.querySelector<HTMLButtonElement>('.trigger');
		if (trigger) tagCommand(trigger, command);
	});
	return gallery;
}
