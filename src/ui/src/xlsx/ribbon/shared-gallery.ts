import type { Command, RibbonControl } from 'ooxml-core/xlsx/ui';
import {
	defineGallery,
	type OfficeUiGallery,
	type OfficeGalleryPickEvent,
} from '../../ribbon/gallery';
import { tagCommand, type RenderScope } from './controls';

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
		const trigger = gallery.querySelector<HTMLButtonElement>('.trigger');
		if (trigger) tagCommand(trigger, command);
	});
	return gallery;
}
