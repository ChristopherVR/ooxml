import { interpolate } from 'ooxml-core/i18n';
import type { Master } from './stencil-catalog';

/** One master button: drag it onto the page, or activate it to add it at the page centre. */
function masterButton(doc: Document, master: Master, t: (text: string) => string): HTMLLIElement {
	const item = doc.createElement('li');
	item.dataset.name = master.name;
	const button = doc.createElement('button');
	button.type = 'button';
	button.className = 'master';
	button.dataset.master = master.id;
	let svg = master.draw?.();
	if (!svg) {
		svg = doc.createElementNS('http://www.w3.org/2000/svg', 'svg');
		svg.setAttribute('viewBox', '0 0 24 24');
		svg.setAttribute('aria-hidden', 'true');
		const outline = doc.createElementNS('http://www.w3.org/2000/svg', 'path');
		outline.setAttribute('d', master.path);
		svg.append(outline);
	}
	const label = doc.createElement('span');
	label.textContent = master.name;
	button.append(svg, label);
	if (master.unsupported) {
		// Listed as Visio lists it, but not offered: the reason is the tooltip.
		button.dataset.unsupported = '';
		button.setAttribute('aria-disabled', 'true');
		button.title = `${master.name}: ${t(master.unsupported)}`;
	} else {
		button.draggable = true;
		button.title = interpolate(
			t('{name}: drag onto the page, or press Enter to add it at the centre.'),
			{
				name: master.name,
			},
		);
	}
	item.append(button);
	return item;
}
export function masterList(
	doc: Document,
	masters: readonly Master[],
	t: (text: string) => string,
): HTMLUListElement {
	const list = doc.createElement('ul');
	list.className = 'masters';
	list.append(...masters.map((master) => masterButton(doc, master, t)));
	return list;
}
export function menuItem(
	doc: Document,
	command: string,
	label: string,
	checked?: boolean,
): HTMLElement {
	const item = doc.createElement('office-ui-menu-item');
	item.setAttribute('command', command);
	item.setAttribute('label', label);
	if (checked !== undefined) item.setAttribute('checked', String(checked));
	return item;
}
