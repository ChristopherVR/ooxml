import { interpolate } from 'ooxml-core/i18n';
import { DOCUMENT_MASTER_LIMIT, DOCUMENT_STENCIL_ID } from './shapes-document';
import { masterList } from './shapes-master-list';
import { BASIC_STENCIL_ID, type Stencil } from './stencil-catalog';

export function createStencilSection(
	doc: Document,
	{
		stencil,
		name,
		open,
		editable,
		omitted,
		t,
	}: {
		stencil: Stencil;
		name: string;
		open: boolean;
		editable: boolean;
		omitted: number;
		t: (text: string) => string;
	},
): HTMLElement {
	const { id } = stencil;
	const section = doc.createElement('section');
	section.dataset.stencil = id;
	section.setAttribute('aria-label', name);
	const header = doc.createElement('div');
	header.className = 'stencil-header';
	const title = doc.createElement('button');
	title.type = 'button';
	title.className = 'stencil-title';
	title.textContent = name;
	const list = masterList(doc, stencil.masters, t);
	list.id = `stencil-${id}-masters`;
	title.setAttribute('aria-controls', list.id);
	title.setAttribute('aria-expanded', String(open));
	list.hidden = !open;
	if (!open) section.dataset.collapsed = '';
	header.append(title);
	if (editable) {
		const actions = doc.createElement('button');
		actions.className = 'stencil-close';
		actions.type = 'button';
		actions.textContent = '⋯';
		actions.setAttribute('aria-haspopup', 'menu');
		actions.setAttribute('aria-label', t('Stencil actions'));
		actions.title = t('Stencil actions');
		actions.dataset.stencilActions = id;
		header.append(actions);
	}
	if (id !== BASIC_STENCIL_ID && id !== DOCUMENT_STENCIL_ID) {
		const close = doc.createElement('button');
		close.type = 'button';
		close.className = 'stencil-close';
		close.dataset.closeStencil = id;
		const closeLabel = interpolate(t('Close {name}'), { name });
		close.setAttribute('aria-label', closeLabel);
		close.title = closeLabel;
		close.textContent = '×';
		header.append(close);
	}
	section.append(header, list);
	if (omitted) {
		const note = doc.createElement('p');
		note.className = 'shapes-empty';
		note.hidden = !open;
		note.textContent = interpolate(
			t('Showing the first {limit} masters; {omitted} more are not listed.'),
			{ limit: DOCUMENT_MASTER_LIMIT, omitted },
		);
		section.append(note);
	}
	return section;
}
