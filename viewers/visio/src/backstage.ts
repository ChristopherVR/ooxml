/** Visio's File (backstage) pages, in Visio's order. Save and Close act without a page. */
export const BACKSTAGE_ITEMS = [
	['info', 'Info'],
	['new', 'New'],
	['open', 'Open'],
	['save', 'Save'],
	['save-as', 'Save As'],
	['print', 'Print'],
	['share', 'Share'],
	['export', 'Export'],
	['close', 'Close'],
] as const;
export const BACKSTAGE_FOOTER = [
	['account', 'Account'],
	['feedback', 'Feedback'],
	['options', 'Options'],
] as const;
export type BackstagePage = Exclude<
	(typeof BACKSTAGE_ITEMS)[number][0] | (typeof BACKSTAGE_FOOTER)[number][0],
	'save' | 'close'
>;
const LOCAL = 'Documents stay in this browser; nothing is uploaded.';

function el<K extends keyof HTMLElementTagNameMap>(
	doc: Document,
	tag: K,
	className?: string,
	text?: string,
): HTMLElementTagNameMap[K] {
	const node = doc.createElement(tag);
	if (className) node.className = className;
	if (text !== undefined) node.textContent = text;
	return node;
}

/** A backstage command tile: large label, optional description, disabled with a reason. */
function action(
	doc: Document,
	id: string,
	label: string,
	description: string,
	unsupported?: string,
): HTMLButtonElement {
	const button = el(doc, 'button', 'backstage-action');
	button.type = 'button';
	button.dataset.backstageAction = id;
	button.append(el(doc, 'strong', '', label), el(doc, 'span', '', description));
	if (unsupported) {
		button.disabled = true;
		button.dataset.unsupported = '';
		button.title = `${label}: not available yet. ${unsupported}`;
	}
	return button;
}

function page(doc: Document, id: BackstagePage, title: string, ...content: Node[]): HTMLElement {
	const section = el(doc, 'section', 'backstage-page');
	section.dataset.backstagePage = id;
	section.hidden = true;
	section.setAttribute('aria-labelledby', `backstage-${id}-title`);
	const heading = el(doc, 'h1', '', title);
	heading.id = `backstage-${id}-title`;
	section.append(heading, ...content);
	return section;
}

/**
 * Visio's backstage: a full-window File view with a navigation column and one page per item.
 * Static application structure only; document values are filled in later as text.
 */
export function createBackstage(doc: Document): HTMLElement {
	const root = el(doc, 'div', 'backstage');
	root.hidden = true;
	root.setAttribute('role', 'dialog');
	root.setAttribute('aria-modal', 'true');
	root.setAttribute('aria-label', 'File');
	const nav = el(doc, 'nav', 'backstage-nav');
	nav.setAttribute('aria-label', 'File');
	const back = el(doc, 'button', 'backstage-back', '←');
	back.type = 'button';
	back.dataset.backstage = 'back';
	back.setAttribute('aria-label', 'Back to the drawing');
	back.title = 'Back (Esc)';
	const item = ([id, label]: readonly [string, string]) => {
		const button = el(doc, 'button', 'backstage-item', label);
		button.type = 'button';
		button.dataset.backstageItem = id;
		return button;
	};
	const footer = el(doc, 'div', 'backstage-footer');
	footer.append(...BACKSTAGE_FOOTER.map(item));
	nav.append(back, ...BACKSTAGE_ITEMS.map(item), footer);

	const facts = el(doc, 'dl', 'backstage-facts');
	for (const [key, label] of [
		['pages', 'Pages'],
		['page', 'Current page'],
		['size', 'Page size'],
		['shapes', 'Shapes'],
		['state', 'State'],
	] as const) {
		const value = el(doc, 'dd');
		value.dataset.info = key;
		facts.append(el(doc, 'dt', '', label), value);
	}
	const notes = el(doc, 'div', 'backstage-notes');
	notes.append(
		el(doc, 'h2', '', 'Compatibility'),
		el(doc, 'p', 'backstage-note-count'),
		action(doc, 'notes', 'View compatibility notes', 'What this viewer approximates or omits.'),
	);
	const file = el(doc, 'h2', 'backstage-file');
	file.dataset.info = 'name';

	const templates = el(doc, 'div', 'backstage-templates');
	const slot = doc.createElement('slot');
	slot.name = 'templates';
	templates.append(
		action(
			doc,
			'new-blank',
			'Blank drawing',
			'Start an empty drawing.',
			'Needs core blank drawing creation.',
		),
		slot,
	);
	const input = el(doc, 'input', 'backstage-file-input');
	input.type = 'file';
	input.accept = '.vsdx,.vsd';
	input.hidden = true;
	input.setAttribute('aria-label', 'Choose a Visio drawing');
	const preview = el(doc, 'div', 'backstage-print-preview');
	preview.setAttribute('aria-label', 'Print preview of the current page');

	const body = el(doc, 'div', 'backstage-body');
	body.append(
		page(doc, 'info', 'Info', file, facts, notes),
		page(doc, 'new', 'New', templates),
		page(
			doc,
			'open',
			'Open',
			action(
				doc,
				'open',
				'Browse',
				'Open a .vsdx drawing or a legacy .vsd preview from this device.',
			),
			input,
			el(doc, 'p', 'backstage-muted', `Recent drawings are not tracked. ${LOCAL}`),
		),
		page(
			doc,
			'save-as',
			'Save As',
			action(
				doc,
				'download',
				'Download a copy',
				'Save the drawing, with your edits, as a .vsdx file.',
			),
			action(
				doc,
				'export-svg',
				'Export the current page as SVG',
				'An approximate picture with compatibility notes.',
			),
		),
		page(doc, 'print', 'Print', action(doc, 'print', 'Print', 'Print the current page.'), preview),
		page(
			doc,
			'share',
			'Share',
			action(doc, 'share', 'Share with people', 'Invite people to the drawing.', LOCAL),
			action(doc, 'email', 'Email', 'Send the drawing as an attachment.', LOCAL),
		),
		page(
			doc,
			'export',
			'Export',
			action(
				doc,
				'export-svg',
				'Export the current page as SVG',
				'An approximate picture with compatibility notes.',
			),
			action(
				doc,
				'export-pdf',
				'Create PDF/XPS Document',
				'A fixed-layout document.',
				'Needs PDF output.',
			),
			action(
				doc,
				'change-type',
				'Change File Type',
				'Save in another Visio format.',
				'Needs core format conversion.',
			),
		),
		page(
			doc,
			'account',
			'Account',
			el(doc, 'p', 'backstage-muted', `No account is used. ${LOCAL}`),
		),
		page(
			doc,
			'feedback',
			'Feedback',
			el(doc, 'p', 'backstage-muted', 'Feedback is not collected by this viewer.'),
		),
		page(
			doc,
			'options',
			'Options',
			el(doc, 'p', 'backstage-muted', 'Visio Options are not available in this viewer.'),
		),
	);
	root.append(nav, body);
	return root;
}
