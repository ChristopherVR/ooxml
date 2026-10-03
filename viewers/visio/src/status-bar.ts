/**
 * Visio's status bar on the shared `office-ui-status-bar`: page and shape metrics, a live
 * message region, compatibility notes, the host footer slot and the shared zoom slider.
 */
export function createStatusBar(doc: Document): HTMLElement {
	const bar = doc.createElement('office-ui-status-bar');
	bar.className = 'status';
	bar.setAttribute('label', 'Status bar');
	const item = (name: string) => {
		const el = doc.createElement('office-ui-status-item');
		el.setAttribute(name, '');
		el.setAttribute('label', '');
		el.setAttribute('value', '');
		return el;
	};
	const message = doc.createElement('div');
	message.className = 'status-message';
	message.setAttribute('role', 'status');
	const text = doc.createElement('span');
	text.dataset.status = '';
	const diagnostics = doc.createElement('span');
	diagnostics.dataset.diagnostics = '';
	message.append(text, diagnostics);
	const notes = doc.createElement('div');
	notes.className = 'notes-strip';
	const notesButton = doc.createElement('button');
	notesButton.type = 'button';
	notesButton.dataset.chrome = 'notes';
	notesButton.setAttribute('aria-controls', 'inspector-pane');
	notesButton.setAttribute('aria-expanded', 'false');
	const count = doc.createElement('span');
	count.className = 'notes-count';
	count.dataset.noteCount = '';
	notesButton.append('Notes', count);
	const local = doc.createElement('span');
	local.textContent = 'Files stay in your browser';
	notes.append(notesButton, local);
	const footer = doc.createElement('slot');
	footer.name = 'workspace-footer';
	const zoom = doc.createElement('office-ui-zoom-slider');
	zoom.slot = 'end';
	zoom.className = 'zoom-controls';
	zoom.setAttribute('label', 'Canvas zoom');
	zoom.setAttribute('fit', 'Fit page to current window');
	zoom.setAttribute('value', '100');
	bar.append(item('data-page-status'), item('data-shape-status'), message, notes, footer, zoom);
	return bar;
}

/** Visio's page tabs under the drawing, on the shared document tab strip. */
export function createPageTabs(doc: Document): HTMLElement {
	const strip = doc.createElement('office-ui-tab-strip');
	strip.className = 'page-tabs';
	strip.setAttribute('label', 'Pages');
	strip.setAttribute('previous-label', 'Previous page');
	strip.setAttribute('next-label', 'Next page');
	return strip;
}
