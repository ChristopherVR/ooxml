/** The shared frame and dismissal behavior for a product's drawing list. */
export function selectionPaneFrame(doc: Document, dismiss: () => void) {
	const element = doc.createElement('aside');
	element.hidden = true;
	const header = doc.createElement('header');
	const heading = doc.createElement('h2');
	const close = doc.createElement('button');
	close.type = 'button';
	close.textContent = '×';
	const list = doc.createElement('ul');
	header.append(heading, close);
	element.append(header, list);
	const hide = () => {
		element.hidden = true;
		dismiss();
	};
	close.addEventListener('click', hide);
	element.addEventListener('keydown', (event) => {
		if (event.key !== 'Escape') return;
		event.preventDefault();
		event.stopPropagation();
		hide();
	});
	return { element, heading, close, list, hide };
}
