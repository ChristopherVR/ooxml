/** File-opening and menu presentation only; loading and editing remain in main/controller. */
export function wireWorkspaceShell(doc: Document): () => void {
	const get = (id: string) => doc.getElementById(id)!;
	const start = get('start-screen');
	const dropzone = get('start-dropzone');
	const menu = get('file-menu');
	const filebar = get('file-commands');
	const closeMenu = () => {
		menu.setAttribute('aria-expanded', 'false');
		filebar.removeAttribute('data-open');
	};
	const reveal = () => {
		start.hidden = true;
		doc.body.dataset.workspace = '';
		closeMenu();
	};
	get('start-browse').addEventListener('click', () => get('open').click());
	get('start-sample').addEventListener('click', () => get('sample').click());
	dropzone.addEventListener('click', (event) => {
		if (!(event.target as Element).closest('button')) get('open').click();
	});
	dropzone.addEventListener('keydown', (event) => {
		if (event.target === dropzone && (event.key === 'Enter' || event.key === ' ')) {
			event.preventDefault();
			get('open').click();
		}
	});
	menu.addEventListener('click', () => {
		const open = menu.getAttribute('aria-expanded') !== 'true';
		menu.setAttribute('aria-expanded', String(open));
		filebar.toggleAttribute('data-open', open);
	});
	doc.addEventListener('keydown', (event) => {
		if (event.key === 'Escape' && filebar.hasAttribute('data-open')) {
			closeMenu();
			menu.focus();
		}
	});
	doc.addEventListener('click', (event) => {
		const target = event.target as Node;
		if (!filebar.contains(target) && !menu.contains(target)) closeMenu();
	});
	const params = new URL(doc.defaultView!.location.href).searchParams;
	if (params.get('sample') === '1' || params.get('embed') === '1') reveal();
	return reveal;
}
