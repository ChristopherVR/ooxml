/**
 * Start screen presentation only. File commands live in the viewer's own File backstage, as in
 * Visio; loading and editing remain in main/controller.
 */
export function wireWorkspaceShell(
	doc: Document,
	actions: { browse(): void; sample(): void },
): () => void {
	const get = (id: string) => doc.getElementById(id)!;
	const start = get('start-screen');
	const dropzone = get('start-dropzone');
	const reveal = () => {
		start.hidden = true;
		doc.body.dataset.workspace = '';
	};
	get('start-browse').addEventListener('click', () => actions.browse());
	get('start-sample').addEventListener('click', () => actions.sample());
	dropzone.addEventListener('click', (event) => {
		if (!(event.target as Element).closest('button')) actions.browse();
	});
	dropzone.addEventListener('keydown', (event) => {
		if (event.target === dropzone && (event.key === 'Enter' || event.key === ' ')) {
			event.preventDefault();
			actions.browse();
		}
	});
	const params = new URL(doc.defaultView!.location.href).searchParams;
	if (params.get('sample') === '1' || params.get('embed') === '1') reveal();
	return reveal;
}
