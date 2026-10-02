import { initVisioTheme } from '../docs/assets/theme.js';

/** The docs, Office launcher and workspace share the same appearance preference. */
export function wireWorkspaceTheme(doc: Document): () => void {
	const view = doc.defaultView;
	if (!view) return () => {};
	if (new URL(view.location.href).searchParams.get('embed') === '1')
		doc.documentElement.dataset.embedded = '';
	return initVisioTheme(view);
}
