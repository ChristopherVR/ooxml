import { mountViewer } from '../src/index.js';
import { createWorkspace } from './workspace.js';

// The vanilla demo: the element through the browser binding. The React, Vue, Angular, Svelte and
// Solid demos (packages/bindings/demos) attach the same workspace through their own bindings.
const workspace = createWorkspace();
const viewer = mountViewer(document.getElementById('viewer')!, {
	document: workspace.initialDocument,
	events: workspace.events,
});
workspace.attach(
	viewer,
	(document) => viewer.update({ document, pageIndex: 0 }),
	() => viewer.destroy(),
);
