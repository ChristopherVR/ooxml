import { mountViewer } from '../../../viewers/visio/src/index';
import { completeDemoSignInPopup } from 'ooxml-demo-onedrive';
import { createWorkspace } from './workspace';

// The vanilla demo: the element through the browser binding. The React, Vue, Angular, Svelte and
// Solid demos (packages/bindings/demos) attach the same workspace through their own bindings.
// In the MSAL sign-in popup the page only hands the response back to the opener.
if (!completeDemoSignInPopup()) {
	const workspace = createWorkspace();
	// The vanilla binding mounts into the caller's own element, so the host's label is set here.
	const host = document.getElementById('viewer')!;
	host.setAttribute('aria-label', 'Visio diagram');
	const viewer = mountViewer(host, {
		document: workspace.initialDocument,
		fileProviders: workspace.fileProviders,
		events: workspace.events,
	});
	workspace.attach(viewer, () => viewer.destroy());
}
