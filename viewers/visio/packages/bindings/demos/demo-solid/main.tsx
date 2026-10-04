import { createSignal } from 'solid-js';
import { render } from 'solid-js/web';
import { VisioViewer } from '../../src/solid.js';
import { createWorkspace } from '../../../../demo/workspace.js';

// The Solid demo: <VisioViewer> with a viewerRef callback and the document as a signal.
const workspace = createWorkspace();

function App() {
	const [document, setDocument] = createSignal(workspace.initialDocument);
	let attached = false;
	return (
		<VisioViewer
			document={document()}
			events={workspace.events}
			viewerRef={(handle) => {
				if (handle && !attached) {
					attached = true;
					workspace.attach(handle, setDocument);
				}
			}}
		/>
	);
}

render(() => <App />, window.document.getElementById('viewer')!);
