import { render } from 'solid-js/web';
import { VisioViewer } from '../../../viewers/visio/packages/bindings/src/solid';
import { completeDemoSignInPopup } from 'ooxml-demo-onedrive';
import { createWorkspace, type Workspace } from '../demo/workspace';

// The Solid demo: <VisioViewer> with a viewerRef callback.
const popup = completeDemoSignInPopup();
// In the MSAL sign-in popup the page only hands the response back to the opener.
const workspace = popup ? undefined : createWorkspace();

function App(props: { workspace: Workspace }) {
	const { workspace } = props;
	let attached = false;
	return (
		<VisioViewer
			document={workspace.initialDocument}
			fileProviders={workspace.fileProviders}
			events={workspace.events}
			aria-label="Visio diagram"
			viewerRef={(handle) => {
				if (handle && !attached) {
					attached = true;
					workspace.attach(handle);
				}
			}}
		/>
	);
}

if (workspace)
	render(() => <App workspace={workspace} />, window.document.getElementById('viewer')!);
