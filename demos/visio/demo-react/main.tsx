/** @jsxImportSource react */
import { useEffect, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { VisioViewer, type ViewerHandle } from '../../../viewers/visio/packages/bindings/src/react';
import { completeDemoSignInPopup } from 'ooxml-demo-onedrive';
import { createWorkspace, type Workspace } from '../demo/workspace';

// The React demo: <VisioViewer> with a ref handle.
const popup = completeDemoSignInPopup();
// In the MSAL sign-in popup the page only hands the response back to the opener.
const workspace = popup ? undefined : createWorkspace();

function App({ workspace }: { workspace: Workspace }) {
	const viewer = useRef<ViewerHandle>(null);
	useEffect(() => {
		if (viewer.current) workspace.attach(viewer.current);
	}, []);
	return (
		<VisioViewer
			ref={viewer}
			document={workspace.initialDocument}
			fileProviders={workspace.fileProviders}
			events={workspace.events}
			aria-label="Visio diagram"
		/>
	);
}

if (workspace)
	createRoot(window.document.getElementById('viewer')!).render(<App workspace={workspace} />);
