/** @jsxImportSource react */
import { useEffect, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { VisioViewer, type ViewerHandle } from '../../src/react.js';
import { createWorkspace } from '../../../../demo/workspace.js';

// The React demo: <VisioViewer> with a ref handle and the document as React state.
const workspace = createWorkspace();

function App() {
	const viewer = useRef<ViewerHandle>(null);
	const [document, setDocument] = useState(workspace.initialDocument);
	useEffect(() => {
		if (viewer.current) workspace.attach(viewer.current, setDocument);
	}, []);
	return (
		<VisioViewer
			ref={viewer}
			document={document}
			events={workspace.events}
			aria-label="Visio diagram"
		/>
	);
}

const host = window.document.getElementById('viewer')!;
createRoot(host).render(<App />);
