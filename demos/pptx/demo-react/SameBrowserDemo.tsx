import { PowerPointViewer } from 'pptx-react-viewer';
import React, { useEffect, useState } from 'react';

import { createSameBrowserDemo } from '../shared/same-browser-collaboration';
import type { SameBrowserDemo } from '../shared/same-browser-collaboration';

export function SameBrowserDemoApp() {
	const [demo, setDemo] = useState<SameBrowserDemo | null>(null);
	const [error, setError] = useState('');
	useEffect(() => {
		let disposed = false;
		let current: SameBrowserDemo | undefined;
		void createSameBrowserDemo()
			.then((value) => {
				if (disposed) {
					value.dispose();
					return;
				}
				current = value;
				setDemo(value);
				return undefined;
			})
			.catch((reason: unknown) => setError(String(reason)));
		return () => {
			disposed = true;
			current?.dispose();
		};
	}, []);
	return (
		<main style={{ position: 'fixed', inset: '28px 0 0' }}>
			{error && <p role='alert'>{error}</p>}
			{demo && (
				<PowerPointViewer
					content={demo.source}
					fileName={demo.fileName}
					collaboration={demo.collaboration}
					onCollaborationChange={demo.report}
					canEdit
				/>
			)}
		</main>
	);
}
