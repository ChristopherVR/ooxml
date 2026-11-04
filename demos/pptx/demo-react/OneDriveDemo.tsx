import { PowerPointViewer } from 'pptx-react-viewer';
import React, { useEffect, useRef, useState } from 'react';

import { mountOneDriveDemo } from '../shared/onedrive-demo';
import type { OneDriveDemoDeck, OneDriveDemoMount } from '../shared/onedrive-demo';

/** `?onedrive=1`: open, edit and save a deck on your own OneDrive (docs/onedrive-quickstart.md). */
export function OneDriveDemoApp() {
	const bar = useRef<HTMLDivElement>(null);
	const [deck, setDeck] = useState<OneDriveDemoDeck | null>(null);
	const [storage, setStorage] = useState<OneDriveDemoMount['storage']>();
	useEffect(() => {
		const mounted = mountOneDriveDemo(bar.current!, setDeck);
		setStorage(mounted.storage);
		return () => mounted.dispose();
	}, []);
	return (
		<>
			<div ref={bar} />
			<main className='od-viewer'>
				{deck && storage ? (
					<PowerPointViewer
						content={deck.bytes}
						fileName={deck.fileName}
						canEdit={deck.canEdit}
						storage={storage}
					/>
				) : (
					<p className='od-empty'>No OneDrive file open yet.</p>
				)}
			</main>
		</>
	);
}
