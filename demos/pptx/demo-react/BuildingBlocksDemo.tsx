import { SlideCanvas, SlidesPane, Toolbar, useViewerBuildingBlocks } from 'pptx-react-viewer';
import type { UseViewerBuildingBlocksInput } from 'pptx-react-viewer';
import React, { useEffect, useRef, useState } from 'react';

type CollaborationInput = UseViewerBuildingBlocksInput['collaboration'];

import { mountBuildingBlocksPanel } from '../shared/building-blocks-panel';

function Blocks({ content }: { content: Uint8Array }) {
	const [canEdit, setCanEdit] = useState(false);
	const [collaboration, setCollaboration] = useState<CollaborationInput | undefined>(undefined);
	const blocks = useViewerBuildingBlocks({
		content,
		canEdit,
		collaboration,
		fileName: 'sample-deck.pptx',
		autosaveEnabled: false,
	});
	const latest = useRef(blocks);
	latest.current = blocks;
	const panelHost = useRef<HTMLElement>(null);
	const panel = useRef<{ refresh(): void } | null>(null);
	useEffect(() => {
		if (!panelHost.current) {
			return;
		}
		panel.current = mountBuildingBlocksPanel(panelHost.current, {
			actions: () => latest.current.actions,
			state: () => ({
				mode: latest.current.mode,
				slideCount: latest.current.slidesPaneProps.slides.length,
				shellState: latest.current.shellState,
				collaborationState: latest.current.collaborationState,
			}),
			setCanEdit,
			setCollaboration: (value) => setCollaboration(value as CollaborationInput | undefined),
		});
	}, []);
	// The panel is plain DOM: refresh it after every render of the blocks.
	useEffect(() => panel.current?.refresh());
	return (
		<main style={{ height: '100vh', display: 'flex', flexDirection: 'column' }}>
			<section ref={panelHost} />
			<Toolbar {...blocks.toolbarProps} />
			<div style={{ flex: 1, minHeight: 0, display: 'flex' }}>
				<SlidesPane {...blocks.slidesPaneProps} />
				<div style={{ flex: 1, minWidth: 0, position: 'relative', display: 'flex' }}>
					<SlideCanvas {...blocks.canvasProps} />
				</div>
			</div>
		</main>
	);
}

/**
 * `?blocks=1`: a custom shell composed from React's building blocks with the shared host panel
 * that finishes their workflows. The framework-neutral `e2e/pptx/building-blocks.spec.ts` drives
 * it; every binding with building blocks has the same page.
 */
export function BuildingBlocksDemoApp() {
	const [content, setContent] = useState<Uint8Array | null>(null);
	useEffect(() => {
		void fetch(`${import.meta.env.BASE_URL}sample-deck.pptx`)
			.then((response) => response.arrayBuffer())
			.then((buffer) => setContent(new Uint8Array(buffer)));
	}, []);
	return content ? <Blocks content={content} /> : null;
}

export const buildingBlocksRequested = (): boolean =>
	new URLSearchParams(window.location.search).get('blocks') === '1';
