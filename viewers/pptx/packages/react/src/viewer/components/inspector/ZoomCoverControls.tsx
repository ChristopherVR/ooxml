import type {
	PptxUiZoomCoverElement,
	ZoomCoverControlState,
	ZoomCoverPatchEvent,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import React, { useEffect, useRef } from 'react';

interface Props extends ZoomCoverControlState {
	onPatch: (patch: Partial<PptxElement>) => void;
}

export function ZoomCoverControls({
	element,
	disabled,
	translate,
	onPatch,
}: Props): React.JSX.Element {
	const ref = useRef<PptxUiZoomCoverElement>(null);
	useEffect(() => {
		if (ref.current) ref.current.state = { element, disabled, translate };
	}, [element, disabled, translate]);
	useEffect(() => {
		const control = ref.current;
		if (!control) return;
		const commit = (event: Event) => onPatch((event as ZoomCoverPatchEvent).detail);
		control.addEventListener('zoom-cover-patch', commit);
		return () => control.removeEventListener('zoom-cover-patch', commit);
	}, [onPatch]);
	return <pptx-ui-zoom-cover ref={ref} />;
}
