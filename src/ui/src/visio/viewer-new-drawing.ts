import { editErrorMessage, isEditCancellation } from 'ooxml-core/visio/ui';

/** File New and its shortcut share the element's source lifecycle. */
export function wireNewDrawing(
	root: ShadowRoot,
	create: () => Promise<void>,
	busy: () => boolean,
	announce: (message: string) => void,
): { run(): void; dispose(): void } {
	const Abort = root.ownerDocument.defaultView?.AbortController ?? AbortController;
	const events = new Abort();
	let disposed = false;
	const run = () => {
		if (disposed || busy()) return;
		void create().catch((error: unknown) => {
			if (!disposed && !isEditCancellation(error)) announce(editErrorMessage(error));
		});
	};
	root.addEventListener(
		'keydown',
		(event) => {
			const key = event as KeyboardEvent;
			if (
				key.defaultPrevented ||
				key.isComposing ||
				key.altKey ||
				key.shiftKey ||
				!(key.ctrlKey || key.metaKey) ||
				key.key.toLowerCase() !== 'n' ||
				key
					.composedPath()
					.some(
						(node) =>
							node instanceof Element &&
							!!node.closest(
								'input, textarea, select, office-ui-select, [contenteditable]:not([contenteditable="false"])',
							),
					)
			)
				return;
			key.preventDefault();
			run();
		},
		{ signal: events.signal },
	);
	return {
		run,
		dispose() {
			disposed = true;
			events.abort();
		},
	};
}
