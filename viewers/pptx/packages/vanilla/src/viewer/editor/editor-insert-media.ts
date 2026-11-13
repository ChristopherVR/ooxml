import { mediaElementFromFile } from 'ooxml-ui/pptx';
import type { CanvasSize } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

/**
 * Insert-media factory for the vanilla editor: opens a native audio/video file
 * picker and builds the centred `media` element with the shared
 * `mediaElementFromFile` (data URL, probed video size, fixed audio strip).
 * Mirrors `pickImageElement`'s async DOM-side-effect shape in `editor-insert.ts`.
 */

/**
 * Open a native audio/video file picker, read the chosen file as a data URL,
 * size it (video: probed + capped, audio: a fixed player strip), and return a
 * centred media element. Resolves `null` when the user cancels, picks an
 * unsupported file type, or the file cannot be read.
 */
export function pickMediaElement(
	doc: Document,
	canvasSize: CanvasSize,
): Promise<PptxElement | null> {
	return new Promise((resolve) => {
		const input = doc.createElement('input');
		input.type = 'file';
		input.accept = 'video/*,audio/*';
		input.style.display = 'none';
		doc.body.appendChild(input);
		const cleanup = (): void => input.remove();
		input.addEventListener('cancel', () => {
			cleanup();
			resolve(null);
		});
		input.addEventListener('change', () => {
			const file = input.files?.[0];
			cleanup();
			if (!file) {
				resolve(null);
				return;
			}
			void mediaElementFromFile(file, canvasSize, doc).then(resolve);
		});
		input.click();
	});
}
