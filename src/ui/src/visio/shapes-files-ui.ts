import type { Stencil } from './stencil-catalog';
import {
	MAX_STENCIL_BYTES,
	StencilFileError,
	stencilErrorMessage,
	type StencilFiles,
} from './stencil-files';
import { loadStencilStrings } from './stencil-locale';

/** What a stencil file picker offers; Visio's own stencils are .vssx, and the rest is refused by name. */
export const STENCIL_FILE_ACCEPT = '.vssx,.vssm,.vstx,.vstm,.vsdx,.vsdm,.vss,.vst';

/**
 * The hidden file input behind More Shapes > Open Stencil and the line that says why a file could
 * not be opened. `opened` runs for each stencil that was read.
 */
export function stencilFilePicker(
	doc: Document,
	files: StencilFiles,
	opened: (stencil: Stencil) => void,
	t: (text: string) => string,
) {
	const input = doc.createElement('input');
	input.type = 'file';
	input.accept = STENCIL_FILE_ACCEPT;
	input.multiple = true;
	input.hidden = true;
	input.dataset.openStencil = '';
	input.setAttribute('aria-label', 'Open Stencil');
	const status = doc.createElement('p');
	status.className = 'stencil-status';
	status.setAttribute('role', 'status');
	status.dataset.stencilStatus = '';
	status.hidden = true;
	input.addEventListener('change', async () => {
		const picked = Array.from(input.files ?? []).slice(0, 16);
		input.value = '';
		status.hidden = true;
		for (const file of picked) {
			try {
				await loadStencilStrings();
				if (file.size > MAX_STENCIL_BYTES)
					throw new StencilFileError(file.name, 'This file is too large to open.');
				opened(await files.open(file.name, new Uint8Array(await file.arrayBuffer())));
			} catch (error) {
				status.textContent = stencilErrorMessage(error, t);
				status.hidden = false;
			}
		}
	});
	return { input, status, open: () => input.click() };
}
