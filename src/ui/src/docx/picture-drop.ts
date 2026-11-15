// Dropping or pasting a picture file onto the page inserts it as an inline picture, through the
// same staging path as Insert > Pictures. Text, HTML and other files are left to the browser and
// ProseMirror, so a copied paragraph that also carries a bitmap still pastes as text.
import { Plugin, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { PICTURE_TYPES } from './picture-commands';

const DROPPABLE = new Set([...Object.keys(PICTURE_TYPES), 'image/svg+xml']);

/** The picture files a drop or paste carries, in order. */
export function pictureFiles(data: DataTransfer | null | undefined): File[] {
	return Array.from(data?.files ?? []).filter((file) => DROPPABLE.has(file.type.toLowerCase()));
}

/** Whether a drag may carry a picture file; during `dragover` only the item types are readable. */
export function dragHasPicture(data: DataTransfer | null | undefined): boolean {
	if (!data?.types.includes('Files')) return false;
	const items = Array.from(data.items ?? []);
	// Some browsers hide item types until the drop; offer the drop and filter the files then.
	return (
		!items.length ||
		items.some((item) => item.kind === 'file' && (!item.type || DROPPABLE.has(item.type)))
	);
}

export interface PictureDropHost {
	/** The editable body view, or undefined while the document cannot take a picture here. */
	editableView(): EditorView | undefined;
	/** Puts the caret at a viewport point on the pages; false to fall back to the view's own hit. */
	placeCaret?(clientX: number, clientY: number): boolean;
	insert(file: File): Promise<boolean>;
}

/** Moves the caret to the drop point in the Draft view; without geometry the caret stays put. */
function caretAtPoint(view: EditorView, clientX: number, clientY: number): boolean {
	let found: ReturnType<EditorView['posAtCoords']>;
	try {
		found = view.posAtCoords({ left: clientX, top: clientY });
	} catch {
		return false;
	}
	if (!found) return false;
	view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(found.pos))));
	return true;
}

/** Inserts the files one after another, each after the one before. */
async function insertAll(host: PictureDropHost, files: File[]): Promise<void> {
	for (const file of files) if (!(await host.insert(file))) return;
}

/**
 * Listens on the canvas, which holds both the Draft paper and the Print Layout pages, so a drop
 * anywhere on the page lands at the nearest text position. Runs in the capture phase, ahead of
 * ProseMirror's own drop handling, and stops the event so a host page's drop-to-open does not
 * also open the file.
 */
export function attachPictureDrop(canvas: HTMLElement, host: PictureDropHost): () => void {
	const listeners = new AbortController();
	const options = { capture: true, signal: listeners.signal };
	canvas.addEventListener(
		'dragover',
		(event) => {
			if (!host.editableView() || !dragHasPicture(event.dataTransfer)) return;
			event.preventDefault();
			if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy';
		},
		options,
	);
	canvas.addEventListener(
		'drop',
		(event) => {
			const view = host.editableView();
			const files = pictureFiles(event.dataTransfer);
			if (!view || !files.length) return;
			event.preventDefault();
			event.stopPropagation();
			const { clientX, clientY } = event;
			if (!host.placeCaret?.(clientX, clientY)) caretAtPoint(view, clientX, clientY);
			void insertAll(host, files);
		},
		options,
	);
	return () => listeners.abort();
}

/**
 * Pastes a copied picture file (an image with no text or HTML alongside it) as a new picture, in
 * the views `accepts` (the body; headers and notes keep the default paste).
 */
export function pictureClipboardPlugin(
	accepts: (view: EditorView) => boolean,
	insert: (file: File) => Promise<boolean>,
): Plugin {
	return new Plugin({
		props: {
			handlePaste(view, event) {
				const data = event.clipboardData;
				if (!view.editable || !data || !accepts(view)) return false;
				if (data.getData('text/html') || data.getData('text/plain')) return false;
				const files = pictureFiles(data);
				if (!files.length) return false;
				void insertAll({ editableView: () => view, insert }, files);
				return true;
			},
		},
	});
}
