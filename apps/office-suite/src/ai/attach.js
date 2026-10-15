/**
 * Getting things into a chat: uploading from the device, and dropping a tab or a file onto it.
 * Office files join the library (so they can be edited and versioned) and are attached by id;
 * plain-text files are attached as notes the request reads as they are.
 */
import { icon } from '../icons.js';

/** The data type a dragged tab carries: the document id. */
export const TAB_TYPE = 'application/x-ooxml-tab';
const OFFICE = /\.(docx|xlsx|pptx|vsdx)$/iu;
const TEXT = /\.(txt|md|markdown|csv|tsv|json|xml|html?|log|ya?ml|ini|css|js|ts|py|sql)$/iu;
export const MAX_NOTE_BYTES = 512 * 1024;
export const ACCEPT = '.docx,.xlsx,.pptx,.vsdx,.txt,.md,.csv,.tsv,.json,.xml,.html,.log,.yaml,.yml';

/** Imports picked or dropped files: `{ ids, notes }` for the library files and the text notes. */
export async function importForChat(files, store) {
	const ids = [],
		notes = [];
	for (const file of files) {
		if (OFFICE.test(file.name)) {
			if (file.size > 32 * 1024 * 1024) throw new Error(`${file.name} is larger than 32 MB.`);
			const doc = await store.create(file.name, new Uint8Array(await file.arrayBuffer()));
			ids.push(doc.id);
		} else if (TEXT.test(file.name) || file.type.startsWith('text/')) {
			if (file.size > MAX_NOTE_BYTES)
				throw new Error(`${file.name} is larger than 512 KB. Attach a shorter extract.`);
			notes.push({ name: file.name.slice(0, 200), text: await file.text() });
		} else
			throw new Error(
				`${file.name} cannot be attached. Attach Word, Excel, PowerPoint or Visio files, or plain text.`,
			);
	}
	return { ids, notes };
}

/** Opens the device's file picker and resolves with the chosen files (none if cancelled). */
export function pickFiles() {
	return new Promise((resolve) => {
		const input = Object.assign(document.createElement('input'), {
			type: 'file',
			multiple: true,
			accept: ACCEPT,
		});
		input.onchange = () => resolve([...input.files]);
		input.oncancel = () => resolve([]);
		input.click();
	});
}

/** The badge of a text note. */
export const noteBadge = () => `<span class="app-badge chat-files-badge">${icon('doc')}</span>`;

/**
 * Makes `target` accept dropped tabs and files. The drop stops here, so the shell's own drop
 * handler (which opens dropped files) does not run as well.
 */
export function acceptDrops(target, { onTab, onFiles, enabled = () => true }) {
	let depth = 0;
	const wanted = (event) =>
		enabled() &&
		(event.dataTransfer?.types.includes(TAB_TYPE) || event.dataTransfer?.types.includes('Files'));
	const leave = () => {
		depth = 0;
		target.classList.remove('drop-target');
	};
	target.addEventListener('dragenter', (event) => {
		if (!wanted(event)) return;
		depth++;
		target.classList.add('drop-target');
	});
	target.addEventListener('dragover', (event) => {
		if (!wanted(event)) return;
		event.preventDefault();
		event.dataTransfer.dropEffect = 'copy';
	});
	target.addEventListener('dragleave', () => {
		if (--depth <= 0) leave();
	});
	target.addEventListener('drop', (event) => {
		if (!wanted(event)) return;
		event.preventDefault();
		event.stopPropagation();
		leave();
		const id = event.dataTransfer.getData(TAB_TYPE);
		if (id) onTab(id);
		else if (event.dataTransfer.files.length) onFiles([...event.dataTransfer.files]);
	});
}
