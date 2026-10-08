import { initTheme } from './theme';
import './style.css';
import { createDocument } from 'docx-core';
import { normalizeEditorLocale, type DocxEditorElement } from 'docx-web-component';
import type { EditorHandle } from '../../../viewers/docx/packages/bindings/src/index';
import { mountFramework } from './framework';
import { createSampleDocument } from './sample-document';

const get = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
initTheme();

const framework = new URLSearchParams(location.search).get('framework') || 'vanilla';
get('build-stamp').textContent = `docx-viewer demo · ${framework}`;

let toastTimer: ReturnType<typeof setTimeout> | undefined;
function toast(message: string) {
	const element = get('toast');
	element.textContent = message;
	element.hidden = false;
	clearTimeout(toastTimer);
	toastTimer = setTimeout(() => (element.hidden = true), 6000);
}

/** Interface language: `?locale=de` (any supported tag), switchable from the picker. */
const localeSelect = get<HTMLSelectElement>('locale-select');
let locale = normalizeEditorLocale(new URLSearchParams(location.search).get('locale'));
localeSelect.value = locale;
localeSelect.addEventListener('change', () => {
	locale = normalizeEditorLocale(localeSelect.value);
	const url = new URL(location.href);
	url.searchParams.set('locale', locale);
	history.replaceState(null, '', url);
	for (const element of document.querySelectorAll<DocxEditorElement>('docx-editor'))
		element.locale = locale;
});

let editor: EditorHandle | undefined;
async function showEditor(): Promise<EditorHandle> {
	get('landing').hidden = true;
	get('workspace').hidden = false;
	editor ??= await mountFramework(get('editor'), {
		documentModel: createDocument(),
		locale,
		onDocumentError(error) {
			toast(error.message);
		},
	});
	editor.element.addEventListener('file-command', onShare);
	return editor;
}

/** The session this window hosts or joined, if any. */
let sessionRoom: string | undefined;
/**
 * Share (the ribbon tab row's button) has no editor default: the host brings the transport. This
 * demo hosts a session of this browser with the open document and opens a second window that
 * joins it (see session.ts), or, once sharing, just opens another window into the same session.
 */
function onShare(event: Event) {
	if ((event as CustomEvent<{ command: string }>).detail.command !== 'share') return;
	event.preventDefault();
	const element = event.currentTarget as DocxEditorElement;
	const hosting = !sessionRoom;
	const room = (sessionRoom ??= `doc-${Math.random().toString(36).slice(2, 8)}`);
	const guest = new URL(location.href);
	guest.searchParams.delete('sample');
	guest.searchParams.set('room', room);
	guest.searchParams.set('name', 'Grace');
	// Opened before any await so the browser still counts it as the click's popup.
	window.open(guest, '_blank');
	if (!hosting) return;
	void import('./session').then(({ runSession }) =>
		runSession(element, {
			room,
			host: true,
			name: 'Ada',
			onStatus: (text) => {
				get('build-stamp').textContent = `docx-viewer demo · ${framework} · ${text}`;
				toast(text);
			},
		}),
	);
}

async function openFile(file: File) {
	try {
		const handle = await showEditor();
		await handle.load(new Uint8Array(await file.arrayBuffer()));
		handle.element.fileName = file.name;
	} catch (error) {
		get('workspace').hidden = true;
		get('landing').hidden = false;
		const message = get('landing-error');
		message.textContent = error instanceof Error ? error.message : String(error);
		message.hidden = false;
	}
}

async function openModel(model: ReturnType<typeof createDocument>, fileName: string) {
	const handle = await showEditor();
	handle.element.documentModel = model;
	handle.element.fileName = fileName;
}

const landingFile = get<HTMLInputElement>('landing-file');
landingFile.addEventListener('change', () => {
	const file = landingFile.files?.[0];
	landingFile.value = '';
	if (file) void openFile(file);
});
get('browse').addEventListener('click', (event) => {
	event.stopPropagation();
	landingFile.click();
});
get('blank').addEventListener('click', (event) => {
	event.stopPropagation();
	void openModel(createDocument(), 'Document1.docx');
});
get('sample').addEventListener('click', (event) => {
	event.stopPropagation();
	void openModel(createSampleDocument(), 'Sample document.docx');
});
// `?sample=1` opens the sample document immediately (used by the docs site's embedded demo).
// `?room=<session>` joins a shared session of this browser: the window that also asks for the sample
// hosts it, the others join by name and receive the document. The editor's framework does not matter,
// so the React demo can join a session started in the Vue demo (see session.ts).
const query = new URLSearchParams(location.search);
const room = query.get('room');
if (room && /^[A-Za-z0-9_-]{1,64}$/.test(room)) {
	sessionRoom = room;
	const host = query.get('sample') === '1';
	void (async () => {
		if (host) await openModel(createSampleDocument(), 'Sample document.docx');
		else await showEditor();
		const { runSession } = await import('./session');
		runSession((await showEditor()).element, {
			room,
			host,
			name: query.get('name') || (host ? 'Ada' : 'Grace'),
			onStatus: (text) => {
				get('build-stamp').textContent = `docx-viewer demo · ${framework} · ${text}`;
			},
		});
	})();
} else if (query.get('sample') === '1') {
	void openModel(createSampleDocument(), 'Sample document.docx');
}
const dropzone = get('dropzone');
dropzone.addEventListener('click', () => landingFile.click());
dropzone.addEventListener('keydown', (event) => {
	if (event.target === dropzone && (event.key === 'Enter' || event.key === ' ')) {
		event.preventDefault();
		landingFile.click();
	}
});
const carriesFiles = (event: DragEvent) => Boolean(event.dataTransfer?.types.includes('Files'));
for (const type of ['dragenter', 'dragover'] as const)
	document.addEventListener(type, (event) => {
		if (!carriesFiles(event)) return;
		event.preventDefault();
		dropzone.classList.add('dragging');
	});
document.addEventListener('dragleave', (event) => {
	if (!event.relatedTarget) dropzone.classList.remove('dragging');
});
document.addEventListener('drop', (event) => {
	if (!carriesFiles(event)) return;
	event.preventDefault();
	dropzone.classList.remove('dragging');
	const file = event.dataTransfer?.files[0];
	if (file) void openFile(file);
});

// Signals that landing actions are wired (used by browser contracts before clicking).
document.documentElement.dataset.demoReady = 'true';
