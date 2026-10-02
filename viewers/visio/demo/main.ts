import { mountViewer, compatibilityNotes, compatibilityText } from '../src/index.js';
import { demoDocument } from '../src/demo-document.js';

const get = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;
const fileName = get('file-name'),
	fileState = get('file-state'),
	errorBox = get('error');
let requestId = 0;
const viewer = mountViewer(get('viewer'), {
	document: demoDocument,
	events: {
		'document-load': () => {
			refreshNotes();
			viewer.fit();
		},
		'page-change': () => {
			refreshNotes();
			viewer.fit();
		},
		'shape-select': (shape) => {
			get('selection').textContent = shape
				? `${shape.name || 'Unnamed shape'}\nShape ID: ${shape.id}`
				: 'Select a shape on the canvas to inspect it.';
		},
	},
});
function refreshNotes(): void {
	const model = viewer.element.document;
	const notes = compatibilityNotes(model?.diagnostics ?? [], viewer.element.renderWarnings);
	get('notes').replaceChildren(
		...notes.map((note) => {
			const li = document.createElement('li');
			li.textContent = compatibilityText(note);
			return li;
		}),
	);
	get('note-count').textContent = String(notes.length);
	get('diagram-label').textContent = model?.pages[viewer.element.pageIndex]?.name ?? 'Diagram';
}
async function openFile(file: File): Promise<void> {
	const request = ++requestId;
	viewer.controller.cancelLoad();
	errorBox.hidden = true;
	try {
		if (!/\.vsdx$/i.test(file.name))
			throw new Error(
				'Choose a .vsdx drawing. Legacy .vsd and macro-enabled files are not supported.',
			);
		if (file.size > 32 * 1024 * 1024) throw new Error('This preview accepts files up to 32 MiB.');
		await viewer.load(file);
		if (request !== requestId) return;
		fileName.textContent = file.name;
		fileState.textContent = 'Local file · Read only';
		get('selection').textContent = 'Select a shape on the canvas to inspect it.';
	} catch (cause) {
		if (request !== requestId) return;
		errorBox.hidden = false;
		errorBox.textContent = cause instanceof Error ? cause.message : String(cause);
	}
}
const input = get<HTMLInputElement>('file');
const downloadUrls = new Set<string>();
get('export-svg').addEventListener('click', () => {
	errorBox.hidden = true;
	get('export-status').textContent = '';
	let url: string | undefined;
	const anchor = document.createElement('a');
	try {
		const result = viewer.exportSvg();
		url = URL.createObjectURL(new Blob([result.svg], { type: 'image/svg+xml;charset=utf-8' }));
		downloadUrls.add(url);
		anchor.href = url;
		anchor.download = `visio-page-${result.pageIndex + 1}.svg`;
		anchor.hidden = true;
		document.body.append(anchor);
		anchor.click();
		get('export-status').textContent =
			`SVG download requested for ${result.pageName}. This is an approximate snapshot. ${result.diagnostics.length} compatibility notes are included in the file.`;
	} catch (cause) {
		errorBox.hidden = false;
		errorBox.textContent = cause instanceof Error ? cause.message : String(cause);
	} finally {
		anchor.remove();
		if (url) {
			const created = url;
			// Keep the URL alive through browser download dispatch, then release it.
			window.setTimeout(() => {
				if (downloadUrls.delete(created)) URL.revokeObjectURL(created);
			}, 1000);
		}
	}
});
get('open').addEventListener('click', () => input.click());
input.addEventListener('change', () => {
	const file = input.files?.[0];
	if (file) void openFile(file);
	input.value = '';
});
get('sample').addEventListener('click', () => {
	++requestId;
	errorBox.hidden = true;
	viewer.update({ document: demoDocument, pageIndex: 0 });
	fileName.textContent = 'Sample workflow';
	fileState.textContent = 'Original sample scene';
	refreshNotes();
	viewer.fit();
});
const overlay = get('drop-overlay');
let dragDepth = 0;
window.addEventListener('dragenter', (event) => {
	if (event.dataTransfer?.types.includes('Files')) {
		event.preventDefault();
		++dragDepth;
		overlay.hidden = false;
	}
});
window.addEventListener('dragover', (event) => {
	if (event.dataTransfer?.types.includes('Files')) event.preventDefault();
});
window.addEventListener('dragleave', () => {
	dragDepth = Math.max(0, dragDepth - 1);
	if (!dragDepth) overlay.hidden = true;
});
window.addEventListener('dragend', () => {
	dragDepth = 0;
	overlay.hidden = true;
});
window.addEventListener('blur', () => {
	dragDepth = 0;
	overlay.hidden = true;
});
window.addEventListener('drop', (event) => {
	event.preventDefault();
	dragDepth = 0;
	overlay.hidden = true;
	const file = event.dataTransfer?.files[0];
	if (file) void openFile(file);
});
window.addEventListener('pagehide', (event) => {
	for (const url of downloadUrls) URL.revokeObjectURL(url);
	downloadUrls.clear();
	if (event.persisted) viewer.controller.cancelLoad();
	else viewer.destroy();
});
window.addEventListener('pageshow', (event) => {
	if (event.persisted) viewer.fit();
});
refreshNotes();
requestAnimationFrame(() => viewer.fit());
if (window.parent !== window)
	window.parent.postMessage({ type: 'visio-viewer-ready' }, window.location.origin);
