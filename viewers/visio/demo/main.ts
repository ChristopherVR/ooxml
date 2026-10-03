import { mountViewer, compatibilityNotes, compatibilityText } from '../src/index.js';
import { demoDocument } from '../src/demo-document.js';
import { wireWorkspaceShell } from './workspace-shell.js';
import { wireWorkspaceTheme } from './workspace-theme.js';

const get = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;
const fileName = get('file-name'),
	fileState = get('file-state'),
	errorBox = get('error');
const disposeTheme = wireWorkspaceTheme(document);
const revealWorkspace = wireWorkspaceShell(document, {
	browse: () => get<HTMLInputElement>('file').click(),
	sample: () => loadSample(),
});
let requestId = 0;
const viewer = mountViewer(get('viewer'), {
	document: demoDocument,
	events: {
		'document-load': () => {
			// Files opened from the viewer's own File > Open land here too.
			fileName.textContent = viewer.element.fileName || 'Untitled drawing';
			revealWorkspace();
			refreshNotes();
			viewer.fit();
		},
		'document-change': () => refreshNotes(),
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
const workspaceReport = document.querySelector<HTMLElement>('.workspace-footer')!;
workspaceReport.slot = 'workspace-footer';
viewer.element.append(workspaceReport);
function refreshEditState(): void {
	const state = viewer.controller.state;
	get('edit-label').textContent =
		state.document?.format === 'vsd'
			? 'LEGACY VSD PREVIEW'
			: !state.edit.sourceAvailable
				? 'MODEL PREVIEW'
				: state.edit.dirty
					? 'EDITED COPY'
					: 'ORIGINAL';
	fileState.textContent =
		state.document?.format === 'vsd'
			? 'Legacy VSD preview: editing and VSDX export are unavailable'
			: !state.edit.sourceAvailable
				? 'Model-only preview · Cannot save VSDX'
				: state.edit.dirty
					? 'Local file · Edited copy'
					: 'Local file · Original bytes';
}
const unsubscribeEdit = viewer.controller.subscribe(() => refreshEditState());
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
		if (!/\.vsdx?$/i.test(file.name))
			throw new Error(
				'Choose a .vsdx or supported legacy .vsd drawing. Macro-enabled files are not supported.',
			);
		if (file.size > 32 * 1024 * 1024) throw new Error('This preview accepts files up to 32 MiB.');
		await viewer.load(file);
		if (request !== requestId) return;
		fileName.textContent = file.name;
		revealWorkspace();
		refreshEditState();
		get('selection').textContent = 'Select a shape on the canvas to inspect it.';
	} catch (cause) {
		if (request !== requestId) return;
		errorBox.hidden = false;
		errorBox.textContent = cause instanceof Error ? cause.message : String(cause);
	}
}
const input = get<HTMLInputElement>('file');
input.addEventListener('change', () => {
	const file = input.files?.[0];
	if (file) void openFile(file);
	input.value = '';
});
/** The sample is offered as a template in the viewer's File > New page. */
const sampleTemplate = document.createElement('button');
sampleTemplate.type = 'button';
sampleTemplate.slot = 'templates';
sampleTemplate.className = 'template-card';
const templateTitle = document.createElement('strong');
templateTitle.textContent = 'Sample workflow';
const templateDescription = document.createElement('span');
templateDescription.textContent = 'A two-page release diagram.';
sampleTemplate.append(templateTitle, templateDescription);
sampleTemplate.addEventListener('click', () => loadSample());
viewer.element.append(sampleTemplate);
function loadSample(): void {
	++requestId;
	errorBox.hidden = true;
	viewer.update({ document: demoDocument, pageIndex: 0 });
	fileName.textContent = 'Sample workflow';
	revealWorkspace();
	refreshEditState();
	refreshNotes();
	viewer.fit();
	viewer.element.closeBackstage();
}
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
	if (event.persisted) {
		viewer.controller.cancelLoad();
		viewer.cancelEdit();
	} else {
		disposeTheme();
		unsubscribeEdit();
		viewer.destroy();
	}
});
window.addEventListener('pageshow', (event) => {
	if (event.persisted) viewer.fit();
});
refreshNotes();
requestAnimationFrame(() => viewer.fit());
if (window.parent !== window)
	window.parent.postMessage({ type: 'visio-viewer-ready' }, window.location.origin);
