import { compatibilityNotes, compatibilityText } from '../src/index.js';
import type { ViewerCallbacks } from '../src/contract.js';
import type { MountedViewer } from '../src/binding.js';
import type { VisioDocument } from 'ooxml-core/visio';
import { demoDocument } from '../src/demo-document.js';
import { wireWorkspaceShell } from './workspace-shell.js';
import { wireWorkspaceTheme } from './workspace-theme.js';

/** What the workspace needs from a mounted viewer: any binding's handle provides it. */
export type WorkspaceViewer = Pick<
	MountedViewer,
	'element' | 'controller' | 'load' | 'fit' | 'cancelEdit'
>;

export interface Workspace {
	/** Pass these to the viewer, through the binding's own event API. */
	readonly events: ViewerCallbacks;
	/** The drawing to start with (the sample). */
	readonly initialDocument: VisioDocument;
	/**
	 * Connect the mounted viewer. `setDocument` gives the binding a new `document` property, the way
	 * that framework updates props (state, signals, inputs); `destroy` tears the binding down.
	 */
	attach(
		viewer: WorkspaceViewer,
		setDocument: (document: VisioDocument) => void,
		destroy?: () => void,
	): void;
}

/**
 * The demo workspace around the viewer: title bar, start screen, file opening, drag and drop, the
 * sample template, theme and the document report. Framework-neutral: every framework demo mounts
 * the viewer with its own binding and attaches it here, so all six demos behave the same.
 */
export function createWorkspace(doc: Document = document): Workspace {
	const get = <T extends HTMLElement>(id: string): T => doc.getElementById(id) as T;
	const fileName = get('file-name'),
		fileState = get('file-state'),
		errorBox = get('error');
	const disposeTheme = wireWorkspaceTheme(doc);
	let viewer: WorkspaceViewer | undefined;
	let setDocument: (document: VisioDocument) => void = () => {};
	let destroyViewer: () => void = () => {};
	let requestId = 0;
	const revealWorkspace = wireWorkspaceShell(doc, {
		browse: () => get<HTMLInputElement>('file').click(),
		sample: () => loadSample(),
	});

	function refreshEditState(): void {
		if (!viewer) return;
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
	function refreshNotes(): void {
		if (!viewer) return;
		const model = viewer.element.document;
		const notes = compatibilityNotes(model?.diagnostics ?? [], viewer.element.renderWarnings);
		get('notes').replaceChildren(
			...notes.map((note) => {
				const li = doc.createElement('li');
				li.textContent = compatibilityText(note);
				return li;
			}),
		);
		get('note-count').textContent = String(notes.length);
		get('diagram-label').textContent = model?.pages[viewer.element.pageIndex]?.name ?? 'Diagram';
	}
	async function openFile(file: File): Promise<void> {
		if (!viewer) return;
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
	function loadSample(): void {
		if (!viewer) return;
		++requestId;
		errorBox.hidden = true;
		// A fresh object, so every binding sees a changed `document` property.
		setDocument({ ...demoDocument });
		fileName.textContent = 'Sample workflow';
		revealWorkspace();
		refreshEditState();
		refreshNotes();
		viewer.fit();
		viewer.element.closeBackstage();
	}

	const events: ViewerCallbacks = {
		'document-load': () => {
			// Files opened from the viewer's own File > Open land here too.
			if (!viewer) return;
			fileName.textContent = viewer.element.fileName || 'Untitled drawing';
			revealWorkspace();
			refreshNotes();
			viewer.fit();
		},
		'document-change': () => refreshNotes(),
		'page-change': () => {
			refreshNotes();
			viewer?.fit();
		},
		'shape-select': (shape) => {
			get('selection').textContent = shape
				? `${shape.name || 'Unnamed shape'}\nShape ID: ${shape.id}`
				: 'Select a shape on the canvas to inspect it.';
		},
	};

	function attach(
		mounted: WorkspaceViewer,
		set: (document: VisioDocument) => void,
		destroy: () => void = () => {},
	): void {
		viewer = mounted;
		setDocument = set;
		destroyViewer = destroy;
		const report = doc.querySelector<HTMLElement>('.workspace-footer')!;
		report.slot = 'workspace-footer';
		mounted.element.append(report);
		/** The sample is offered as a template in the viewer's File > New page. */
		const template = doc.createElement('button');
		template.type = 'button';
		template.slot = 'templates';
		template.className = 'template-card';
		const title = doc.createElement('strong');
		title.textContent = 'Sample workflow';
		const description = doc.createElement('span');
		description.textContent = 'A two-page release diagram.';
		template.append(title, description);
		template.addEventListener('click', () => loadSample());
		mounted.element.append(template);
		const unsubscribe = mounted.controller.subscribe(() => refreshEditState());
		const input = get<HTMLInputElement>('file');
		input.addEventListener('change', () => {
			const file = input.files?.[0];
			if (file) void openFile(file);
			input.value = '';
		});
		wireDrop((file) => void openFile(file));
		const view = doc.defaultView!;
		view.addEventListener('pagehide', (event) => {
			if (event.persisted) {
				mounted.controller.cancelLoad();
				mounted.cancelEdit();
			} else {
				disposeTheme();
				unsubscribe();
				destroyViewer();
			}
		});
		view.addEventListener('pageshow', (event) => {
			if (event.persisted) mounted.fit();
		});
		refreshNotes();
		view.requestAnimationFrame(() => mounted.fit());
		if (view.parent !== view)
			view.parent.postMessage({ type: 'visio-viewer-ready' }, view.location.origin);
	}

	function wireDrop(open: (file: File) => void): void {
		const view = doc.defaultView!;
		const overlay = get('drop-overlay');
		let depth = 0;
		const reset = () => {
			depth = 0;
			overlay.hidden = true;
		};
		view.addEventListener('dragenter', (event) => {
			if (event.dataTransfer?.types.includes('Files')) {
				event.preventDefault();
				++depth;
				overlay.hidden = false;
			}
		});
		view.addEventListener('dragover', (event) => {
			if (event.dataTransfer?.types.includes('Files')) event.preventDefault();
		});
		view.addEventListener('dragleave', () => {
			depth = Math.max(0, depth - 1);
			if (!depth) overlay.hidden = true;
		});
		view.addEventListener('dragend', reset);
		view.addEventListener('blur', reset);
		view.addEventListener('drop', (event) => {
			event.preventDefault();
			reset();
			const file = event.dataTransfer?.files[0];
			if (file) open(file);
		});
	}

	return { events, initialDocument: demoDocument, attach };
}
