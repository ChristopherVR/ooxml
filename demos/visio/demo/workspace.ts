import { compatibilityNotes, compatibilityText } from '../../../viewers/visio/src/index';
import type { ViewerCallbacks } from 'ooxml-ui/visio';
import type { MountedViewer } from 'ooxml-ui/visio';
import type { CloudFileProvider } from 'ooxml-core/embedding';
import type { VisioDocument } from 'ooxml-core/visio';
import { createSampleVsdx, demoDocument } from 'ooxml-core/visio/ui';
import { pickCloudFile } from 'ooxml-demo-onedrive';
import { createDemoFileProviders } from './cloud';
import { wireWorkspaceShell } from './workspace-shell';
import { wireWorkspaceTheme } from './workspace-theme';

/** What the workspace needs from a mounted viewer: any binding's handle provides it. */
export type WorkspaceViewer = Pick<
	MountedViewer,
	'element' | 'controller' | 'load' | 'fit' | 'cancelEdit'
>;

export interface Workspace {
	/** Pass these to the viewer, through the binding's own event API. */
	readonly events: ViewerCallbacks;
	/** The File > Open cloud places every demo binding passes as `fileProviders`. */
	readonly fileProviders: readonly CloudFileProvider[];
	/** What the binding mounts with; `attach` replaces it with the editable sample package. */
	readonly initialDocument: VisioDocument;
	/**
	 * Connect the mounted viewer; `destroy` tears the binding down. Every drawing, the sample
	 * included, then opens through the viewer's `load`, so the binding's `document` prop stays the
	 * initial placeholder.
	 */
	attach(viewer: WorkspaceViewer, destroy?: () => void): void;
}

/**
 * The demo workspace around the viewer: title bar, start screen, file opening, drag and drop, the
 * sample template, theme and the state values the tests read. Framework-neutral: every framework demo mounts
 * the viewer with its own binding and attaches it here, so all six demos behave the same.
 */
export function createWorkspace(doc: Document = document): Workspace {
	const fileProviders = createDemoFileProviders();
	const get = <T extends HTMLElement>(id: string): T => doc.getElementById(id) as T;
	const fileName = get('file-name'),
		fileState = get('file-state'),
		errorBox = get('error');
	const disposeTheme = wireWorkspaceTheme(doc);
	let viewer: WorkspaceViewer | undefined;
	let destroyViewer: () => void = () => {};
	let requestId = 0;
	const revealWorkspace = wireWorkspaceShell(doc, {
		browse: () => get<HTMLInputElement>('file').click(),
		sample: () => void loadSample(),
		onedrive: () => void openFromOneDrive(),
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
	/** Start screen: pick a drawing in the shared OneDrive browser, then hand it to the viewer. */
	async function openFromOneDrive(): Promise<void> {
		const provider = fileProviders[0];
		if (!provider) return;
		const file = await pickCloudFile(provider, ['vsdx']);
		if (!file || !viewer) return;
		const request = ++requestId;
		errorBox.hidden = true;
		try {
			await viewer.element.openCloudFile(file);
			if (request !== requestId) return;
			fileName.textContent = file.fileName;
			revealWorkspace();
			refreshEditState();
		} catch (cause) {
			if (request !== requestId) return;
			errorBox.hidden = false;
			errorBox.textContent = cause instanceof Error ? cause.message : String(cause);
		}
	}
	/**
	 * The sample is a real VSDX built by the core, so it is source-backed: it can be edited, saved
	 * and shared like any opened file. (The bare `demoDocument` model is only the mount placeholder.)
	 */
	async function loadSample(fromBackstage = false): Promise<void> {
		if (!viewer) return;
		const request = ++requestId;
		viewer.controller.cancelLoad();
		errorBox.hidden = true;
		try {
			const file = new File([(await createSampleVsdx()) as BlobPart], 'Sample workflow.vsdx');
			if (request !== requestId) return;
			await viewer.load(file);
			if (request !== requestId) return;
			fileName.textContent = 'Sample workflow';
			revealWorkspace();
			refreshEditState();
			refreshNotes();
			viewer.fit();
			if (fromBackstage) viewer.element.closeBackstage();
		} catch (cause) {
			if (request !== requestId) return;
			errorBox.hidden = false;
			errorBox.textContent = cause instanceof Error ? cause.message : String(cause);
		} finally {
			// Browser tests wait for this before acting, so they never race the sample load.
			doc.body.dataset.sample = request === requestId ? 'loaded' : 'superseded';
		}
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

	function attach(mounted: WorkspaceViewer, destroy: () => void = () => {}): void {
		viewer = mounted;
		destroyViewer = destroy;
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
		template.addEventListener('click', () => void loadSample(true));
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
		// `?share=<session>` joins that session on load, so frames built from different framework demos
		// can share one drawing by name. Same-browser only: without a server the session runs over a
		// BroadcastChannel. (The viewer's Share dialog does the same from an invitation link.)
		const shareRoom = new URL(view.location.href).searchParams.get('share');
		if (shareRoom) {
			const viewer = mounted.element as HTMLElement & { collaboration: { roomId: string } | null };
			viewer.collaboration = { roomId: shareRoom };
		}
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
		// A demo that starts on the sample (`?sample=1` or the embedded landing frame) replaces the
		// mount placeholder with the editable package; any file opened meanwhile wins. A window that
		// joins a share session without asking for the sample adopts the session's drawing instead.
		const params = new URL(view.location.href).searchParams;
		if (params.get('sample') === '1' || (params.get('embed') === '1' && !shareRoom))
			void loadSample();
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

	return { events, fileProviders, initialDocument: demoDocument, attach };
}
