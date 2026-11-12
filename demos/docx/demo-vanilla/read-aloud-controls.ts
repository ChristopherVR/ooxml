import type { NativeHarness } from './building-blocks-types';

/**
 * A host's own Read Aloud control, drawn from `toolbarProps` of whichever binding is mounted. It
 * shows the playing state and the issue the editor reports, and never touches the editor itself.
 */
export function mountHostReadAloud(host: HTMLElement): () => void {
	document.querySelector('#host-read-aloud-panel')?.remove();
	const panel = document.createElement('div');
	panel.id = 'host-read-aloud-panel';
	const button = document.createElement('button');
	button.id = 'host-read-aloud';
	button.type = 'button';
	// Keep the document selection: pressing a host control must not take focus from the editor.
	button.addEventListener('mousedown', (event) => event.preventDefault());
	const state = document.createElement('span');
	state.id = 'host-read-aloud-state';
	const start = document.createElement('span');
	start.id = 'host-read-aloud-start';
	panel.append(button, state, start);
	host.after(panel);
	let disposed = false;
	let element: HTMLElement | null = null;
	const render = () => {
		let props: ReturnType<NativeHarness['toolbar']> | undefined;
		try {
			props = window.nativeWord?.toolbar();
		} catch {
			props = null;
		}
		if (!props) {
			button.disabled = true;
			state.textContent = 'none';
			return;
		}
		button.disabled = !props.canReadAloud;
		button.textContent = props.readingAloud ? 'Stop reading' : 'Read it aloud';
		button.setAttribute('aria-pressed', String(props.readingAloud));
		state.dataset.playing = String(props.readingAloud);
		state.dataset.issue = props.readAloudIssue ?? '';
		state.textContent = props.readingAloud ? 'reading' : (props.readAloudIssue ?? 'idle');
	};
	button.addEventListener('click', () => {
		const props = window.nativeWord?.toolbar();
		if (!props) return;
		if (props.readingAloud) {
			props.onStopReadAloud();
			start.textContent = 'stopped';
		} else {
			start.textContent = props.onReadAloud();
		}
		render();
	});
	const attach = () => {
		if (disposed) return;
		let next: HTMLElement | null = null;
		try {
			// The handle throws while no editor is mounted (before the first mount, after destroy).
			next = (window.nativeWord as NativeHarness | undefined)?.handle?.element ?? null;
		} catch {
			next = null;
		}
		if (next && next !== element) {
			element?.removeEventListener('editor-state-change', render);
			element = next;
			element.addEventListener('editor-state-change', render);
		}
		render();
		timer = window.setTimeout(attach, 100);
	};
	let timer = window.setTimeout(attach, 0);
	return () => {
		disposed = true;
		window.clearTimeout(timer);
		element?.removeEventListener('editor-state-change', render);
		panel.remove();
	};
}
