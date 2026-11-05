/**
 * Demo-only: "Open from OneDrive" for a demo's start page, where File > Open is not reachable yet
 * because no document is open. It shows the shared `<office-ui-cloud-files>` browser with the
 * demo's own provider in a modal dialog that follows the page theme. A real host builds this from
 * its own UI and calls the editor's `openCloudFile(file)` with what the user picked.
 */
import type { CloudFileKind, CloudFileProvider, CloudOpenedFile } from 'ooxml-core/embedding';
import { defineCloudFiles } from 'ooxml-ui/controls';
import { installOfficeUiThemeFor } from 'ooxml-ui/theme';

const STYLE = `
.ooxml-demo-cloud { --bg: #fff; --fg: #242424; --line: #e0e0e0; --accent: #0f6cbd; --accent-fg: #fff;
	position: fixed; inset: 0; margin: auto; padding: 0; box-sizing: border-box; display: flex;
	flex-direction: column; width: min(720px, calc(100vw - 32px)); height: min(560px, calc(100vh - 32px));
	border: 1px solid var(--line); border-radius: 8px; overflow: hidden; color-scheme: light;
	font: 14px/1.45 'Segoe UI', system-ui, sans-serif; color: var(--fg); background: var(--bg); }
.ooxml-demo-cloud[open] { display: flex; }
.ooxml-demo-cloud.dark { --bg: #1f1f1f; --fg: #f3f3f3; --line: #3d3d3d; --accent: #479ef5;
	--accent-fg: #0a0a0a; color-scheme: dark; }
.ooxml-demo-cloud::backdrop { background: rgb(0 0 0 / 0.45); }
.ooxml-demo-cloud header { display: flex; align-items: center; justify-content: space-between;
	padding: 12px 16px; border-bottom: 1px solid var(--line); }
.ooxml-demo-cloud h2 { margin: 0; font-size: 16px; font-weight: 600; }
.ooxml-demo-cloud button { font: inherit; padding: 4px 12px; border-radius: 4px; cursor: pointer;
	border: 1px solid var(--accent); background: transparent; color: var(--accent); }
.ooxml-demo-cloud office-ui-cloud-files { flex: 1; min-height: 0; }`;

/** Whether the page shows a dark theme: the demo toggle, an editor's theme, else the system. */
function prefersDark(): boolean {
	const root = document.documentElement;
	const explicit = root.dataset.theme ?? root.getAttribute('data-office-theme');
	if (explicit === 'dark') return true;
	if (explicit === 'light') return false;
	return matchMedia('(prefers-color-scheme: dark)').matches;
}

/** Resolves with the file the user opened, or undefined when the dialog was closed. */
export function pickCloudFile(
	provider: CloudFileProvider,
	accept: readonly CloudFileKind[],
	title = 'Open from OneDrive',
): Promise<CloudOpenedFile | undefined> {
	defineCloudFiles();
	return new Promise((resolve) => {
		const dark = prefersDark();
		const dialog = document.createElement('dialog');
		dialog.className = dark ? 'ooxml-demo-cloud dark' : 'ooxml-demo-cloud';
		dialog.setAttribute('data-office-theme', dark ? 'dark' : 'light');
		dialog.setAttribute('aria-label', title);
		const style = document.createElement('style');
		style.textContent = STYLE;
		const header = document.createElement('header');
		const heading = document.createElement('h2');
		heading.textContent = title;
		const close = document.createElement('button');
		close.type = 'button';
		close.textContent = 'Close';
		close.addEventListener('click', () => dialog.close());
		header.append(heading, close);
		const browser = document.createElement('office-ui-cloud-files') as HTMLElement & {
			providers: readonly CloudFileProvider[];
			accept: readonly CloudFileKind[];
		};
		browser.providers = [provider];
		browser.accept = accept;
		let result: CloudOpenedFile | undefined;
		browser.addEventListener('office-cloud-file-open', (event) => {
			result = (event as CustomEvent<{ file: CloudOpenedFile }>).detail.file;
			dialog.close();
		});
		dialog.addEventListener('close', () => {
			dialog.remove();
			resolve(result);
		});
		dialog.append(style, header, browser);
		document.body.append(dialog);
		installOfficeUiThemeFor(dialog);
		dialog.showModal();
	});
}
