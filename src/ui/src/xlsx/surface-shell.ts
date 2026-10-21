import type { EditorCore } from './editor-core';
import { mountGrid } from './grid';
import { editorStyleText } from './theme';
import { fontCommands, alignmentCommands, numberCommands } from 'ooxml-core/xlsx/ui';

/** The existing grid engine with no application chrome, dialogs or file shortcuts. */
export function buildSurface(core: EditorCore): () => void {
	const root = core.host.shadowRoot ?? core.host.attachShadow({ mode: 'open' });
	const style = core.host.ownerDocument.createElement('style');
	style.textContent = editorStyleText;
	const frame = core.host.ownerDocument.createElement('section');
	frame.className = 'xve-frame';
	const grid = core.host.ownerDocument.createElement('div');
	grid.className = 'xve-grid-host';
	grid.setAttribute('part', 'grid');
	frame.append(grid);
	root.replaceChildren(style, frame);
	core.commands.registerAll([...fontCommands(), ...alignmentCommands(), ...numberCommands()]);
	const dispose = mountGrid(core.ctx, grid);
	core.shell = {
		render: () => core.ctx.grid()?.invalidate(),
		relocalize: () => core.ctx.grid()?.invalidate(),
		toast: () => {},
	};
	return dispose;
}
