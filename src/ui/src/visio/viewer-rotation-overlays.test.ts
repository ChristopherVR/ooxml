import { afterEach, expect, it, vi } from 'vitest';
import { ViewerRotationHandle } from './viewer-rotation-handle';
import { pointerViewer } from './__fixtures__/pointer-viewer';
afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});
it('hides the saved resize overlay during rotation and restores only captured overlay nodes', async () => {
	const ui = await pointerViewer();
	ui.select();
	const overlay = document.createElementNS('http://www.w3.org/2000/svg', 'g');
	overlay.dataset.resizeOverlay = '';
	overlay.style.visibility = 'collapse';
	ui.svg.append(overlay);
	const handle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
	handle.dataset.rotationHandle = '1';
	ui.svg.append(handle);
	const rotation = new ViewerRotationHandle(ui.viewport, ui.controller, {
		active: () => true,
		announce: vi.fn(),
	});
	const dispose = rotation.wire();
	ui.pointer('pointerdown', handle, 40, 30);
	expect(overlay.style.visibility).toBe('hidden');
	const replacement = document.createElementNS('http://www.w3.org/2000/svg', 'g');
	replacement.dataset.resizeOverlay = '';
	overlay.replaceWith(replacement);
	ui.viewport.dispatchEvent(
		new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
	);
	expect(overlay.style.visibility).toBe('collapse');
	expect(replacement.style.visibility).toBe('');
	expect(ui.edits).toHaveLength(0);
	dispose();
	ui.dispose();
});
