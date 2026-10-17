import { afterEach, describe, expect, it, vi } from 'vitest';
import { registerViewerControls } from './office-ui';
import { RecoverDialog } from './viewer-recover-dialog';
import type { RecoveryOffer } from './viewer-autorecover';

afterEach(() => document.body.replaceChildren());

function setup(name = 'Plan.vsdx') {
	registerViewerControls();
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	const said: string[] = [];
	const dialog = new RecoverDialog(
		root,
		(message) => said.push(message),
		() => 'en-US',
	);
	const offer: RecoveryOffer = {
		name,
		savedAt: Date.UTC(2026, 0, 2, 3, 4),
		restore: vi.fn(async () => {}),
		discard: vi.fn(async () => {}),
	};
	const press = (command: string) =>
		dialog.dialog
			.querySelector(`[command="recover-${command}"]`)!
			.dispatchEvent(new CustomEvent('office-command', { detail: { command } }));
	return { dialog, offer, said, press };
}

describe('Recover Drawing dialog', () => {
	it('names the file and when the copy was kept', () => {
		const { dialog, offer } = setup();
		dialog.show(offer);
		expect(dialog.dialog.open).toBe(true);
		expect(dialog.dialog.getAttribute('heading')).toBe('Recover Drawing');
		const text = dialog.dialog.querySelector('p')!.textContent!;
		expect(text).toContain('Plan.vsdx');
		expect(text).toContain('2026');
	});

	it('shows a hostile file name as text, never as markup', () => {
		const { dialog, offer } = setup('<img src=x onerror=alert(1)>.vsdx');
		dialog.show(offer);
		expect(dialog.dialog.querySelector('img')).toBeNull();
		expect(dialog.dialog.querySelector('p')!.textContent).toContain('<img src=x onerror=alert(1)>');
	});

	it('Restore opens the kept copy and says so', async () => {
		const { dialog, offer, said, press } = setup();
		dialog.show(offer);
		press('restore');
		await vi.waitFor(() => expect(said).toHaveLength(1));
		expect(offer.restore).toHaveBeenCalledOnce();
		expect(said[0]).toMatch(/Restored Plan\.vsdx.*Save a copy/);
		expect(dialog.dialog.open).toBe(false);
	});

	it('Discard deletes the kept copy', async () => {
		const { dialog, offer, said, press } = setup();
		dialog.show(offer);
		press('discard');
		await vi.waitFor(() => expect(said).toHaveLength(1));
		expect(offer.discard).toHaveBeenCalledOnce();
		expect(offer.restore).not.toHaveBeenCalled();
		expect(said[0]).toMatch(/Discarded/);
	});

	it('Keep for later leaves the copy alone', () => {
		const { dialog, offer, said, press } = setup();
		dialog.show(offer);
		press('later');
		expect(offer.restore).not.toHaveBeenCalled();
		expect(offer.discard).not.toHaveBeenCalled();
		expect(said).toEqual([]);
		expect(dialog.dialog.open).toBe(false);
	});

	it('reports a copy that cannot be restored instead of failing silently', async () => {
		const { dialog, offer, said, press } = setup();
		offer.restore = vi.fn(async () => {
			throw new Error('The saved copy is damaged.');
		});
		dialog.show(offer);
		press('restore');
		await vi.waitFor(() => expect(said).toEqual(['The saved copy is damaged.']));
	});
});
