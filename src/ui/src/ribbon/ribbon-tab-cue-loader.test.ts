import { expect, it, vi } from 'vitest';
import { RibbonTabCueLoader } from './ribbon-tab-cue-loader';

it('retries a failed tab cue chunk on the next pass and disposes the cue it wired', async () => {
	const ribbon = document.createElement('div');
	ribbon.setAttribute('overflow-label', 'More tabs');
	const dispose = vi.fn();
	const wireRibbonTabCue = vi.fn(() => dispose);
	const load = vi
		.fn<() => Promise<{ wireRibbonTabCue: typeof wireRibbonTabCue }>>()
		.mockRejectedValueOnce(new Error('chunk failed to load'))
		.mockResolvedValue({ wireRibbonTabCue });
	const loader = new RibbonTabCueLoader(ribbon, load);
	loader.ensure();
	// The next pass after the failure settles loads again; passes before it do not.
	await vi.waitFor(() => {
		loader.ensure();
		expect(load).toHaveBeenCalledTimes(2);
	});
	await vi.waitFor(() => expect(wireRibbonTabCue).toHaveBeenCalledWith(ribbon));
	loader.ensure();
	expect(load).toHaveBeenCalledTimes(2);
	loader.stop();
	expect(dispose).toHaveBeenCalledTimes(1);
});

it('loads nothing for a ribbon without an overflow label', () => {
	const load = vi.fn(() => new Promise<never>(() => {}));
	new RibbonTabCueLoader(document.createElement('div'), load).ensure();
	expect(load).not.toHaveBeenCalled();
});
