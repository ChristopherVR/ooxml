/**
 * @vitest-environment jsdom
 *
 * jsdom supplies the sessionStorage the acknowledgements live in.
 */
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { AutosaveRecoveryOffer } from '../autosave-recovery';
import { autosaveRecoveryPrompt } from '../autosave-recovery';
import { getAcknowledgedAutosaveRecoveryTimestamp } from '../autosave-recovery-acknowledgement';
import type { AutosaveRecoveryCheckInput } from './autosave-recovery-controller';
import { createAutosaveRecoveryController } from './autosave-recovery-controller';

const NOW = 1_000_000_000;
const record = { key: 'deck.pptx', data: new Uint8Array([7]), timestamp: NOW - 60_000, size: 1 };

function offer(): AutosaveRecoveryOffer {
	const prompt = autosaveRecoveryPrompt({ record, now: NOW });
	if (!prompt) {
		throw new Error('expected a prompt');
	}
	return { prompt, record };
}

const ready: AutosaveRecoveryCheckInput = {
	filePath: 'deck.pptx',
	loading: false,
	error: null,
	slideCount: 3,
	autosaveAllowed: true,
};

const flush = () => new Promise((resolve) => setTimeout(resolve, 0));

describe('createAutosaveRecoveryController', () => {
	afterEach(() => {
		sessionStorage.clear();
	});

	it('probes once per load key and offers the prompt', async () => {
		const probe = vi.fn(async () => offer());
		const onOffered = vi.fn();
		const controller = createAutosaveRecoveryController({ load: vi.fn(), probe, onOffered });
		controller.check({ ...ready, loading: true });
		expect(probe).not.toHaveBeenCalled();
		controller.check(ready);
		controller.check(ready);
		await flush();
		expect(probe).toHaveBeenCalledTimes(1);
		expect(controller.getSnapshot().prompt?.filePath).toBe('deck.pptx');
		expect(onOffered).toHaveBeenCalledWith(record);
		controller.check({ ...ready, loadKey: 2 });
		expect(probe).toHaveBeenCalledTimes(2);
	});

	it('does not probe when the host forbids autosave', () => {
		const probe = vi.fn(async () => offer());
		const controller = createAutosaveRecoveryController({ load: vi.fn(), probe });
		controller.check({ ...ready, autosaveAllowed: false });
		expect(probe).not.toHaveBeenCalled();
	});

	it('restore loads the bytes, acknowledges them and does not re-probe its own load', async () => {
		const probe = vi.fn(async () => offer());
		let controllerRef: ReturnType<typeof createAutosaveRecoveryController> | undefined;
		const load = vi.fn(async () => {
			controllerRef?.check({ ...ready, loadKey: 'restored' });
		});
		const controller = createAutosaveRecoveryController({ load, probe });
		controllerRef = controller;
		controller.check(ready);
		await flush();
		await controller.restore();
		expect(load).toHaveBeenCalledWith(record.data);
		expect(controller.getSnapshot().prompt).toBeNull();
		expect(getAcknowledgedAutosaveRecoveryTimestamp('deck.pptx')).toBe(record.timestamp);
		expect(probe).toHaveBeenCalledTimes(1);
	});

	it('a failed load forgets the acknowledgement', async () => {
		const controller = createAutosaveRecoveryController({
			load: async () => {
				throw new Error('bad');
			},
			probe: async () => offer(),
		});
		controller.check(ready);
		await flush();
		await expect(controller.restore()).rejects.toThrow('bad');
		expect(getAcknowledgedAutosaveRecoveryTimestamp('deck.pptx')).toBeUndefined();
	});

	it('discard deletes the record; a failed delete keeps the prompt for a retry', async () => {
		const discardRecord = vi
			.fn<(r: typeof record) => Promise<void>>()
			.mockRejectedValueOnce(new Error('blocked'))
			.mockResolvedValueOnce(undefined);
		const controller = createAutosaveRecoveryController({
			load: vi.fn(),
			probe: async () => offer(),
			discardRecord,
		});
		controller.check(ready);
		await flush();
		await controller.discard();
		expect(controller.getSnapshot()).toMatchObject({ discarding: false });
		expect(controller.getSnapshot().prompt).not.toBeNull();
		await controller.discard();
		expect(discardRecord).toHaveBeenCalledTimes(2);
		expect(controller.getSnapshot().prompt).toBeNull();
	});

	it('drops a stale probe answer when a newer load probed meanwhile', async () => {
		let first: ((value: AutosaveRecoveryOffer | null) => void) | undefined;
		const probe = vi
			.fn()
			.mockImplementationOnce(
				() =>
					new Promise((resolve) => {
						first = resolve;
					}),
			)
			.mockResolvedValueOnce(null);
		const controller = createAutosaveRecoveryController({ load: vi.fn(), probe });
		controller.check({ ...ready, loadKey: 1 });
		controller.check({ ...ready, loadKey: 2 });
		first?.(offer());
		await flush();
		expect(controller.getSnapshot().prompt).toBeNull();
	});
});
