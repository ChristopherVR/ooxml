/**
 * The File > Open cloud place mounts the shared `pptx-ui-cloud-files` browser into its own
 * element, keeps it across option updates, swaps it for another provider and forwards opens.
 */
import {
	DestroyRef,
	ElementRef,
	Injector,
	runInInjectionContext,
	ɵChangeDetectionScheduler as ChangeDetectionScheduler,
	ɵEffectScheduler as EffectScheduler,
} from '@angular/core';
import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { CloudFilesPlaceComponent } from './cloud-files-place.component';

function provider(id: string): CloudFileProvider {
	return {
		id,
		label: id,
		account: () => null,
		signIn: vi.fn(),
		list: vi.fn(async () => ({ entries: [] })),
		open: vi.fn(),
	} as unknown as CloudFileProvider;
}

const hosts: HTMLElement[] = [];
afterEach(() => {
	for (const host of hosts.splice(0)) {
		host.remove();
	}
});

function createPlace() {
	const host = document.createElement('div');
	document.body.append(host);
	hosts.push(host);
	const destroyRef: Pick<DestroyRef, 'onDestroy'> = { onDestroy: () => () => {} };
	const place = runInInjectionContext(
		Injector.create({
			providers: [
				{ provide: ElementRef, useValue: new ElementRef(host) },
				{ provide: DestroyRef, useValue: destroyRef },
				// Same bare-injector harness as `autosave-recovery.service.test.ts`. The effect that
				// reads the (unbound) `provider` input is never flushed: `apply` is driven directly.
				{ provide: ChangeDetectionScheduler, useValue: { notify: () => {} } },
				{
					provide: EffectScheduler,
					useValue: { add: () => {}, schedule: () => {}, remove: () => {}, flush: () => {} },
				},
			],
		}),
		() => new CloudFilesPlaceComponent(),
	);
	return { host, place };
}

describe('CloudFilesPlaceComponent', () => {
	it('mounts one browser per provider and forwards an opened file', () => {
		const { host, place } = createPlace();
		const opened = vi.fn();
		place.open.subscribe(opened);
		const onedrive = provider('onedrive');
		const options = {
			provider: onedrive,
			translate: (key: string) => key,
			onOpen: (file: CloudOpenedFile) => place.open.emit(file),
		};
		place.apply(options);
		const element = host.querySelector('pptx-ui-cloud-files');
		expect(element).not.toBeNull();

		place.apply({ ...options, locale: 'fr' });
		expect(host.querySelector('pptx-ui-cloud-files')).toBe(element);

		const file = { fileName: 'Deck.pptx' } as CloudOpenedFile;
		element!.dispatchEvent(
			new CustomEvent('office-cloud-file-open', { detail: { providerId: 'onedrive', file } }),
		);
		expect(opened).toHaveBeenCalledWith(file);

		place.apply({ ...options, provider: provider('box') });
		expect(host.querySelectorAll('pptx-ui-cloud-files')).toHaveLength(1);
		expect(host.querySelector('pptx-ui-cloud-files')).not.toBe(element);
	});
});
