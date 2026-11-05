/**
 * File > Open with cloud places: the Open page lists "This device" and one place per provider
 * (from the viewer's `fileProviders`, read through `ViewerFileIOService`), shows the provider's
 * file browser, and hands a file opened there to the viewer before closing the backstage.
 */
import { Injector, runInInjectionContext } from '@angular/core';
import type { CloudFileProvider, CloudOpenedFile, OpenPlace } from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';

import { RibbonFileSectionComponent } from './ribbon-file-section.component';
import { ViewerFileIOService } from './viewer-file-io.service';

interface OpenPageInternals {
	page: { set: (id: string) => void };
	places: () => readonly OpenPlace[];
	placeId: { set: (id: string) => void };
	place: () => OpenPlace | undefined;
	cloudPlace: () => CloudFileProvider | undefined;
	openCloudFile: (file: CloudOpenedFile) => void;
}

const provider = {
	id: 'onedrive',
	label: 'OneDrive',
	account: () => null,
	signIn: vi.fn(),
	list: vi.fn(async () => ({ entries: [] })),
	open: vi.fn(),
} as unknown as CloudFileProvider;

function createFileSection(fileIO?: Partial<ViewerFileIOService>) {
	const providers = fileIO ? [{ provide: ViewerFileIOService, useValue: fileIO }] : [];
	const section = runInInjectionContext(
		Injector.create({ providers }),
		() => new RibbonFileSectionComponent(),
	);
	return { section, internals: section as unknown as OpenPageInternals };
}

describe('RibbonFileSectionComponent open places', () => {
	it('keeps the Open page as it was without providers', () => {
		const { internals } = createFileSection({ fileProviders: () => [], storage: () => undefined });
		internals.page.set('open');
		expect(internals.places()).toEqual([]);
		expect(internals.cloudPlace()).toBeUndefined();
	});

	it('lists this device and the provider, and opens a provider file', () => {
		const openCloudFile = vi.fn();
		const { section, internals } = createFileSection({
			fileProviders: () => [provider],
			storage: () => undefined,
			openCloudFile,
		});
		const closes = vi.fn();
		section.close.subscribe(closes);
		internals.page.set('open');

		expect(internals.places().map((place) => place.id)).toEqual(['device', 'cloud:onedrive']);
		expect(internals.places()[1]?.label).toBe('OneDrive');
		// "This device" is shown first, so the device search and recent list stay.
		expect(internals.place()?.id).toBe('device');
		expect(internals.cloudPlace()).toBeUndefined();

		internals.placeId.set('cloud:onedrive');
		expect(internals.cloudPlace()).toBe(provider);

		const file = { fileName: 'Deck.pptx' } as CloudOpenedFile;
		internals.openCloudFile(file);
		expect(openCloudFile).toHaveBeenCalledWith(file);
		expect(closes).toHaveBeenCalledOnce();
	});

	it('shows the cloud browser only on the Open page', () => {
		const { internals } = createFileSection({
			fileProviders: () => [provider],
			storage: () => undefined,
		});
		internals.placeId.set('cloud:onedrive');
		internals.page.set('home');
		expect(internals.cloudPlace()).toBeUndefined();
		internals.page.set('open');
		expect(internals.cloudPlace()).toBe(provider);
	});
});
