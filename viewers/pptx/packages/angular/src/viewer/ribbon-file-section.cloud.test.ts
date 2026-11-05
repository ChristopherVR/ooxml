/**
 * File > Open with cloud places: the Open page renders the shared places element with the providers
 * (from the viewer's `fileProviders`, read through `ViewerFileIOService`), and hands a file opened there to the viewer before closing the backstage.
 */
import { Injector, runInInjectionContext } from '@angular/core';
import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';

import { RibbonFileSectionComponent } from './ribbon-file-section.component';
import { ViewerFileIOService } from './viewer-file-io.service';

interface OpenPageInternals {
	page: { set: (id: string) => void };
	placeProviders: () => readonly CloudFileProvider[];
	placeLabels: () => Record<string, string>;
	openAccept: readonly string[];
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
		expect(internals.placeProviders()).toEqual([]);
	});

	it('hands the providers to the places element and opens a provider file', () => {
		const openCloudFile = vi.fn();
		const { section, internals } = createFileSection({
			fileProviders: () => [provider],
			storage: () => undefined,
			openCloudFile,
		});
		const closes = vi.fn();
		section.close.subscribe(closes);
		internals.page.set('open');

		expect(internals.placeProviders()).toEqual([provider]);
		expect(internals.openAccept).toEqual(['pptx']);
		expect(internals.placeLabels()['thisDevice']).toBe('pptx.backstage.thisDevice');

		const file = { fileName: 'Deck.pptx' } as CloudOpenedFile;
		internals.openCloudFile(file);
		expect(openCloudFile).toHaveBeenCalledWith(file);
		expect(closes).toHaveBeenCalledOnce();
	});

	it('offers the providers on the Open page only', () => {
		const { internals } = createFileSection({
			fileProviders: () => [provider],
			storage: () => undefined,
		});
		internals.page.set('home');
		expect(internals.placeProviders()).toEqual([]);
		internals.page.set('open');
		expect(internals.placeProviders()).toEqual([provider]);
	});
});
