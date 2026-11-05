/**
 * File > Account with `fileProviders`: the connected services list gets the providers (read through
 * `ViewerFileIOService`) and translated labels, and has none without providers.
 */
import { Injector, runInInjectionContext } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import type { CloudFileProvider } from 'ooxml-ui/pptx';
import { EMPTY } from 'rxjs';
import { describe, expect, it, vi } from 'vitest';

import { AccountPageComponent } from './account-page.component';
import { ViewerFileIOService } from './viewer-file-io.service';

vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	getLocalStorageUsageSummary: async () => ({ presentationCount: 0, totalBytes: 0 }),
}));

interface AccountInternals {
	providers: () => readonly CloudFileProvider[];
	connectedLabels: () => Record<string, string>;
}

const provider = {
	id: 'onedrive',
	label: 'OneDrive',
	account: () => null,
	signIn: vi.fn(),
	list: vi.fn(async () => ({ entries: [] })),
	open: vi.fn(),
} as unknown as CloudFileProvider;

function createPage(fileIO?: Partial<ViewerFileIOService>): AccountInternals {
	const providers = [
		{
			provide: TranslateService,
			useValue: {
				instant: (key: string) => `t:${key}`,
				onLangChange: EMPTY,
				onTranslationChange: EMPTY,
			},
		},
		...(fileIO ? [{ provide: ViewerFileIOService, useValue: fileIO }] : []),
	];
	const page = runInInjectionContext(
		Injector.create({ providers }),
		() => new AccountPageComponent(),
	);
	return page as unknown as AccountInternals;
}

describe('AccountPageComponent connected services', () => {
	it('has no providers standalone or without fileProviders', () => {
		expect(createPage().providers()).toEqual([]);
		expect(createPage({ fileProviders: () => [], storage: () => undefined }).providers()).toEqual(
			[],
		);
	});

	it('hands the fileProviders and translated labels to the shared list', () => {
		const page = createPage({ fileProviders: () => [provider], storage: () => undefined });
		expect(page.providers()).toEqual([provider]);
		expect(page.connectedLabels()['heading']).toBe('t:pptx.cloudFiles.connectedHeading');
		expect(page.connectedLabels()['connect']).toBe('t:pptx.cloudFiles.connectedConnect');
	});
});
