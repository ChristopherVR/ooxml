// @vitest-environment jsdom
import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-core/embedding';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { OfficeUiCloudFiles } from '../../controls';
import { mountCloudFilesPlace } from './cloud-files-place';

afterEach(() => {
	document.body.replaceChildren();
});

const provider = {
	id: 'onedrive',
	label: 'OneDrive',
	account: () => null,
	signIn: vi.fn(),
	list: vi.fn(),
	open: vi.fn(),
} as unknown as CloudFileProvider;
const translate = (key: string): string => (key === 'pptx.cloudFiles.signIn' ? 'Anmelden' : key);

describe('mountCloudFilesPlace', () => {
	it('shows one provider for pptx files with translated labels', () => {
		const host = document.createElement('div');
		document.body.append(host);
		mountCloudFilesPlace(host, { provider, translate, locale: 'de', onOpen: vi.fn() });
		const element = host.querySelector('pptx-ui-cloud-files') as OfficeUiCloudFiles;
		expect(element.providers).toEqual([provider]);
		expect(element.accept).toEqual(['pptx']);
		expect(element.labels.signIn).toBe('Anmelden');
		expect(element.locale).toBe('de');
	});

	it('hands an opened file to the latest callback and unmounts cleanly', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const first = vi.fn();
		const second = vi.fn();
		const place = mountCloudFilesPlace(host, { provider, translate, onOpen: first });
		place.update({ provider, translate, onOpen: second });
		const file = { fileName: 'a.pptx' } as CloudOpenedFile;
		host
			.querySelector('pptx-ui-cloud-files')!
			.dispatchEvent(
				new CustomEvent('office-cloud-file-open', { detail: { providerId: 'onedrive', file } }),
			);
		expect(first).not.toHaveBeenCalled();
		expect(second).toHaveBeenCalledWith(file);
		place.destroy();
		expect(host.querySelector('pptx-ui-cloud-files')).toBeNull();
	});
});
