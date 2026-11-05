import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-core/embedding';
import { describe, expect, it, vi } from 'vitest';

import { DEFAULT_CLOUD_FILES_LABELS } from '../../files/cloud-files-labels';
import { translationsEn } from '../i18n';
import {
	cloudFilesLabels,
	createCloudDocumentController,
	resolveCloudDocumentView,
	openPlacesLabels,
} from './cloud-open';

const provider = { id: 'onedrive', label: 'OneDrive' } as CloudFileProvider;
const storage = { save: vi.fn() };
const file = (canEdit: boolean): CloudOpenedFile => ({
	entry: { id: 'd!1', name: 'Deck.pptx', type: 'file', kind: 'pptx' },
	kind: 'pptx',
	fileName: 'Deck.pptx',
	bytes: new Uint8Array([1]),
	canEdit,
	storage,
});

describe('open places labels', () => {
	it('adds the place tab labels to the cloud browser labels', () => {
		const labels = openPlacesLabels((key) =>
			key === 'pptx.backstage.thisDevice'
				? 'Dieses Geraet'
				: key === 'pptx.backstage.places'
					? 'Orte'
					: key,
		);
		expect(labels.thisDevice).toBe('Dieses Geraet');
		expect(labels.places).toBe('Orte');
		expect(labels.signIn).toBe('Sign in');
	});
});

describe('cloud files labels', () => {
	it('has a dictionary entry for every label of the element', () => {
		for (const name of Object.keys(DEFAULT_CLOUD_FILES_LABELS)) {
			expect(translationsEn).toHaveProperty(`pptx.cloudFiles.${name}`);
		}
	});

	it('hands the element its own placeholders back', () => {
		const translate = (key: string, params?: Record<string, string | number>): string =>
			(translationsEn[key as keyof typeof translationsEn] ?? key).replace(
				/\{\{(\w+)\}\}/g,
				(_, name: string) => String(params?.[name] ?? ''),
			);
		const labels = cloudFilesLabels(translate);
		expect(labels.signInHeading).toBe('Sign in to {provider}');
		expect(labels.openingPercent).toBe('Opening {name} ({percent}%)');
		expect(labels).toEqual(DEFAULT_CLOUD_FILES_LABELS);
	});

	it('falls back to English for a missing key', () => {
		expect(cloudFilesLabels((key) => key).signIn).toBe('Sign in');
	});
});

describe('cloud document', () => {
	const host = {
		storage: undefined,
		canEdit: true,
		fileName: 'host.pptx',
		filePath: 'host/path',
	};

	it('leaves the host values alone without a cloud file', () => {
		expect(resolveCloudDocumentView(null, host)).toBe(host);
	});

	it('routes saving to the file storage and takes its name', () => {
		const view = resolveCloudDocumentView(file(true), host);
		expect(view.storage).toBe(storage);
		expect(view.fileName).toBe('Deck.pptx');
		expect(view.canEdit).toBe(true);
	});

	it('lets the stricter edit right win', () => {
		expect(resolveCloudDocumentView(file(false), host).canEdit).toBe(false);
		expect(resolveCloudDocumentView(file(true), { ...host, canEdit: false }).canEdit).toBe(false);
	});

	it('remembers the file until another document is loaded', () => {
		const controller = createCloudDocumentController();
		const listener = vi.fn();
		controller.subscribe(listener);
		const opened = file(true);
		controller.open(opened);
		expect(controller.getSnapshot().file).toBe(opened);
		controller.clear();
		expect(controller.getSnapshot().file).toBeNull();
		controller.clear();
		expect(listener).toHaveBeenCalledTimes(2);
	});
});
