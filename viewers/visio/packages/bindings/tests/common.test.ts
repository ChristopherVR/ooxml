import { beforeEach, describe, expect, it } from 'vitest';
import { current, mocks, reset } from './mock-binding.js';
import {
	viewerOptions,
	viewerHandle,
	propertyKeys,
	eventKeys,
	withEventEmitter,
} from '../src/common.js';
import { mountViewer } from '../src/vanilla.js';
import type { MountedViewer, ViewerProps } from '../src/common.js';
import { vi } from 'vitest';
beforeEach(reset);
describe('complete shared adapter contract', () => {
	it('copies every property, keeps null and false, omits undefined and clears removed callbacks', () => {
		const props: ViewerProps = {
			document: null,
			pageIndex: 0,
			zoom: 2,
			showToolbar: false,
			events: { 'zoom-change': vi.fn() },
		};
		expect(Object.keys(viewerOptions(props)).sort()).toEqual([...propertyKeys, 'events'].sort());
		expect(viewerOptions(props)).toEqual(props);
		expect(viewerOptions({ zoom: undefined })).toEqual({ events: {} });
	});
	it('native event forwarding covers every event and invokes callback before native output', () => {
		const received: string[] = [];
		const options = withEventEmitter(
			{ events: { 'zoom-change': () => received.push('callback') } },
			(name) => received.push(name),
		);
		expect(Object.keys(options.events!).sort()).toEqual([...eventKeys].sort());
		options.events?.['zoom-change']?.(2);
		expect(received).toEqual(['callback', 'zoom-change']);
	});
	it('guards the shared lazy handle before mount and after unmount', async () => {
		let binding: MountedViewer | undefined;
		const handle = viewerHandle(() => binding);
		expect(() => handle.element).toThrow('not mounted');
		expect(() => handle.controller).toThrow('not mounted');
		expect(() => handle.fit()).toThrow('not mounted');
		expect(() => handle.setLayerVisibility('page', 'layer', false)).toThrow('not mounted');
		expect(() => handle.resetLayerVisibility()).toThrow('not mounted');
		expect(() => handle.exportSvg()).toThrow('not mounted');
		expect(() => handle.createPrintSnapshot()).toThrow('not mounted');
		await expect(handle.load(new Uint8Array())).rejects.toThrow('not mounted');
		const host = document.createElement('div');
		binding = mountViewer(host, {});
		expect(handle.element).toBe(current().binding.element);
		handle.fit();
		handle.setLayerVisibility('page', 'layer', false);
		expect(current().binding.setLayerVisibility).toHaveBeenLastCalledWith('page', 'layer', false);
		handle.setLayerVisibility('page', 'layer', null);
		expect(current().binding.setLayerVisibility).toHaveBeenLastCalledWith('page', 'layer', null);
		handle.resetLayerVisibility('page');
		expect(current().binding.resetLayerVisibility).toHaveBeenLastCalledWith('page');
		handle.resetLayerVisibility();
		expect(current().binding.resetLayerVisibility).toHaveBeenLastCalledWith(undefined);
		expect(handle.exportSvg({ maxBytes: 1000 }).svg).toBe('<svg/>');
		expect(current().binding.exportSvg).toHaveBeenCalledWith({ maxBytes: 1000 });
		expect(handle.createPrintSnapshot({ limits: { maxPages: 1 } }).appearance).toBe(
			'saved-display',
		);
		expect(current().binding.createPrintSnapshot).toHaveBeenCalledWith({ limits: { maxPages: 1 } });
		await handle.load(new Uint8Array());
		expect(current().binding.fit).toHaveBeenCalledOnce();
		binding.destroy();
		binding = undefined;
		expect(() => handle.fit()).toThrow('not mounted');
		expect(() => handle.setLayerVisibility('page', 'layer', true)).toThrow('not mounted');
		expect(() => handle.resetLayerVisibility('page')).toThrow('not mounted');
		expect(() => handle.exportSvg()).toThrow('not mounted');
		expect(() => handle.createPrintSnapshot()).toThrow('not mounted');
	});
	it('vanilla is the exact shared mount without another renderer or lifecycle', () => {
		const host = document.createElement('div');
		const result = mountViewer(host, { zoom: 2 });
		expect(mocks.mount).toHaveBeenCalledWith(host, { zoom: 2 });
		expect(result).toBe(current().binding);
	});
});

import { mountFrameworkViewer } from '../src/common.js';
it('native prop snapshots do not overwrite loaded documents or user zoom during unrelated updates', () => {
	const a = { format: 'vsdx' as const, pages: [], diagnostics: [] };
	const b = { format: 'vsdx' as const, pages: [], diagnostics: [] };
	const callbacks = { 'zoom-change': vi.fn() };
	const viewer = mountFrameworkViewer(document.createElement('div'), { document: a, zoom: 1 });
	viewer.update({ document: a, zoom: 1, events: callbacks });
	expect(current().update).toHaveBeenLastCalledWith({ events: callbacks });
	viewer.update({ document: b, zoom: 1 });
	expect(current().update).toHaveBeenLastCalledWith({ events: {}, document: b });
	viewer.update({ document: a, zoom: 1 });
	expect(current().update).toHaveBeenLastCalledWith({ events: {}, document: a });
	viewer.update({});
	viewer.update({ document: a, zoom: 1 });
	expect(current().update).toHaveBeenLastCalledWith({ events: {}, document: a, zoom: 1 });
	viewer.destroy();
	viewer.destroy();
	expect(current().destroy).toHaveBeenCalledOnce();
	expect(() => viewer.update({})).toThrow('destroyed');
});
it('a failed older update cannot roll back a newer reentrant native snapshot', () => {
	const viewer = mountFrameworkViewer(document.createElement('div'), { zoom: 1 });
	const failure = new Error('older update failed');
	current().update.mockImplementationOnce(() => {
		viewer.update({ zoom: 3 });
		throw failure;
	});
	expect(() => viewer.update({ zoom: 2 })).toThrow(failure);
	viewer.update({ zoom: 3 });
	expect(current().update).toHaveBeenLastCalledWith({ events: {} });
	viewer.destroy();
});
it('a failed current update remains retryable against the previous native snapshot', () => {
	const viewer = mountFrameworkViewer(document.createElement('div'), { zoom: 1 });
	current().update.mockImplementationOnce(() => {
		throw new Error('current update failed');
	});
	expect(() => viewer.update({ zoom: 2 })).toThrow('current update failed');
	viewer.update({ zoom: 2 });
	expect(current().update).toHaveBeenLastCalledWith({ events: {}, zoom: 2 });
	viewer.destroy();
});
it('a throwing callback still delivers the framework-native output without changing the error', () => {
	const failure = new Error('host callback failed');
	const emit = vi.fn();
	const options = withEventEmitter(
		{
			events: {
				'zoom-change': () => {
					throw failure;
				},
			},
		},
		emit,
	);
	expect(() => options.events?.['zoom-change']?.(2)).toThrow(failure);
	expect(emit).toHaveBeenCalledWith('zoom-change', 2);
});
