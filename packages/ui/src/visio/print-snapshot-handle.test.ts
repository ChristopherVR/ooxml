import { afterEach, describe, expect, it, vi } from 'vitest';
import { mountViewer } from './binding.js';
import { ViewerController } from './controller.js';
import { demoDocument } from './demo-document.js';
import * as snapshots from './print-snapshot.js';

afterEach(() => {
	vi.restoreAllMocks();
	document.body.replaceChildren();
});
function duringSerialization(action: () => void): void {
	const serialize = XMLSerializer.prototype.serializeToString;
	vi.spyOn(XMLSerializer.prototype, 'serializeToString').mockImplementationOnce((node) => {
		const result = serialize.call(new XMLSerializer(), node);
		action();
		return result;
	});
}
describe('current-page print snapshot handle', () => {
	it('forwards through the standalone implementation without changing viewer state, DOM or events', () => {
		const viewer = mountViewer(document.createElement('div'), {
			document: demoDocument,
			pageIndex: 1,
			zoom: 3,
		});
		viewer.controller.setSearchQuery('framework');
		viewer.controller.selectShape({ id: 'a1', name: 'Your framework', pageId: '2' });
		const events = vi.fn();
		viewer.controller.onEvent(events);
		const state = viewer.controller.state,
			before = viewer.element.shadowRoot!.innerHTML;
		const create = vi.spyOn(snapshots, 'createPrintSnapshot');
		const result = viewer.createPrintSnapshot({ limits: { maxPages: 1 } });
		expect(create).toHaveBeenCalledExactlyOnceWith(demoDocument, {
			pageIndices: [1],
			limits: { maxPages: 1 },
		});
		expect(result.pageIndices).toEqual([1]);
		expect(result.pages[0]!.pageName).toBe('Architecture');
		expect(viewer.controller.state).toBe(state);
		expect(viewer.element.shadowRoot!.innerHTML).toBe(before);
		expect(events).not.toHaveBeenCalled();
		expect(viewer.element.createPrintSnapshot().pages).toHaveLength(1);
		expect(() => viewer.createPrintSnapshot({ limits: { maxTotalBytes: 1 } })).toThrow();
		expect(viewer.controller.state).toBe(state);
		expect(viewer.element.shadowRoot!.innerHTML).toBe(before);
		viewer.destroy();
		expect(() => viewer.createPrintSnapshot()).toThrow('destroyed');
		expect(() => viewer.element.createPrintSnapshot()).toThrow('destroyed');
	});
	it('rejects empty/unloaded documents and invalid handle options without accepting page overrides', () => {
		const viewer = mountViewer(document.createElement('div'));
		expect(() => viewer.createPrintSnapshot()).toThrow('Open a document');
		viewer.update({ document: { format: 'vsdx', pages: [], diagnostics: [] } });
		expect(() => viewer.createPrintSnapshot()).toThrow('page index');
		viewer.update({ document: demoDocument });
		for (const options of [null, [], { pageIndices: [1] }, { scale: 0.5 }])
			expect(() =>
				viewer.createPrintSnapshot(options as snapshots.CurrentPagePrintSnapshotOptions),
			).toThrow(/options|option/);
		viewer.destroy();
	});
	it.each(['replace', 'same', 'clear', 'destroy', 'controller-destroy', 'delete-page'] as const)(
		'rejects stale artifacts when serialization reenters to %s',
		(action) => {
			const model = structuredClone(demoDocument);
			const viewer = mountViewer(document.createElement('div'), { document: model, pageIndex: 1 });
			const fresh = structuredClone(demoDocument);
			fresh.pages[0]!.name = 'New document';
			duringSerialization(() => {
				if (action === 'replace') viewer.controller.setDocument(fresh);
				if (action === 'same') viewer.controller.setDocument(model);
				if (action === 'clear') viewer.controller.setDocument(null);
				if (action === 'destroy') viewer.destroy();
				if (action === 'controller-destroy') viewer.controller.destroy();
				if (action === 'delete-page') model.pages.pop();
			});
			expect(() => viewer.createPrintSnapshot()).toThrow(/document changed|destroyed/);
			if (action === 'replace') expect(viewer.controller.state.document).toBe(fresh);
			viewer.destroy();
		},
	);
	it('captures the original page while allowing reentrant zoom, search, selection and page navigation', () => {
		const viewer = mountViewer(document.createElement('div'), { document: demoDocument });
		const generation = viewer.controller.documentGeneration;
		duringSerialization(() => {
			viewer.controller.setZoom(3);
			viewer.controller.setSearchQuery('framework');
			viewer.controller.setPage(1);
			viewer.controller.selectShape({ id: 'a1', name: 'Your framework', pageId: '2' });
		});
		const result = viewer.createPrintSnapshot();
		expect(result.pageIndices).toEqual([0]);
		expect(viewer.controller.documentGeneration).toBe(generation);
		expect(viewer.controller.state).toMatchObject({
			pageIndex: 1,
			zoom: 3,
			search: { query: 'framework' },
			selectedShape: { id: 'a1' },
		});
		viewer.destroy();
	});
});

describe('DOM-free document generation tracking', () => {
	it('advances only for accepted documents and successful loads, not ordinary state or rejected input', async () => {
		const parser = vi
			.fn()
			.mockRejectedValueOnce(new Error('bad'))
			.mockResolvedValueOnce(demoDocument);
		const controller = new ViewerController(parser);
		expect(controller.documentGeneration).toBe(0);
		controller.setDocument(demoDocument);
		expect(controller.documentGeneration).toBe(1);
		controller.setZoom(2);
		controller.setPage(1);
		controller.setSearchQuery('framework');
		controller.cancelLoad();
		expect(controller.documentGeneration).toBe(1);
		const invalid = structuredClone(demoDocument);
		invalid.pages[0]!.width = -1;
		expect(() => controller.setDocument(invalid)).toThrow();
		await expect(controller.load(new Uint8Array())).rejects.toThrow('bad');
		expect(controller.documentGeneration).toBe(1);
		await controller.load(new Uint8Array());
		expect(controller.documentGeneration).toBe(2);
		controller.setDocument(demoDocument);
		expect(controller.documentGeneration).toBe(3);
		controller.setDocument(null);
		expect(controller.documentGeneration).toBe(4);
		controller.destroy();
		expect(() => controller.documentGeneration).toThrow('destroyed');
	});
});
