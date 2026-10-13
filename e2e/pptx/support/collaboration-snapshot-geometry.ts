import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';

import { expect } from '@playwright/test';
import type { Page } from '@playwright/test';
import type { PptxElement } from 'pptx-viewer-core';

const coreRequire = createRequire(createRequire(import.meta.url).resolve('pptx-viewer-core'));
const { PptxHandler } = coreRequire('pptx-viewer-core') as typeof import('pptx-viewer-core');
const { createCollabSession, createWebSocketTransport, transportProvider } = coreRequire(
	'ooxml-core/collab',
) as typeof import('../../../src/core/collab');

type ShapeGeometry = Pick<
	PptxElement,
	'x' | 'y' | 'width' | 'height' | 'xEmu' | 'yEmu' | 'widthEmu' | 'heightEmu'
>;

/** Read the exact synced model through the public protocol, without writing document content. */
export async function syncedShapeGeometry(
	page: Page,
	roomId: string,
	elementId: string,
): Promise<ShapeGeometry> {
	const server = new URL(page.url()).searchParams.get('server');
	expect(server).not.toBeNull();
	const session = createCollabSession({
		roomId,
		user: { name: 'Snapshot verifier', role: 'viewer' },
		provider: transportProvider({
			transport: createWebSocketTransport({ url: `${server}/${encodeURIComponent(roomId)}` }),
		}),
		teardown: false,
		heartbeatMs: 0,
	});
	try {
		await expect.poll(() => session.provider.synced, { timeout: 30_000 }).toBe(true);
		expect(session.canWrite()).toBe(false);
		const slides = session.doc.getArray('pptx:slides').toJSON() as {
			elements: (ShapeGeometry & { id: string })[];
		}[];
		const element = slides[0]?.elements.find((candidate) => candidate.id === elementId);
		expect(element).toBeDefined();
		expect([element!.x, element!.y, element!.width, element!.height].every(Number.isFinite)).toBe(
			true,
		);
		return element!;
	} finally {
		session.destroy();
	}
}

/** Compare saved EMUs with the full-precision model, then return the parser's pixel contract. */
export async function savedShapeGeometry(
	path: string,
	elementId: string,
	model: ShapeGeometry,
): Promise<number[]> {
	const data = await new PptxHandler().load(new Uint8Array(readFileSync(path)));
	const saved = data.slides[0]?.elements.find((candidate) => candidate.id === elementId);
	expect(saved).toBeDefined();
	expect(model.widthEmu).toBeDefined();
	expect(model.heightEmu).toBeDefined();
	// This scenario moves both axes and leaves the size untouched. An unchanged
	// size retains its original EMUs; a moved axis is quantized once by the writer.
	expect(saved!.xEmu).toBe(Math.round(model.x * 9525));
	expect(saved!.yEmu).toBe(Math.round(model.y * 9525));
	expect(saved!.widthEmu).toBe(model.widthEmu);
	expect(saved!.heightEmu).toBe(model.heightEmu);
	return [saved!.x, saved!.y, saved!.width, saved!.height];
}
