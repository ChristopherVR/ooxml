// Property-level merging over live in-memory sessions (memory hub, transport provider).
import { afterEach, expect, it } from 'vitest';
import { createVsdx, editVsdx } from '../index';
import { createRoom, published, settle, until } from './test-support/peers';

const room = createRoom();
afterEach(() => room.destroy());

it('keeps a move and a text and fill change to one shape on both live peers', async () => {
	const seed = await editVsdx(await createVsdx(), [
		{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 1, y: 2, width: 1, height: 1 },
	]);
	const a = room.join('Ada', seed.bytes);
	await published(a);
	const b = room.join('Bob', await createVsdx());
	await until(async () => (await b.side.model()).pages[0]!.shapes.length === 1, 'join');
	await settle(a, b);
	await Promise.all([
		a.side.edit({ type: 'move-shape', pageId: '0', shapeId: '1', x: 5, y: 2 }),
		b.side.edit(
			{ type: 'replace-plain-text', pageId: '0', shapeId: '1', text: 'Hello' },
			{ type: 'format-shape', pageId: '0', shapeId: '1', fillColor: '#00FF00' },
		),
	]);
	await settle(a, b);
	for (const peer of [a, b]) {
		const shape = (await peer.side.model()).pages[0]!.shapes[0]!;
		expect(shape.transform[4]).toBe(4.5);
		expect(shape.text.plainText.trim()).toBe('Hello');
		expect(shape.style.fill.toLowerCase()).toBe('#00ff00');
	}
	expect(a.errors).toEqual([]);
	expect(b.errors).toEqual([]);
});
