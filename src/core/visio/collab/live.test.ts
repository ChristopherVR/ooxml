// Property-level merging over live in-memory sessions (memory hub, transport provider).
import { afterEach, expect, it } from 'vitest';
import { createVsdx, editVsdx } from '../index';
import { createRoom, settle, until, wait } from './test-support/peers';

const room = createRoom();
afterEach(() => room.destroy());

it('keeps a move and a text and fill change to one shape on both live peers', async () => {
	const seed = await editVsdx(await createVsdx(), [
		{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 1, y: 2, width: 1, height: 1 },
	]);
	const a = room.join('Ada', seed.bytes);
	await wait();
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
	await until(async () => {
		const shapes = await Promise.all(
			[a, b].map(async (p) => (await p.side.model()).pages[0]!.shapes[0]!),
		);
		return shapes.every(
			(s) =>
				s.transform[4] === 4.5 &&
				s.text.plainText.trim() === 'Hello' &&
				s.style.fill.toLowerCase() === '#00ff00',
		);
	}, 'both changes on both peers');
	expect(a.errors).toEqual([]);
	expect(b.errors).toEqual([]);
});
