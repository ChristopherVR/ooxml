// Masters and media added by two peers at the same time both survive.
import type JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createVsdx, editVsdx } from '../index';
import { type DocPeer, converge, joinFrom, seedPeer } from './test-support/doc-peers';
import { type Added, addMaster } from './test-support/masters';

const mastersOf = async (peer: DocPeer) => {
	const list = (await peer.text('visio/masters/masters.xml')) ?? '';
	return [...list.matchAll(/<Master ID="(\d+)"[^>]*UniqueID="([^"]+)"/g)].map((m) => ({
		id: m[1]!,
		unique: m[2]!,
	}));
};
/** The picture bytes a master's relationships lead to. */
async function pictureOf(peer: DocPeer, unique: string): Promise<number[]> {
	const master = (await mastersOf(peer)).find((m) => m.unique === unique)!;
	const list = (await peer.text('visio/masters/masters.xml'))!;
	const rid = new RegExp(`ID="${master.id}"[^>]*>[\\s\\S]*?r:id="([^"]+)"`).exec(list)![1]!;
	const rels = (await peer.text('visio/masters/_rels/masters.xml.rels'))!;
	const target = new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`).exec(rels)![1]!;
	const part = `visio/masters/${target}`;
	const partRels = (await peer.text(
		`${part.replace(/[^/]+$/, '_rels/')}${part.split('/').pop()}.rels`,
	))!;
	const image = /Target="\.\.\/([^"]+)"/.exec(partRels)![1]!;
	const bytes = await (
		await import('jszip')
	).default
		.loadAsync(peer.bytes)
		.then((z) => z.file(`visio/${image}`)!.async('uint8array'));
	return [...bytes];
}

async function pair() {
	const start = await editVsdx(await createVsdx(), [
		{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 1, y: 2, width: 1, height: 1 },
	]);
	const a = await seedPeer(start.bytes, 'a');
	const b = await joinFrom(a, 'b');
	return { a, b };
}

describe('masters added concurrently', () => {
	const first: Added = {
		id: '2',
		unique: '{A}',
		shapeId: '5',
		media: { path: 'image1.png', bytes: [1, 2, 3], ext: 'png', type: 'image/png' },
	};
	const second: Added = {
		id: '2',
		unique: '{B}',
		shapeId: '6',
		media: { path: 'image1.png', bytes: [4, 5, 6], ext: 'png', type: 'image/png' },
	};

	it('keeps both masters, renumbers the colliding id and follows it in the shape', async () => {
		const { a, b } = await pair();
		await a.patch((zip) => addMaster(zip, first));
		await b.patch((zip) => addMaster(zip, second));
		await converge(a, b);
		for (const peer of [a, b]) {
			const masters = await mastersOf(peer);
			expect(masters.map((m) => m.unique).sort()).toEqual(['{A}', '{B}']);
			expect(new Set(masters.map((m) => m.id)).size).toBe(2);
			const page = (await peer.text('visio/pages/page1.xml'))!;
			const idOf = (u: string) => masters.find((m) => m.unique === u)!.id;
			expect(new RegExp(`<Shape ID="5"[^>]*Master="${idOf('{A}')}"`).test(page)).toBe(true);
			expect(new RegExp(`<Shape ID="6"[^>]*Master="${idOf('{B}')}"`).test(page)).toBe(true);
			const types = (await peer.text('[Content_Types].xml'))!;
			expect(types.match(/application\/vnd\.ms-visio\.master\+xml/g)).toHaveLength(2);
		}
		expect(await mastersOf(a)).toEqual(await mastersOf(b));
	});

	it('keeps both pictures when they share a name and follows the relationship', async () => {
		const { a, b } = await pair();
		await a.patch((zip) => addMaster(zip, first));
		await b.patch((zip) => addMaster(zip, second));
		await converge(a, b);
		for (const peer of [a, b]) {
			expect(await pictureOf(peer, '{A}')).toEqual([1, 2, 3]);
			expect(await pictureOf(peer, '{B}')).toEqual([4, 5, 6]);
			const media = (await peer.names()).filter((n) => n.startsWith('visio/media/'));
			expect(media).toHaveLength(2);
		}
	});

	it('declares the content type of each picture although the types part is shared whole', async () => {
		const { a, b } = await pair();
		await a.patch((zip) => addMaster(zip, first));
		await b.patch((zip) =>
			addMaster(zip, {
				...second,
				media: { path: 'image2.jpeg', bytes: [7, 8], ext: 'jpeg', type: 'image/jpeg' },
			}),
		);
		await converge(a, b);
		for (const peer of [a, b]) {
			const types = (await peer.text('[Content_Types].xml'))!;
			expect(types).toContain('Extension="png"');
			expect(types).toContain('Extension="jpeg"');
			expect((await peer.names()).filter((n) => n.startsWith('visio/media/'))).toHaveLength(2);
		}
	});

	it('stores identical pictures once and keeps the same path', async () => {
		const { a, b } = await pair();
		const same = { ...first.media! };
		await a.patch((zip) => addMaster(zip, { ...first, media: same }));
		await b.patch((zip) => addMaster(zip, { ...second, media: same }));
		await converge(a, b);
		expect(a.shared.media.size).toBe(1);
		expect(a.shared.blobs.size).toBe(1);
		expect((await a.names()).filter((n) => n.startsWith('visio/media/'))).toEqual([
			'visio/media/image1.png',
		]);
		expect(await pictureOf(b, '{B}')).toEqual([1, 2, 3]);
	});

	it('keeps the bytes when a media delete races an add of the same bytes', async () => {
		const { a, b } = await pair();
		const bytes = new Uint8Array([9, 9, 9]);
		const addPng = (path: string) => async (zip: JSZip) => {
			zip.file(path, bytes);
			const types = (await zip.file('[Content_Types].xml')!.async('string')) as string;
			if (!types.includes('Extension="png"'))
				zip.file(
					'[Content_Types].xml',
					types.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>'),
				);
		};
		await a.patch(addPng('visio/media/x.png'));
		await converge(a, b);
		expect(await b.names()).toContain('visio/media/x.png');
		await a.patch((zip) => void zip.remove('visio/media/x.png'));
		await b.patch(addPng('visio/media/y.png'));
		await converge(a, b);
		expect(a.shared.blobs.size).toBe(1);
		for (const peer of [a, b]) {
			const zip = await (await import('jszip')).default.loadAsync(peer.bytes);
			expect([...(await zip.file('visio/media/y.png')!.async('uint8array'))]).toEqual([9, 9, 9]);
		}
	});
});
