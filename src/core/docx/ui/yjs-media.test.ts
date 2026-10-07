import { afterEach, expect, it } from 'vitest';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollabSession,
} from '../../collab/index';
import { WordYjsMedia } from './yjs-media';

const sessions: CollabSession[] = [];
afterEach(() => {
	for (const session of sessions.splice(0)) session.destroy();
});
function pair() {
	const hub = createMemoryHub();
	return [false, true].map((viewer) => {
		const session = createCollabSession({
			roomId: 'assets',
			provider: transportProvider({ transport: hub.createTransport('assets') }),
			user: { name: viewer ? 'Viewer' : 'Author', ...(viewer ? { role: 'viewer' as const } : {}) },
			heartbeatMs: 0,
			teardown: false,
		});
		sessions.push(session);
		return session;
	});
}
it('shares immutable binary media, isolates callers and rejects viewer publication', () => {
	const [a, b] = pair();
	const source = new WordYjsMedia(a!);
	const reader = new WordYjsMedia(b!);
	const name = source.partName('image/png');
	const bytes = new Uint8Array([137, 80, 78, 71]);
	source.publish(name, { bytes, contentType: 'image/png' });
	bytes[0] = 0;
	expect(reader.get(name)!.bytes).toEqual(new Uint8Array([137, 80, 78, 71]));
	reader.get(name)!.bytes[0] = 0;
	expect(reader.get(name)!.bytes[0]).toBe(137);
	expect(() => reader.publish(name, { bytes, contentType: 'image/png' })).toThrow(/disabled/);
	expect(() => source.publish(name, { bytes, contentType: 'image/png' })).toThrow(/immutable/);
});
it('allocates fresh names after reattachment and validates package part names', () => {
	const [a] = pair();
	const first = new WordYjsMedia(a!);
	const name = first.partName('image/png');
	first.publish(name, { bytes: new Uint8Array([1]), contentType: 'image/png' });
	expect(new WordYjsMedia(a!).partName('image/png')).not.toBe(name);
	expect(() =>
		first.publish('../image.png', { bytes: new Uint8Array([1]), contentType: 'image/png' }),
	).toThrow(/Unsupported/);
	expect(() => first.partName('image/tiff')).toThrow(/Unsupported/);
});
